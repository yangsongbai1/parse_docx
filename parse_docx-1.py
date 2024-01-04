import collections
import copy
import json
import os
import uuid
import zipfile
from copy import deepcopy
from pprint import pprint
import shutil

import langid
from langdetect import detect

import xml2dict
from lxml import etree
import pysbd
import pymongo
from trans_dict import trans_dict

myclient = pymongo.MongoClient("mongodb://localhost:27017/")
mydb = myclient["docx"]

corpus_dict = {
    "实验室": "laboratory"
}


def insert_dict(my_dict, key_to_insert_after, new_dict):
    keys = list(my_dict.keys())
    my_dict_list = list(my_dict.items())
    index = keys.index(key_to_insert_after)
    result_dict = dict(my_dict_list[:index + 1] + list(new_dict.items()) + my_dict_list[index + 1:])
    return result_dict


class DocxParser:

    def __init__(self, src_filename, trans_filename=None, download_type=0):
        self.namespaces_dict = {}
        self.namespaces = {}
        self.file_infos_dict = {}
        self.src_filename = src_filename
        self.trans_filename = trans_filename or f"{self.src_filename.rsplit('.', 1)[0]}-译文.docx"
        # 0 译文， 1-原译文对照，2-译原文对照
        self.download_type = download_type

    @staticmethod
    def split_sentence(text):
        seg = pysbd.Segmenter(language="en", clean=False)
        return seg.segment(text)

    def parse_language(self):
        with zipfile.ZipFile(self.src_filename, "r") as z:
            tree = etree.parse(z.open('word/document.xml', "r"))
            text_list = tree.xpath(".//text()")
            text = "".join(text_list)
            # text = "我是繁體，真的是正體字"
            language_list = ['en', 'zh']
            langid.set_languages(language_list)
            language = langid.classify(text)[0]
            # if language == "zh":
            #     language = detect(text)
            print(language)

    def parse_p(self, p):
        """
        解析p标签
        :param p:
        :return:
        """
        # 获取段落内容
        # 遇到公式，需要将公式左右切割成两句话
        children = p.xpath('./*', namespaces=self.namespaces)
        text_list = []
        one_text_list = []
        textbox_p_infos = []
        rd_list = []
        for child in children:
            tag = etree.QName(child).localname
            if tag == "oMath":
                text_list.append("".join(one_text_list))
                one_text_list = []
            elif child.xpath("mc:AlternateContent", namespaces=self.namespaces):
                textbox_p_list = child.xpath(".//w:p", namespaces=self.namespaces)
                for textbox_p in textbox_p_list:
                    p_info, _ = self.parse_p(textbox_p)
                    textbox_p_infos.append(p_info)
            else:
                # 获取段落内容，如果有问题，就将 .//w:t 改成 ./w:t 并特殊处理其他标签
                r_text_list = child.xpath('.//w:t/text()', namespaces=self.namespaces)
                one_text_list.extend(r_text_list)
                # print(r_text_list, '-----text-----')
                # 获取r
                if tag == "r" and child.xpath(".//w:t", namespaces=self.namespaces):
                    r_dict = xml2dict.parse(etree.tostring(child, encoding="utf-8").decode("utf-8"),
                                            strip_whitespace=False)
                    rd_list.append(r_dict)
                else:
                    r_list = child.xpath('.//w:r', namespaces=self.namespaces)
                    for r in r_list:
                        if r.xpath(".//w:t", namespaces=self.namespaces):
                            r_dict = xml2dict.parse(etree.tostring(r, encoding="utf-8").decode("utf-8"),
                                                    strip_whitespace=False)
                            rd_list.append(r_dict)

        if one_text_list:
            text_list.append("".join(one_text_list))
        # 拆句，拆成句子列表
        all_sentence_list = []
        for text in text_list:
            sentence_list = self.split_sentence(text)
            all_sentence_list.extend(sentence_list)
        # 将块归类到每个句子下边
        p_info = self.get_sentence_r_list(all_sentence_list, rd_list)
        return p_info, textbox_p_infos

    @staticmethod
    def get_sentence_r_list(sentence_list, rd_list):
        """
        将句子块，归类的每个句子下边
        :param sentence_list:
        :param rd_list:
        :return:
        """
        sentence_list_all = []
        if len(sentence_list) == 1:
            sentence_dict = {
                "origin_text": sentence_list[0],
                "rs": rd_list
            }
            sentence_list_all.append(sentence_dict)
            return sentence_list_all

        run_on_start = 0
        start_index = 0
        for index, sentence in enumerate(sentence_list):
            sentence_dict = {
                "origin_text": sentence
            }
            # 句子长度
            sentence_len = len(sentence)
            sentence_rd_list = []
            count = 0
            print(f"--{sentence}--{len(sentence)}", start_index)
            print(rd_list[start_index:])
            for rd in rd_list[start_index:]:
                """
                如果小于句子长度，就加入句子块
                如果等于句子长度，就加入句子块，并跳出循环
                如果句子长度大于句子长度，就进行切割
                """
                wt = rd["w:r"]["w:t"]
                if isinstance(wt, dict):
                    r_text = wt.get("#text", "")
                else:
                    r_text = wt
                r_text = "" if r_text is None else r_text
                count += len(r_text)
                if count < sentence_len:
                    sentence_rd_list.append(rd)
                    start_index += 1
                elif count == sentence_len:
                    sentence_rd_list.append(rd)
                    start_index += 1
                    break
                else:
                    run_on_end = run_on_start + sentence_len
                    sentence_text = r_text[run_on_start:run_on_end]
                    rd_copy = copy.deepcopy(rd)
                    if isinstance(wt, dict):
                        rd_copy["w:r"]["w:t"]["#text"] = sentence_text
                    else:
                        rd_copy["w:r"]["w:t"] = sentence_text
                    sentence_rd_list.append(rd_copy)
                    run_on_start = run_on_end
                    break

            sentence_dict["rs"] = sentence_rd_list
            sentence_list_all.append(sentence_dict)

        return sentence_list_all

    def parse_document(self):
        with zipfile.ZipFile(self.src_filename, "r") as z:
            tree = etree.parse(z.open('word/document.xml', "r"))
            self.namespaces = tree.getroot().nsmap
            self.namespaces_dict = {self.namespaces[k]: f"{k}:" for k in self.namespaces}
            children = tree.xpath("./w:body/*", namespaces=self.namespaces)
            p_infos = []
            for child in children:
                tag = etree.QName(child.tag).localname
                if tag == "p":
                    p_info, textbox_p_infos = self.parse_p(child)
                    # 文本框节点前置，适用于前置处理文本框,然后组装p节点
                    # p_infos.extend(textbox_p_infos)
                    p_infos.append(p_info)
                    p_infos.extend(textbox_p_infos)
                elif tag == "tbl":
                    p_list = child.xpath(".//w:p", namespaces=self.namespaces)
                    for p in p_list:
                        p_info, _ = self.parse_p(p)
                        p_infos.append(p_info)
                else:
                    # print(tag, "其他标签")
                    # todo: 解析其他标签
                    pass
                # 文本框节点后置，适用于后置处理文本框
                # p_infos.extend(textbox_p_infos)
            # print(json.dumps(p_infos, ensure_ascii=False))
        self.file_infos_dict["word/document.xml"] = p_infos
        return p_infos

    def parse_footnotes(self):
        """
        解析脚注
        :return:
        """
        with zipfile.ZipFile(self.src_filename, "r") as z:
            if "word/footnotes.xml" not in z.namelist():
                return []
            tree = etree.parse(z.open('word/footnotes.xml', "r"))
            root = tree.getroot()
            self.namespaces = root.nsmap
            children = root.xpath("./w:footnote", namespaces=self.namespaces)
            footnotes_infos = []
            for child in children:
                p_list = child.xpath("./w:p", namespaces=self.namespaces)
                for p in p_list:
                    footnote_infos, _ = self.parse_p(p)
                    footnotes_infos.append(footnote_infos)
        self.file_infos_dict["word/footnotes.xml"] = footnotes_infos
        return footnotes_infos

    def parse_endnotes(self):
        """
        解析尾注
        :return:
        """
        with zipfile.ZipFile(self.src_filename, "r") as z:
            if "word/endnotes.xml" not in z.namelist():
                return []
            tree = etree.parse(z.open('word/endnotes.xml', "r"))
            root = tree.getroot()
            self.namespaces = root.nsmap
            children = root.xpath("./w:endnote", namespaces=self.namespaces)
            endnote_infos = []
            for child in children:
                p_list = child.xpath("./w:p", namespaces=self.namespaces)
                for p in p_list:
                    endnote_info, _ = self.parse_p(p)
                    endnote_infos.append(endnote_info)
        self.file_infos_dict["word/endnotes.xml"] = endnote_infos
        return endnote_infos

    def parse_comments(self):
        """
        解析批注
        :return:
        """
        with zipfile.ZipFile(self.src_filename, "r") as z:
            if "word/comments.xml" not in z.namelist():
                return []
            tree = etree.parse(z.open('word/comments.xml', "r"))
            root = tree.getroot()
            self.namespaces = root.nsmap
            children = root.xpath("./w:comment", namespaces=self.namespaces)
            comments_infos = []
            for child in children:
                p_list = child.xpath("./w:p", namespaces=self.namespaces)
                for p in p_list:
                    comment_info, _ = self.parse_p(p)
                    comments_infos.append(comment_info)
        self.file_infos_dict["word/comments.xml"] = comments_infos
        return comments_infos

    def parse_headers(self):
        """
        解析页眉
        :return:
        """
        with zipfile.ZipFile(self.src_filename, "r") as z:
            for full_filename in z.namelist():
                filename = full_filename.split("/")[-1]
                if filename.startswith("header") and filename.endswith(".xml"):
                    tree = etree.parse(z.open(full_filename, "r"))
                    root = tree.getroot()
                    self.namespaces = root.nsmap
                    children = root.xpath("./w:p", namespaces=self.namespaces)
                    headers_infos = []
                    for child in children:
                        header_info, _ = self.parse_p(child)
                        headers_infos.append(header_info)
                    self.file_infos_dict[full_filename] = headers_infos
        return self.file_infos_dict

    def parse_footers(self):
        """
        解析页脚
        :return:
        """
        with zipfile.ZipFile(self.src_filename, "r") as z:
            for full_filename in z.namelist():
                filename = full_filename.split("/")[-1]
                if filename.startswith("footer") and filename.endswith(".xml"):
                    tree = etree.parse(z.open(full_filename, "r"))
                    root = tree.getroot()
                    self.namespaces = root.nsmap
                    children = root.xpath("./w:p", namespaces=self.namespaces)
                    footer_infos = []
                    for child in children:
                        footer_info, _ = self.parse_p(child)
                        footer_infos.append(footer_info)
                    self.file_infos_dict[full_filename] = footer_infos
        return self.file_infos_dict

    def parse_file(self):
        self.parse_document()
        self.parse_footnotes()
        self.parse_endnotes()
        self.parse_comments()
        self.parse_headers()
        self.parse_footers()
        return self.file_infos_dict

    def translate_file(self):
        """
        翻译
        :param text:
        :return:
        """
        for filename in self.file_infos_dict:
            file_infos = self.file_infos_dict[filename]
            for file_info in file_infos:
                for sentence in file_info:
                    origin_text = sentence.get("origin_text")
                    trans_text = trans_dict.get(origin_text, f"【{origin_text}】")
                    # trans_text = f"【{origin_text}】"
                    sentence['trans_text'] = trans_text
                    rs = sentence.get("rs")
                    trans_rs = []
                    r = copy.deepcopy(rs[0])
                    v = r["w:r"]
                    wt = v["w:t"]
                    if isinstance(wt, dict):
                        wt["#text"] = trans_text
                    else:
                        v["w:t"] = trans_text
                    trans_rs.append(r)
                    # self.align_corner_mark(origin_rs=rs, trans_rs=trans_rs)
                    sentence["trans_r"] = trans_rs
        return self.file_infos_dict

    def get_r_style_dict(self, rs):
        """
        翻译
        先把段落样式写进去，构造一个空壳
        将原文每个块的样式组成字典, 如果有重复的块内容，则将块内容携带上前一个块的内容
        eg: {块内容: 本块样式}
        循环原文快，匹配语料库，如果匹配到语料库，用语料译文在翻译文中查找，
        如果查找到，则将匹配到的译文 赋予样式，存入块列表中
        循环完毕，对比匹配到的块列表和翻译文，将未匹配的内容，按照翻译文顺序，插入到块列表中，并继承上一个块的样式
        :param text:
        :return:
        """
        r_style_dict = {}
        for t, v in rs.items():
            if isinstance(v, dict) and (wt := v.get("w:t")):
                r_text = wt
                if isinstance(wt, dict):
                    r_text = wt.get("#text")
                if not r_text:
                    continue
                r_style_dict[r_text] = {t: v}
        return r_style_dict

    def translate_file2(self):
        """
        翻译
        :param text:
        :return:
        """
        for filename in self.file_infos_dict:
            file_infos = self.file_infos_dict[filename]
            for file_info in file_infos:
                for sentence in file_info:
                    origin_text = sentence.get("origin_text")
                    trans_text = trans_dict.get(origin_text, origin_text)
                    sentence['trans_text'] = trans_text
                    rs = sentence.get("rs")
                    r_style_dict = self.get_r_style_dict(rs)
                    trans_rs = {}
                    # 匹配到样式的块列表
                    style_text = []
                    for t, v in rs.items():
                        if not isinstance(v, dict):
                            trans_rs[t] = v
                            continue
                        tag = t.split(".")[0]
                        wt = v.get("w:t")
                        if tag != "w:r" or wt is None:
                            trans_rs[t] = v
                        else:
                            r_text = wt
                            if isinstance(wt, dict):
                                r_text = wt.get("#text")
                            if not r_text:
                                trans_rs[t] = v
                            else:
                                # 去匹配语料
                                r_trans_text = corpus_dict.get(r_text)
                                if r_trans_text:
                                    # 拿语料译文，去翻译文中找
                                    index = trans_text.find(r_trans_text)
                                    if index != -1:
                                        # 如果存在，则查找对应的样式
                                        r_style = r_style_dict.get(r_text)
                                        # 创造译文样式块
                                        # r_style = {'w:r': {'w:rPr': {'w:u': {'@w:val': 'single'}}, 'w:t': '实验室'}}
                                        trans_r_style = copy.deepcopy(r_style)
                                        trans_r_style["w:r"]["w:t"] = r_trans_text
                                        style_text.append(trans_r_style)
                    # print(style_text)
                    # print('-----')
                    # print(json.dumps(trans_rs, ensure_ascii=False))
                    # print('===++++===')
                    base_r = {"w:rPr": {"w:rFonts": {"@w:hint": "eastAsia"},
                                        "w:lang": {"@w:val": "en-US", "@w:eastAsia": "zh-CN"}}, "w:t": "我是第二行的第二句话。"}
                    trans_text_start = 0
                    # style_text = [{'w:r': {'w:rPr': {'w:u': {'@w:val': 'single'}}, 'w:t': '实验室'}}, ...]
                    for i in style_text:
                        t_text = i["w:r"]["w:t"]
                        # print(t_text, '=-=-')
                        trans_text_end = trans_text.index(t_text)
                        # print(trans_text_start, trans_text_end)
                        front_no_style_text = trans_text[trans_text_start:trans_text_end]
                        # print(trans_text)
                        # print(front_no_style_text, '--')
                        if front_no_style_text:
                            # print(trans_rs)
                            front_no_style = copy.deepcopy(base_r)
                            front_no_style["w:t"] = front_no_style_text
                            # print(front_no_style)
                            trans_rs[f"w:r.{str(uuid.uuid4())}"] = front_no_style
                            # print(trans_rs)
                        trans_rs[f"w:r.{str(uuid.uuid4())}"] = list(i.values())[0]
                        trans_text_start = trans_text_end + len(t_text)
                    front_no_style_text = trans_text[trans_text_start:]
                    if front_no_style_text:
                        front_no_style = copy.deepcopy(base_r)
                        front_no_style["w:t"] = front_no_style_text
                        trans_rs[f"w:r.{str(uuid.uuid4())}"] = front_no_style
                    sentence["trans_r"] = trans_rs
                    # print(json.dumps(trans_rs, ensure_ascii=False))
                    # print('------我是最终-')
        # print(json.dumps(self.file_infos_dict, ensure_ascii=False))
        # print("-========-----------------")
        return self.file_infos_dict

    @staticmethod
    def align_corner_mark(origin_rs, trans_rs):
        """
        对齐角标
        :param origin_rs: 原文样式
        :param trans_rs: 译文样式
        :return:
        """
        pre_text = ""
        style_text_dict = {}
        # 循环原文样式，找到含有上下标的样式
        # 含有角标是，将角标和前面的文本组合起来，尽量减少重复，作为key，样式作为value
        for tag, v in origin_rs.items():
            tag = tag.split(".")[0]
            if tag == "w:r":
                if "w:vertAlign" in str(v):
                    # 有角标的内容
                    mark_text = v.get("w:t")
                    zuhe_text = pre_text + mark_text
                    d = {
                        "pre_text": pre_text,
                        "mark_text": mark_text,
                        "style": v
                    }
                    style_text_dict[zuhe_text] = d
                    pre_text = ""
                else:
                    pre_text = v.get("w:t", "")
        # print(style_text_dict)
        # 循环有角标的样式，找到对应的译文
        for zuhe_text, d in style_text_dict.items():
            # 译文的全部样式
            trans_rs_copy = copy.deepcopy(trans_rs)
            for tag, t_rs in trans_rs_copy.items():
                # tag = w:r.1
                # t_rs 某一个块或r
                # {
                #     "w:rPr": {
                #         "w:u": {
                #             "@w:val": "single"
                #         }
                #     },
                #     "w:t": "译文"
                # }
                if not isinstance(t_rs, dict):
                    continue
                r_wt_text = t_rs.get("w:t", "")
                if (index := r_wt_text.find(zuhe_text)) != -1:
                    # 第二句的样式，继承第一句的样式
                    text2_rs = copy.deepcopy(t_rs)
                    # 要添加的样式，即中间的样式
                    middle_rs = d["style"]
                    pre_text = d["pre_text"]
                    middle_text = d["mark_text"]
                    # left_text, right_text = r_wt_text.split(zuhe_text)
                    left_text = r_wt_text[:index]
                    right_text = r_wt_text[index + len(zuhe_text):]
                    left_text += pre_text
                    t_rs["w:t"] = left_text
                    text2_rs["w:t"] = right_text
                    middle_rs["w:t"] = middle_text
                    # 给原有的翻译样式添加角标
                    trans_rs[tag] = t_rs
                    # trans_rs = insert_dict(trans_rs, tag, tag)
                    trans_rs[f"w:r.{str(uuid.uuid4())}"] = middle_rs
                    trans_rs[f"w:r.{str(uuid.uuid4())}"] = text2_rs
                    # print(trans_rs)

    def translate_file3(self):
        """
        翻译 角标还原
        :param text:
        :return:
        """
        for filename in self.file_infos_dict:
            file_infos = self.file_infos_dict[filename]
            for file_info in file_infos:
                for sentence in file_info:
                    origin_text = sentence.get("origin_text")
                    trans_text = trans_dict.get(origin_text, origin_text)
                    sentence['trans_text'] = trans_text
                    rs = sentence.get("rs")
                    trans_rs = {}
                    write_wt = True
                    for t, v in rs.items():
                        if not isinstance(v, dict):
                            trans_rs[t] = v
                            continue
                        tag = t.split(".")[0]
                        wt = v.get("w:t")
                        if tag != "w:r" or wt is None:
                            trans_rs[t] = v
                        elif write_wt:
                            r_copy = copy.deepcopy(v)
                            r_copy["w:t"] = trans_text
                            trans_rs[t] = r_copy
                            write_wt = False
                        else:
                            continue

                    # 角标还原
                    pre_text = ""
                    style_text_dict = {}
                    for t, v in rs.items():
                        tag = t.split(".")[0]
                        if tag == "w:r":
                            if "w:vertAlign" in str(v):
                                text = v.get("w:t")
                                zuhe_text = pre_text + text
                                d = {
                                    "pre_text": pre_text,
                                    "text": text,
                                    "style": v
                                }
                                style_text_dict[zuhe_text] = d
                                pre_text = ""
                            else:
                                pre_text = v.get("w:t", "")
                    # print(style_text_dict)
                    for zuhe_text, d in style_text_dict.items():
                        # 译文的全部样式
                        trans_rs_copy = copy.deepcopy(trans_rs)
                        for node, t_rs in trans_rs_copy.items():
                            # t_rs 某一个块或r
                            # {
                            #     "w:rPr": {
                            #         "w:u": {
                            #             "@w:val": "single"
                            #         }
                            #     },
                            #     "w:t": "译文"
                            # }
                            if not isinstance(t_rs, dict):
                                continue
                            tst = t_rs.get("w:t", "")
                            if (index := tst.find(zuhe_text)) != -1:
                                # print(node, tst, index)
                                # 第二句的样式，继承第一句的样式
                                text2_rs = copy.deepcopy(t_rs)
                                # 要添加的样式，即中间的样式
                                middle_rs = d["style"]
                                pre_text = d["pre_text"]
                                middle_text = d["text"]
                                text1, text2 = trans_text.split(zuhe_text)
                                text1 += pre_text
                                t_rs["w:t"] = text1
                                text2_rs["w:t"] = text2
                                middle_rs["w:t"] = middle_text

                                trans_rs["w:r"] = t_rs
                                trans_rs["w:r.1"] = middle_rs
                                trans_rs["w:r.2"] = text2_rs
                                # print(t_rs)
                                # print(text2_rs)
                                # print(middle_rs)
                                print(trans_rs)
                    sentence["trans_r"] = trans_rs

        # print(json.dumps(self.file_infos_dict, ensure_ascii=False))
        # print("-========-----------------")
        return self.file_infos_dict

    @staticmethod
    def create_node(p_info, p_namespaces_dict, key):
        p_dict = {}
        p_dict.update(p_namespaces_dict)
        rk_no = collections.defaultdict(int)
        for sentence in p_info:
            sentence_r_dict = sentence[key]
            for rk, rv in sentence_r_dict.items():
                if rk in p_dict:
                    rk = rk.split(".")[0]
                    no = rk_no[rk]
                    p_dict[f"{rk}.{no}"] = rv
                    rk_no[rk] += 1
                else:
                    p_dict[rk] = rv
        p_xml_dict = {
            "w:p": p_dict
        }
        p_xml_str = xml2dict.unparse(p_xml_dict)
        p_xml_str = p_xml_str.replace("""<?xml version="1.0" encoding="utf-8"?>""", '').strip()
        p_node = etree.fromstring(p_xml_str)
        return p_node

    def join_child_xml(self, p_node, p_index, p_infos, p_namespaces_dict, key):
        start_p_index = p_index
        AlternateContent_list = p_node.xpath(".//mc:AlternateContent", namespaces=self.namespaces)
        for AlternateContent in AlternateContent_list:
            AlternateContent_parent_node = AlternateContent.getparent()
            AlternateContent_index = AlternateContent_parent_node.index(AlternateContent)
            p_list = AlternateContent.xpath(".//w:p", namespaces=self.namespaces)
            for p in p_list:
                parent_node = p.getparent()
                index = parent_node.index(p)
                textbox_p_info = p_infos[p_index]
                textbox_p_node = self.create_node(textbox_p_info, p_namespaces_dict, key)
                parent_node[index] = textbox_p_node
                p_index += 1
            AlternateContent_parent_node[AlternateContent_index] = AlternateContent
        return p_index - start_p_index

    def join_xml(self, child, p_infos, p_index, p_namespaces_dict):
        parent_node = child.getparent()
        index = parent_node.index(child)
        p_info = p_infos[p_index]
        origin_p_node = self.create_node(p_info, p_namespaces_dict, "rs")
        trans_p_node = self.create_node(p_info, p_namespaces_dict, "trans_r")
        p_index += 1
        step_index = self.join_child_xml(trans_p_node, p_index, p_infos, p_namespaces_dict, "trans_r")
        p_index += step_index
        # 0 译文， 1-原译文对照，2-译原文对照
        if self.download_type == 0:
            parent_node[index] = trans_p_node
        elif self.download_type == 1:
            parent_node[index] = trans_p_node
            parent_node.insert(index, origin_p_node)
        elif self.download_type == 2:
            parent_node[index] = origin_p_node
            parent_node.insert(index, trans_p_node)
        return p_index

    def get_document_xml_str(self):
        with zipfile.ZipFile(self.src_filename, "r") as z:
            p_infos = self.file_infos_dict["word/document.xml"]
            tree = etree.parse(z.open('word/document.xml', "r"))
            root = tree.getroot()
            namespaces = root.nsmap
            p_namespaces_dict = {f"@xmlns:{k}": v for k, v in namespaces.items()}
            body = root.xpath("./w:body", namespaces=namespaces)[0]
            children = body.xpath("./*", namespaces=self.namespaces)
            p_index = 0
            for child_index, child in enumerate(children):
                tag = etree.QName(child.tag).localname
                if tag == "p":
                    p_index = self.join_xml(child, p_infos, p_index, p_namespaces_dict)
                elif tag == "tbl":
                    p_list = child.xpath(".//w:p", namespaces=self.namespaces)
                    for p in p_list:
                        p_index = self.join_xml(p, p_infos, p_index, p_namespaces_dict)
                else:
                    # print(tag, "其他标签")
                    # print(json.dumps(p_infos, ensure_ascii=False))
                    pass

        return etree.tostring(tree, encoding="utf-8", pretty_print=True).decode()

    def get_footnotes_xml_str(self):
        """
        获取脚注xml字符串
        :return:
        """
        with zipfile.ZipFile(self.src_filename, "r") as z:
            if "word/footnotes.xml" not in z.namelist():
                return ""
            footnotes_infos = self.file_infos_dict["word/footnotes.xml"]
            tree = etree.parse(z.open('word/footnotes.xml', "r"))
            root = tree.getroot()
            self.namespaces = root.nsmap
            p_namespaces_dict = {f"@xmlns:{k}": v for k, v in self.namespaces.items()}
            children = root.xpath("./w:footnote", namespaces=self.namespaces)
            p_index = 0
            for child in children:
                p_list = child.xpath("./w:p", namespaces=self.namespaces)
                for p in p_list:
                    parent_node = p.getparent()
                    index = parent_node.index(p)
                    p_info = footnotes_infos[p_index]
                    p_dict = {}
                    p_dict.update(p_namespaces_dict)
                    rk_no = collections.defaultdict(int)
                    for sentence in p_info:
                        sentence_r_dict = sentence["trans_r"]
                        for rk, rv in sentence_r_dict.items():
                            if rk in p_dict:
                                rk = rk.split(".")[0]
                                no = rk_no[rk]
                                p_dict[f"{rk}.{no}"] = rv
                                rk_no[rk] += 1
                            else:
                                p_dict[rk] = rv

                    p_xml_dict = {
                        "w:p": p_dict
                    }
                    p_xml_str = xml2dict.unparse(p_xml_dict)
                    p_xml_str = p_xml_str.replace("""<?xml version="1.0" encoding="utf-8"?>""", '').strip()
                    p_node = etree.fromstring(p_xml_str)
                    parent_node[index] = p_node
                    p_index += 1
        return etree.tostring(tree, encoding="utf-8", pretty_print=True).decode()

    def get_endnotes_xml_str(self):
        with zipfile.ZipFile(self.src_filename, "r") as z:
            if "word/endnotes.xml" not in z.namelist():
                return ""
            endnote_infos = self.file_infos_dict["word/endnotes.xml"]
            end_notes_tree = etree.parse(z.open('word/endnotes.xml', "r"))
            end_notes_root = end_notes_tree.getroot()
            self.namespaces = end_notes_root.nsmap
            p_namespaces_dict = {f"@xmlns:{k}": v for k, v in self.namespaces.items()}
            children = end_notes_root.xpath("./w:endnote", namespaces=self.namespaces)
            p_index = 0
            for child in children:
                p_list = child.xpath("./w:p", namespaces=self.namespaces)
                for p in p_list:
                    parent_node = p.getparent()
                    index = parent_node.index(p)
                    p_info = endnote_infos[p_index]
                    p_dict = {}
                    p_dict.update(p_namespaces_dict)
                    rk_no = collections.defaultdict(int)
                    for sentence in p_info:
                        sentence_r_dict = sentence["trans_r"]
                        for rk, rv in sentence_r_dict.items():
                            if rk in p_dict:
                                rk = rk.split(".")[0]
                                no = rk_no[rk]
                                p_dict[f"{rk}.{no}"] = rv
                                rk_no[rk] += 1
                            else:
                                p_dict[rk] = rv

                    p_xml_dict = {
                        "w:p": p_dict
                    }
                    p_xml_str = xml2dict.unparse(p_xml_dict)
                    p_xml_str = p_xml_str.replace("""<?xml version="1.0" encoding="utf-8"?>""", '').strip()
                    p_node = etree.fromstring(p_xml_str)
                    parent_node[index] = p_node
                    p_index += 1
        return etree.tostring(end_notes_tree, encoding="utf-8", pretty_print=True).decode()

    def get_comments_xml_str(self):
        """
        获取批注xml字符串
        :return:
        """
        with zipfile.ZipFile(self.src_filename, "r") as z:
            if "word/comments.xml" not in z.namelist():
                return ""
            comments_infos = self.file_infos_dict["word/comments.xml"]
            tree = etree.parse(z.open('word/comments.xml', "r"))
            root = tree.getroot()
            self.namespaces = root.nsmap
            p_namespaces_dict = {f"@xmlns:{k}": v for k, v in self.namespaces.items()}
            children = root.xpath("./w:comment", namespaces=self.namespaces)
            p_index = 0
            for child in children:
                p_list = child.xpath("./w:p", namespaces=self.namespaces)
                for p in p_list:
                    parent_node = p.getparent()
                    index = parent_node.index(p)
                    p_info = comments_infos[p_index]
                    p_dict = {}
                    p_dict.update(p_namespaces_dict)
                    rk_no = collections.defaultdict(int)
                    for sentence in p_info:
                        sentence_r_dict = sentence["trans_r"]
                        for rk, rv in sentence_r_dict.items():
                            if rk in p_dict:
                                rk = rk.split(".")[0]
                                no = rk_no[rk]
                                p_dict[f"{rk}.{no}"] = rv
                                rk_no[rk] += 1
                            else:
                                p_dict[rk] = rv

                    p_xml_dict = {
                        "w:p": p_dict
                    }
                    p_xml_str = xml2dict.unparse(p_xml_dict)
                    p_xml_str = p_xml_str.replace("""<?xml version="1.0" encoding="utf-8"?>""", '').strip()
                    p_node = etree.fromstring(p_xml_str)
                    parent_node[index] = p_node
                    p_index += 1
        return etree.tostring(tree, encoding="utf-8", pretty_print=True).decode()

    def get_header_footer_xml_str(self, file_name):
        """
        获取页眉页脚xml字符串
        :return:
        """
        with zipfile.ZipFile(self.src_filename, "r") as z:
            if file_name not in z.namelist():
                return ""
            tree = etree.parse(z.open(file_name, "r"))
            root = tree.getroot()
            self.namespaces = root.nsmap
            p_namespaces_dict = {f"@xmlns:{k}": v for k, v in self.namespaces.items()}
            p_list = root.xpath("./w:p", namespaces=self.namespaces)
            p_index = 0
            for p in p_list:
                parent_node = p.getparent()
                index = parent_node.index(p)
                p_info = self.file_infos_dict[file_name][p_index]
                p_dict = {}
                p_dict.update(p_namespaces_dict)
                rk_no = collections.defaultdict(int)
                for sentence in p_info:
                    sentence_r_dict = sentence["trans_r"]
                    for rk, rv in sentence_r_dict.items():
                        if rk in p_dict:
                            rk = rk.split(".")[0]
                            no = rk_no[rk]
                            p_dict[f"{rk}.{no}"] = rv
                            rk_no[rk] += 1
                        else:
                            p_dict[rk] = rv

                p_xml_dict = {
                    "w:p": p_dict
                }
                p_xml_str = xml2dict.unparse(p_xml_dict)
                p_xml_str = p_xml_str.replace("""<?xml version="1.0" encoding="utf-8"?>""", '').strip()
                p_node = etree.fromstring(p_xml_str)
                parent_node[index] = p_node
                p_index += 1
        return etree.tostring(tree, encoding="utf-8", pretty_print=True).decode()

    def get_xml_str(self, full_filename):
        filename = full_filename.split("/")[-1]
        pure_filename = filename.split(".")[0]
        func_str = f"get_{pure_filename}_xml_str"
        if hasattr(self, func_str):
            func = getattr(self, func_str)
            return func()
        else:
            return self.get_header_footer_xml_str(full_filename)

    def compose_docx(self):
        """
        1. 复制文件
        2. 替换文件内容
        :return:
        """
        with zipfile.ZipFile(self.src_filename, "r") as z:
            with zipfile.ZipFile(self.trans_filename, "w") as new_z:
                for item in z.infolist():
                    if item.filename not in self.file_infos_dict:
                        new_z.writestr(item, z.read(item.filename))
                for filename in self.file_infos_dict:
                    xml_str = self.get_xml_str(filename)
                    new_z.writestr(filename, xml_str)


if __name__ == '__main__':
    filename = "file/1.docx"
    docx_parser = DocxParser(filename, download_type=0)
    docx_parser.parse_language()
    file_infos_dict = docx_parser.parse_file()
    # print(json.dumps(file_infos_dict["word/document.xml"], ensure_ascii=False))
    # json.dump(file_infos_dict["word/document.xml"], open("test.json", "w", encoding="utf-8"), ensure_ascii=False, indent=4)
    # print('-------------------')
    trans_file_infos_dict = docx_parser.translate_file()
    # json.dump(trans_file_infos_dict["word/document.xml"], open("test1.json", "w", encoding="utf-8"), ensure_ascii=False,
    #           indent=4)
    print(json.dumps(trans_file_infos_dict["word/document.xml"], ensure_ascii=False))
    # print('-------------------')
    # print(json.dumps(docx_parser.translate_file(), ensure_ascii=False))
    # docx_parser.compose_docx()
    # # docx_parser.json2xml()
    # a = docx_parser.split_sentence("你好啊，我是谁。你睡吗？")
    # print(a)
