import collections
import copy
import json
import os
import re
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
        self.target_filename = trans_filename or f"{self.src_filename.rsplit('.', 1)[0]}-译文.docx"
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
        children = p.xpath('./*', namespaces=self.namespaces)
        text_list = []
        one_text_list = []
        rd_list = []
        for child in children:
            child_tag = etree.QName(child).localname
            if child.xpath(".//w:p", namespaces=self.namespaces):
                continue
            # 遇到oMath标签，要拆句
            if child_tag == "oMath":
                text_list.append("".join(one_text_list))
                one_text_list = []
            else:
                r_text_list = child.xpath('.//w:t//text()', namespaces=self.namespaces)
                one_text_list.extend(r_text_list)
                if child_tag == "r" and child.xpath(".//w:t//text()", namespaces=self.namespaces):
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
            if text.strip() == "":
                all_sentence_list.append(text)
            else:
                sentence_list = self.split_sentence(text)
                all_sentence_list.extend(sentence_list)
        # 将块归类到每个句子下边
        if any(all_sentence_list):
            p_info = self.get_sentence_r_list(all_sentence_list, rd_list)
        else:
            p_info = []
        return p_info

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
            # print(f"--{sentence}--{len(sentence)}", start_index)
            # print(rd_list[start_index:])
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

    def parse_sub_file(self, filename):
        """
        解析word/document.xml
        :param filename: eg: word/document.xml
        :return:
        """
        with zipfile.ZipFile(self.src_filename, "r") as z:
            if filename not in z.namelist():
                return []
            tree = etree.parse(z.open(filename, "r"))
            self.namespaces = tree.getroot().nsmap
            self.namespaces_dict = {self.namespaces[k]: f"{k}:" for k in self.namespaces}
            p_list = tree.xpath(".//w:p", namespaces=self.namespaces)
            p_infos = []
            for p in p_list:
                p_info = self.parse_p(p)
                if p_info:
                    p_infos.append(p_info)
            # print(json.dumps(p_infos, ensure_ascii=False))
        self.file_infos_dict[filename] = p_infos
        return p_infos

    def parse_file(self):
        filename_list = ["word/document.xml", "word/footnotes.xml", "word/endnotes.xml", "word/comments.xml"]
        pattern1 = re.compile("word/header\d*\.xml$")
        pattern2 = re.compile("word/footer\d*\.xml$")
        with zipfile.ZipFile(self.src_filename, "r") as z:
            for filename in z.namelist():
                if filename in filename_list or pattern1.findall(filename) or pattern2.findall(filename):
                    self.parse_sub_file(filename)
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
                    # trans_text = trans_dict.get(origin_text, f"【{origin_text}】")
                    trans_text = f"【{origin_text}】"
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

    def join_xml(self, p_node, p_infos, p_info_index, p_namespaces_dict, tree):
        trans_p_node = copy.deepcopy(p_node)
        child_node_list = trans_p_node.xpath("./*", namespaces=self.namespaces)

        separator_node = []
        has_content = False
        for child_node in child_node_list:
            if child_node.xpath(".//w:p", namespaces=self.namespaces):
                continue
            child_tag = etree.QName(child_node).localname
            if child_tag == "oMath":
                continue
            elif child_tag == "r" and child_node.xpath(".//w:t//text()", namespaces=self.namespaces):
                r_parent_node = child_node.getparent()
                r_parent_node.remove(child_node)
                has_content = True
            else:
                r_list = child_node.xpath('.//w:r', namespaces=self.namespaces)
                for r in r_list:
                    r_parent_node = r.getparent()
                    if r.xpath(".//w:t//text()", namespaces=self.namespaces):
                        r_parent_node.remove(r)
                        has_content = True
        if not has_content:
            return p_info_index

        # 将句子的r放到一个列表中
        p_info = p_infos[p_info_index]
        sentence_rs = []
        for sentence in p_info:
            sentence_rs.extend(sentence["trans_r"])

        # 组装新的p
        for sentence_r_dict in sentence_rs:
            sentence_r_dict["w:r"].update(p_namespaces_dict)
            r_xml_str = xml2dict.unparse(sentence_r_dict)
            r_xml_str = r_xml_str.replace("""<?xml version="1.0" encoding="utf-8"?>""", '').strip()
            trans_p_node.append(etree.fromstring(r_xml_str))

        # 将p放入文档
        parent_node = p_node.getparent()
        p_node_index = parent_node.index(p_node)

        # parent_node.insert(p_node_index+1, trans_p_node)
        parent_node[p_node_index] = trans_p_node

        p_info_index += 1
        return p_info_index

    def get_xml_str(self, filename):
        with zipfile.ZipFile(self.src_filename, "r") as z:
            p_infos = self.file_infos_dict[filename]
            tree = etree.parse(z.open(filename, "r"))
            self.namespaces = tree.getroot().nsmap
            p_namespaces_dict = {f"@xmlns:{k}": v for k, v in self.namespaces.items()}
            p_node_list = tree.xpath(".//w:p", namespaces=self.namespaces)
            p_info_index = 0
            for p_node in p_node_list:
                p_info_index = self.join_xml(p_node, p_infos, p_info_index, p_namespaces_dict, tree)

        # if filename == "word/document.xml":
        #     print(etree.tostring(tree, encoding="utf-8", pretty_print=True).decode())

        return etree.tostring(tree, encoding="utf-8", pretty_print=True).decode()

    def compose_docx(self):
        """
        合成文件
        :return:
        """
        # self.file_infos_dict.pop("word/comments.xml", None)
        with zipfile.ZipFile(self.src_filename, "r") as z:
            with zipfile.ZipFile(self.target_filename, "w") as new_z:
                for item in z.infolist():
                    if item.filename not in self.file_infos_dict:
                        new_z.writestr(item, z.read(item.filename))
                for filename in self.file_infos_dict:
                    xml_str = self.get_xml_str(filename)
                    new_z.writestr(filename, xml_str)


if __name__ == '__main__':
    filename = "file/1.docx"
    docx_parser = DocxParser(filename, download_type=0)
    # docx_parser.parse_language()
    file_infos_dict = docx_parser.parse_file()
    # print(json.dumps(file_infos_dict["word/document.xml"], ensure_ascii=False))
    # json.dump(file_infos_dict["word/document.xml"], open("test.json", "w", encoding="utf-8"), ensure_ascii=False, indent=4)
    # print('-------------------')
    trans_file_infos_dict = docx_parser.translate_file()
    # json.dump(trans_file_infos_dict["word/document.xml"], open("test1.json", "w", encoding="utf-8"), ensure_ascii=False,
    #           indent=4)
    # print(json.dumps(trans_file_infos_dict["word/document.xml"], ensure_ascii=False))
    # print('-------------------')
    # print(json.dumps(docx_parser.translate_file(), ensure_ascii=False))
    docx_parser.compose_docx()
    # a = docx_parser.get_document_xml_str()
    # print(a)
    # # docx_parser.json2xml()
    # a = docx_parser.split_sentence("你好啊，我是谁。你睡吗？")
    # print(a)
