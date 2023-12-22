import collections
import copy
import json
import os
import zipfile
from copy import deepcopy
from pprint import pprint
import shutil
import xml2dict
from lxml import etree
import pysbd


class DocxParser:

    def __init__(self, src_filename, trans_filename=None):
        self.namespaces_dict = {}
        self.namespaces = {}
        self.p_infos = []
        self.endnote_infos = []
        self.footnotes_infos = []
        self.comments_infos = []
        self.file_infos_dict = {}
        self.src_filename = src_filename
        self.trans_filename = trans_filename or f"{self.src_filename.rsplit('.', 1)[0]}-译文.docx"
        self.trans_list = [self.p_infos, self.endnote_infos]

    @staticmethod
    def split_sentence(text):
        seg = pysbd.Segmenter(language="en", clean=False)
        return seg.segment(text)

    def parse_p2(self, p):
        """
        解析p标签
        :param p:
        :return:
        """
        # 获取段落内容
        r_text_list = p.xpath('./w:r/w:t/text()', namespaces=self.namespaces)
        p_text = "".join(r_text_list)
        # 拆句，拆成句子列表
        sentence_list = self.split_sentence(p_text)
        # 转成json
        p_xml_str = etree.tostring(p, encoding="utf-8").decode("utf-8")
        p_dict = xml2dict.parse(p_xml_str)
        # 将块归类到每个句子下边
        p_info = self.get_sentence_r_list2(sentence_list, p_dict)
        return p_info

    @staticmethod
    def get_sentence_r_list2(sentence_list, p_dict):
        """
        将句子块，归类的每个句子下边
        :param sentence_list:
        :param r_list:
        :return:
        """
        r_dict = p_dict["w:p"]
        sentence_list_all = []
        if len(sentence_list) == 1:
            sentence_dict = {
                "origin_text": sentence_list[0],
                "rs": r_dict
            }
            sentence_list_all.append(sentence_dict)
            return sentence_list_all

        if len(sentence_list) == 0:
            sentence_dict = {
                "origin_text": "",
                "rs": r_dict
            }
            sentence_list_all.append(sentence_dict)
            return sentence_list_all

        run_on_start = 0
        for sentence in sentence_list:
            sentence_dict = {
                "origin_text": sentence
            }
            # 句子长度
            sentence_len = len(sentence)
            sentence_r_dict = {}
            count = 0
            rr_dict = copy.deepcopy(r_dict)
            # print(rr_dict, '-------')
            for r, values in rr_dict.items():
                """
                如果小于句子长度，就加入句子块
                如果等于句子长度，就加入句子块，并跳出循环
                如果句子长度大于句子长度，就进行切割
                """
                if not isinstance(values, dict):
                    # todo: 这里有问题，需要处理
                    print(values, '------------------', type(values))
                    continue
                print(values, '================')
                wt = values.get("w:t", "")
                if isinstance(wt, dict):
                    r_text = wt.get("#text", "")
                else:
                    r_text = wt
                r_text = "" if r_text is None else r_text
                count += len(r_text)
                if count < sentence_len:
                    sentence_r_dict[r] = rr_dict[r]
                    r_dict.pop(r, None)
                elif count == sentence_len:
                    sentence_r_dict[r] = rr_dict[r]
                    r_dict.pop(r, None)
                    break
                else:
                    run_on_end = run_on_start + sentence_len
                    # print(run_on_start, run_on_end)
                    sentence_text = r_text[run_on_start:run_on_end]
                    r_copy = copy.deepcopy(values)
                    if isinstance(wt, dict):
                        r_copy["w:t"]["#text"] = sentence_text
                    else:
                        r_copy["w:t"] = sentence_text
                    sentence_r_dict[r] = r_copy
                    run_on_start = run_on_end
                    break

            sentence_dict["rs"] = sentence_r_dict
            sentence_list_all.append(sentence_dict)

        return sentence_list_all

    def parse_docx(self):
        with zipfile.ZipFile(self.src_filename, "r") as z:
            tree = etree.parse(z.open('word/document.xml', "r"))
            self.namespaces = tree.getroot().nsmap
            self.namespaces_dict = {self.namespaces[k]: f"{k}:" for k in self.namespaces}
            children = tree.xpath("./w:body/*", namespaces=self.namespaces)
            p_infos = []
            for child in children:
                tag = etree.QName(child.tag).localname
                if tag == "p":
                    p_info = self.parse_p2(child)
                    p_infos.append(p_info)
                elif tag == "tbl":
                    p_list = child.xpath(".//w:p", namespaces=self.namespaces)
                    for p in p_list:
                        p_info = self.parse_p2(p)
                        p_infos.append(p_info)
                else:
                    # print(tag, "其他标签")
                    # todo: 解析其他标签
                    pass
            # print(json.dumps(p_infos, ensure_ascii=False))
            self.p_infos = p_infos
            return p_infos

    def translate2(self):
        """
        翻译
        :param text:
        :return:
        """
        for p_info in self.p_infos:
            for sentence in p_info:
                origin_text = sentence.get("origin_text")
                trans_text = f"【{origin_text}】"
                sentence['trans_text'] = trans_text
                rs = sentence.get("rs")
                trans_rs = {}
                write_wt = True
                for t, v in rs.items():
                    if not isinstance(v, dict):
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
                sentence["trans_r"] = trans_rs

    def translate_endnotes(self):
        """
        翻译
        :param text:
        :return:
        """
        for endnote_info in self.endnote_infos:
            for sentence in endnote_info:
                origin_text = sentence.get("origin_text")
                trans_text = f"【{origin_text}】"
                sentence['trans_text'] = trans_text
                rs = sentence.get("rs")
                trans_rs = {}
                write_wt = True
                for t, v in rs.items():
                    if not isinstance(v, dict):
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
                sentence["trans_r"] = trans_rs

    def translate_footnotes(self):
        """
        翻译
        :param text:
        :return:
        """
        for footnote_info in self.footnotes_infos:
            for sentence in footnote_info:
                origin_text = sentence.get("origin_text")
                trans_text = f"【{origin_text}】"
                sentence['trans_text'] = trans_text
                rs = sentence.get("rs")
                trans_rs = {}
                write_wt = True
                for t, v in rs.items():
                    if not isinstance(v, dict):
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
                sentence["trans_r"] = trans_rs

    def translate_comments(self):
        """
        翻译
        :param text:
        :return:
        """
        for footnote_info in self.comments_infos:
            for sentence in footnote_info:
                origin_text = sentence.get("origin_text")
                trans_text = f"【{origin_text}】"
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
                sentence["trans_r"] = trans_rs
        # print(json.dumps(self.comments_infos, ensure_ascii=False))
        # print("-========-----------------")

    def translate_file(self):
        """
        翻译
        :param text:
        :return:
        """
        for headers_filename in self.file_infos_dict:
            headers_infos = self.file_infos_dict[headers_filename]
            for header_info in headers_infos:
                for sentence in header_info:
                    origin_text = sentence.get("origin_text")
                    trans_text = f"【{origin_text}】"
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
                    sentence["trans_r"] = trans_rs
        # print(json.dumps(self.file_infos_dict, ensure_ascii=False))
        # print("-========-----------------")
        return self.file_infos_dict

    def parse_end_notes(self):
        """
        解析尾注
        :return:
        """
        try:
            with zipfile.ZipFile(self.src_filename, "r") as z:
                tree = etree.parse(z.open('word/endnotes.xml', "r"))
                root = tree.getroot()
                self.namespaces = root.nsmap
                children = root.xpath("./w:endnote", namespaces=self.namespaces)
                endnote_infos = []
                for child in children:
                    p_list = child.xpath("./w:p", namespaces=self.namespaces)
                    for p in p_list:
                        endnote_info = self.parse_p2(p)
                        endnote_infos.append(endnote_info)
            self.endnote_infos = endnote_infos
            return endnote_infos
        except Exception as e:
            print(e)
            return []

    def parse_foot_notes(self):
        """
        解析脚注
        :return:
        """
        try:
            with zipfile.ZipFile(self.src_filename, "r") as z:
                tree = etree.parse(z.open('word/footnotes.xml', "r"))
                root = tree.getroot()
                self.namespaces = root.nsmap
                children = root.xpath("./w:footnote", namespaces=self.namespaces)
                footnotes_infos = []
                for child in children:
                    p_list = child.xpath("./w:p", namespaces=self.namespaces)
                    for p in p_list:
                        footnote_infos = self.parse_p2(p)
                        footnotes_infos.append(footnote_infos)
            self.footnotes_infos = footnotes_infos
            return footnotes_infos
        except Exception as e:
            print(e)
            return []

    def parse_comments(self):
        """
        解析批注
        :return:
        """
        try:
            with zipfile.ZipFile(self.src_filename, "r") as z:
                tree = etree.parse(z.open('word/comments.xml', "r"))
                root = tree.getroot()
                self.namespaces = root.nsmap
                children = root.xpath("./w:comment", namespaces=self.namespaces)
                comments_infos = []
                for child in children:
                    p_list = child.xpath("./w:p", namespaces=self.namespaces)
                    for p in p_list:
                        comment_info = self.parse_p2(p)
                        comments_infos.append(comment_info)
            self.comments_infos = comments_infos
            return comments_infos
        except Exception as e:
            print(e)
            return []

    def parse_headers(self):
        """
        解析页眉
        :return:
        """
        with zipfile.ZipFile(self.src_filename, "r") as z:
            for i in z.infolist():
                full_filename = i.filename
                filename = full_filename.split("/")[-1]
                if filename.startswith("header"):
                    tree = etree.parse(z.open(full_filename, "r"))
                    root = tree.getroot()
                    self.namespaces = root.nsmap
                    children = root.xpath("./w:p", namespaces=self.namespaces)
                    headers_infos = []
                    for child in children:
                        header_info = self.parse_p2(child)
                        headers_infos.append(header_info)
                    self.file_infos_dict[full_filename] = headers_infos
        return self.file_infos_dict

    def parse_footers(self):
        """
        解析页脚
        :return:
        """
        with zipfile.ZipFile(self.src_filename, "r") as z:
            for i in z.infolist():
                full_filename = i.filename
                filename = full_filename.split("/")[-1]
                if filename.startswith("footer"):
                    tree = etree.parse(z.open(full_filename, "r"))
                    root = tree.getroot()
                    self.namespaces = root.nsmap
                    children = root.xpath("./w:p", namespaces=self.namespaces)
                    headers_infos = []
                    for child in children:
                        header_info = self.parse_p2(child)
                        headers_infos.append(header_info)
                    self.file_infos_dict[full_filename] = headers_infos
        return self.file_infos_dict

    def get_comments_xml_str(self):
        """
        获取批注xml字符串
        :return:
        """
        try:
            with zipfile.ZipFile(self.src_filename, "r") as z:
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
                        p_info = self.comments_infos[p_index]
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
        except Exception as e:
            print(e)
            return ""

    def get_footnotes_xml_str(self):
        """
        获取脚注xml字符串
        :return:
        """
        try:
            with zipfile.ZipFile(self.src_filename, "r") as z:
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
                        p_info = self.footnotes_infos[p_index]
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
        except Exception as e:
            print(e)
            return ""

    def get_endnotes_xml_str(self):
        try:
            with zipfile.ZipFile(self.src_filename, "r") as z:
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
                        p_info = self.endnote_infos[p_index]
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
        except Exception as e:
            print(e)
            return ""

    def get_document_xml_str(self):
        with zipfile.ZipFile(self.src_filename, "r") as z:
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
                    # 获取段落
                    p_info = self.p_infos[p_index]
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
                    body[child_index] = p_node
                    p_index += 1
                elif tag == "tbl":
                    p_list = child.xpath(".//w:p", namespaces=self.namespaces)
                    for p in p_list:
                        parent_node = p.getparent()
                        index = parent_node.index(p)
                        p_info = self.p_infos[p_index]
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
                else:
                    # print(tag, "其他标签")
                    # print(json.dumps(p_infos, ensure_ascii=False))
                    pass

        return etree.tostring(tree, encoding="utf-8", pretty_print=True).decode()

    def get_file_xml_str(self, file_name):
        """
        获取页眉页脚xml字符串
        :return:
        """
        with zipfile.ZipFile(self.src_filename, "r") as z:
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

    def compose_docx(self):
        """
        1. 复制文件
        2. 替换文件内容
        :return:
        """
        document_xml_str = self.get_document_xml_str()
        endnotes_xml_str = self.get_endnotes_xml_str()
        footnotes_xml_str = self.get_footnotes_xml_str()
        comments_xml_str = self.get_comments_xml_str()
        with zipfile.ZipFile(self.src_filename, "r") as z:
            with zipfile.ZipFile(self.trans_filename, "w") as new_z:
                for item in z.infolist():
                    if all([
                        item.filename not in [
                            "word/document.xml",
                            "word/endnotes.xml",
                            "word/footnotes.xml",
                            "word/comments.xml"
                        ],
                        item.filename not in self.file_infos_dict
                    ]):
                        new_z.writestr(item, z.read(item.filename))
                new_z.writestr("word/document.xml", document_xml_str)
                # new_z.writestr("word/endnotes.xml", endnotes_xml_str)
                # new_z.writestr("word/footnotes.xml", footnotes_xml_str)
                # new_z.writestr("word/comments.xml", comments_xml_str)
                # for file_name in self.file_infos_dict:
                #     file_xml_str = self.get_file_xml_str(file_name)
                #     new_z.writestr(file_name, file_xml_str)


if __name__ == '__main__':
    docx_parser = DocxParser("file/日语.docx")
    # docx_parser.parse_footers()
    # docx_parser.parse_headers()
    docx_parser.parse_docx()
    docx_parser.parse_end_notes()
    docx_parser.parse_foot_notes()
    docx_parser.parse_comments()
    docx_parser.translate_comments()
    # # a = docx_parser.get_comments_xml_str()
    # # print(a)
    docx_parser.translate_endnotes()
    docx_parser.translate_footnotes()
    docx_parser.translate2()
    docx_parser.translate_file()
    # print(json.dumps(docx_parser.translate_file(), ensure_ascii=False))
    docx_parser.compose_docx()
    # # docx_parser.json2xml()
