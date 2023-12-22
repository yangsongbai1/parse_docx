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

    @staticmethod
    def split_sentence(text):
        seg = pysbd.Segmenter(language="en", clean=False)
        return seg.segment(text)

    def parse_p(self, p):
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
        p_info = self.get_sentence_r_list(sentence_list, p_dict)
        return p_info

    @staticmethod
    def get_sentence_r_list(sentence_list, p_dict):
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
                    # print(values, '------------------', type(values))
                    continue
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
                    p_info = self.parse_p(child)
                    p_infos.append(p_info)
                elif tag == "tbl":
                    p_list = child.xpath(".//w:p", namespaces=self.namespaces)
                    for p in p_list:
                        p_info = self.parse_p(p)
                        p_infos.append(p_info)
                else:
                    # print(tag, "其他标签")
                    # todo: 解析其他标签
                    pass
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
                    footnote_infos = self.parse_p(p)
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
                    endnote_info = self.parse_p(p)
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
                    comment_info = self.parse_p(p)
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
                if filename.startswith("header"):
                    tree = etree.parse(z.open(full_filename, "r"))
                    root = tree.getroot()
                    self.namespaces = root.nsmap
                    children = root.xpath("./w:p", namespaces=self.namespaces)
                    headers_infos = []
                    for child in children:
                        header_info = self.parse_p(child)
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
                if filename.startswith("footer"):
                    tree = etree.parse(z.open(full_filename, "r"))
                    root = tree.getroot()
                    self.namespaces = root.nsmap
                    children = root.xpath("./w:p", namespaces=self.namespaces)
                    footer_infos = []
                    for child in children:
                        footer_info = self.parse_p(child)
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
                    # 获取段落
                    p_info = p_infos[p_index]
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
                        p_info = p_infos[p_index]
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
    docx_parser = DocxParser("file/1.docx")
    file_infos_dict = docx_parser.parse_file()
    # print(json.dumps(file_infos_dict, ensure_ascii=False))
    # print('-------------------')
    trans_file_infos_dict = docx_parser.translate_file()
    # print(json.dumps(trans_file_infos_dict, ensure_ascii=False))
    # print('-------------------')
    # print(json.dumps(docx_parser.translate_file(), ensure_ascii=False))
    docx_parser.compose_docx()
    # # docx_parser.json2xml()
