import copy

from lxml import etree

a = etree.parse("test.xml")

pp = a.xpath(".//w:p", namespaces=a.getroot().nsmap)
for p in pp:
    if p.xpath(".//w:p", namespaces=a.getroot().nsmap):
        print(p.xpath(".//text()", namespaces=a.getroot().nsmap))
        p_parent = p.getparent()
        b_index = p_parent.index(p)
        copy_b = copy.deepcopy(p)
        copy_b.text = "我是哈哈"
        p_parent[b_index] = copy_b

# print(etree.tostring(a, encoding="utf-8").decode("utf-8"))
