import copy

from lxml import etree

a = etree.parse("test.xml")

pp = a.xpath(".//p", namespaces=a.getroot().nsmap)
ss = True
for p in pp:
    parent = p.getparent()
    index = parent.index(p)
    # parent[index] = None
    text = p.text
    if text == "3" and ss:
        parent.remove(p)
        ss = False

print(etree.tostring(a, pretty_print=True, encoding="utf-8").decode("utf-8"))