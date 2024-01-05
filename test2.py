from lxml import etree

a = etree.parse("test.xml")

dd = a.xpath("preceding-sibling::a[1]")
print(dd)
for i in dd:
    x = etree.tostring(i, encoding="utf-8", pretty_print=True, xml_declaration=False).decode()
    print(x)
    print('----------------')