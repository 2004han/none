from lxml import etree as E
from pathlib import Path
r=E.parse(str(Path(r'E:\试切实验\测量系统公式转换\document.xml')))
n={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
for i,p in enumerate(r.findall('.//w:p',n)):
 if i in [4,5,37,38,39,40,47,48,49,50,51]:print(i,E.tostring(p,encoding='unicode')[:2800], '\nPARENT',p.getparent().tag)
print('SECTION',E.tostring(r.find('.//w:sectPr',n),encoding='unicode'))
