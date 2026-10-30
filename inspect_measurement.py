import zipfile
from pathlib import Path
import xml.etree.ElementTree as E
src=Path(r'E:\腾讯电脑管家软件搬家\xwechat_files\wxid_ihkuwle3qrse22_615a\temp\RWTemp\2026-10\2705c282f41ed16ead29d6bebe8aa335\8 测量系统要求.docx')
out=Path(r'E:\试切实验\测量系统公式转换');out.mkdir(exist_ok=True)
z=zipfile.ZipFile(src)
ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main','v':'urn:schemas-microsoft-com:vml','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
root=E.fromstring(z.read('word/document.xml'))
rels=E.fromstring(z.read('word/_rels/document.xml.rels'))
print('RELATIONSHIPS',[(x.get('Id'),x.get('Target')) for x in rels if 'image' in x.get('Type','') or 'oleObject' in x.get('Type','')])
for i,p in enumerate(root.findall('.//w:p',ns)):
 text=''.join(x.text or '' for x in p.findall('.//w:t',ns))
 imgs=[x.attrib for x in p.findall('.//a:blip',ns)+p.findall('.//v:imagedata',ns)]
 if text or imgs: print(i,repr(text),'IMGS',imgs)
for name in z.namelist():
 if name.startswith('word/media/'):
  dest=out/Path(name).name;dest.write_bytes(z.read(name));print('MEDIA',str(dest),len(z.read(name)))
(out/'document.xml').write_bytes(z.read('word/document.xml'))
