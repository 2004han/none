from pathlib import Path
from zipfile import ZipFile
from lxml import etree as E
import json, struct, pythoncom, win32com.client, pymupdf

BASE=Path(__file__).resolve().parent
OUT=Path('E:/试切实验/8 测量系统要求_MathType可编辑版.docx')
SRC=Path(r'E:/腾讯电脑管家软件搬家/xwechat_files/wxid_ihkuwle3qrse22_615a/temp/RWTemp/2026-10/2705c282f41ed16ead29d6bebe8aa335/8 测量系统要求.docx')
NS={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main','o':'urn:schemas-microsoft-com:office:office'}
with ZipFile(SRC) as z: original=E.fromstring(z.read('word/document.xml'))
with ZipFile(OUT) as z:
    assert z.testzip() is None
    doc=E.fromstring(z.read('word/document.xml'))
    objects=doc.findall('.//o:OLEObject',NS)
    assert len(objects)==25 and all(x.get('ProgID')=='Equation.DSMT4' for x in objects)
    assert not doc.findall('.//w:drawing',NS)
    alltext=''.join(doc.xpath('//w:t/text()',namespaces=NS))
    oldtext=original.xpath('//w:t/text()',namespaces=NS)
    offset=0
    for t in oldtext:
        idx=alltext.find(t,offset)
        assert idx>=0,('Original text missing or reordered',t)
        offset=idx+len(t)
    for n in range(1,11): assert alltext.count(f'({n})')==1
    assert alltext.index('10.3')<alltext.index('(2)')<alltext.index('(3)')<alltext.index('10.4')<alltext.index('(4)')
    assert '距离的样本标准差为：' in alltext
    assert '不能使用无限延长线掩盖 P35 的超越。' in alltext

# Check the MTEF record stream using the documented format, independent of previews.
def parse_records(b,pos):
    records=[]
    while True:
        typ=b[pos];pos+=1
        if typ==0: return pos,records
        if typ==1:
            opts=b[pos];pos+=1
            assert opts in (0,1)
            children=[]
            if opts==0: pos,children=parse_records(b,pos)
            records.append(('line',children))
        elif typ==2:
            opts,face=b[pos:pos+2];pos+=2
            assert opts==0
            c=chr(struct.unpack_from('<H',b,pos)[0]);pos+=2
            records.append(('char',c,face-128))
        elif typ==3:
            opts,selector,var,extra=b[pos:pos+4];pos+=4
            assert opts==0 and extra==0
            pos,children=parse_records(b,pos)
            lines=[c for c in children if c[0]=='line']
            expected={1:1,5:1,10:2,11:2,13:1,16:3,23:3,27:2,28:2}
            assert len(lines)==expected[selector],(selector,len(lines))
            records.append(('template',selector,var,children))
        elif 10<=typ<=14: records.append(('size',typ))
        else: raise AssertionError(('Unknown record',typ,pos))

native_audit=[]
for path in sorted(BASE.glob('mathtype*.bin')):
    st=pythoncom.StgOpenStorage(str(path),None,0x10)
    assert str(st.Stat()[8])=='{0002CE03-0000-0000-C000-000000000046}'
    stream=st.OpenStream('Equation Native',None,0x10);raw=stream.Read(stream.Stat()[2])
    assert struct.unpack_from('<I',raw,0)[0]==28
    assert struct.unpack_from('<I',raw,8)[0]==len(raw)-28
    data=raw[28:];start=data.index(bytes([10,1,0]))
    end,tree=parse_records(data,start)
    assert end==len(data)
    native_audit.append({'file':path.name,'records':tree})
    stream=None;st=None
(BASE/'native_audit.json').write_text(json.dumps(native_audit,ensure_ascii=False,indent=2),encoding='utf-8')

a=win32com.client.Dispatch('KWPS.Application')
d=a.Documents.Open(str(OUT),ConfirmConversions=False,ReadOnly=True,AddToRecentFiles=False,Visible=False)
try:
    recognized=[]
    for i in range(1,d.InlineShapes.Count+1):
        sh=d.InlineShapes.Item(i)
        if sh.Type==1: recognized.append(sh.OLEFormat.ProgID)
    assert recognized.count('Equation.DSMT4')==25
    d.ExportAsFixedFormat(str(BASE/'preview.pdf'),17)
finally: d.Close(0)
pdf=pymupdf.open(BASE/'preview.pdf')
for i,p in enumerate(pdf):
    p.get_pixmap(matrix=pymupdf.Matrix(1.3,1.3)).save(BASE/f'page{i+1}.png')
print(json.dumps({'all_original_text_retained':True,'mathtype_objects':len(objects),'native_streams_checked':len(native_audit),'recognized_by_wps':len(recognized),'pages':len(pdf)},ensure_ascii=False))
