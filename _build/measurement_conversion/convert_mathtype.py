from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from copy import deepcopy
import struct, json, io, re
import pythoncom
from lxml import etree as E
import matplotlib
matplotlib.use('Agg')
from matplotlib.mathtext import math_to_image
from matplotlib.font_manager import FontProperties
from PIL import Image

BASE = Path(__file__).resolve().parent
SRC = Path(r'E:/腾讯电脑管家软件搬家/xwechat_files/wxid_ihkuwle3qrse22_615a/temp/RWTemp/2026-10/2705c282f41ed16ead29d6bebe8aa335/8 测量系统要求.docx')
OUT = Path('E:/试切实验/8 测量系统要求_MathType可编辑版.docx')
NS = {'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      'v':'urn:schemas-microsoft-com:vml','o':'urn:schemas-microsoft-com:office:office',
      'a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
def tag(x):
    pre,name=x.split(':'); return '{'+NS[pre]+'}'+name
def el(x, **attrs):
    e=E.Element(tag(x))
    for k,v in attrs.items(): e.set(tag(k) if ':' in k else k,str(v))
    return e

# A single expression tree supplies both the native editable data and its preview.
def T(s, roman=False): return ('text',s,roman)
def seq(*items): return ('seq',*[T(x) if isinstance(x,str) else x for x in items])
def sub(a,b): return ('sub',T(a) if isinstance(a,str) else a,T(b) if isinstance(b,str) else b)
def sup(a,b): return ('sup',a,T(b) if isinstance(b,str) else b)
def frac(a,b): return ('frac',T(a) if isinstance(a,str) else a,T(b) if isinstance(b,str) else b)
def bar(a): return ('bar',T(a) if isinstance(a,str) else a)
def par(a): return ('par',a)
def norm(a): return sub(('norm',a),'2')
def summ(a,lo,hi): return ('sum',a,T(lo) if isinstance(lo,str) else lo,T(hi) if isinstance(hi,str) else hi)
def maximum(lo): return ('lim',T('max',True),T(lo) if isinstance(lo,str) else lo)

fs=sub('f','s'); vmax=sub('v',T('max',True)); ds=sub(seq('Δ','s'),T('max',True))
qi=sub('q','i'); rij=sub('r','ij'); gi=sub('g','i'); lij=sub('l','ij'); li=sub(bar('l'),'i')
sli=sub('S',sub('l','i')); ejk=sub('e','jk'); ajk=sub('a','jk'); sjk=sub('S','jk')
ej=sub('E','j'); emax=sub('E',T('max',True)); nj=sub('N','j')
EQ = [seq(fs,'≥',frac(vmax,ds)),
      seq(gi,'=',frac('1','n'),summ(rij,'j=1','n')),
      seq(sub('AT','p'),'=',maximum('1≤i≤m'),norm(seq(gi,'−',qi))),
      seq(lij,'=',norm(seq(rij,'−',gi))),
      seq(li,'=',frac('1','n'),summ(lij,'j=1','n')),
      seq(sli,'=',('sqrt',frac(summ(sup(par(seq(lij,'−',li)),'2'),'j=1','n'),seq('n','−','1')))),
      seq(sub('RT','p'),'=',maximum('1≤i≤m'),par(seq(li,'+','3',sli))),
      seq(ejk,'=',T('dist',True),par(seq(ajk,',',sjk))),
      seq(ej,'=',maximum(seq('1≤k≤',nj)),ejk),
      seq(emax,'=',maximum('1≤j≤n'),ej)]

def latex(x):
    k=x[0]
    if k=='text':
        v=x[1]
        for a,b in [('Δ',r'\Delta '),('≥',r'\geq '),('≤',r'\leq '),('−','-'),('…',r'\ldots '),('·',r'\cdot ')]: v=v.replace(a,b)
        return r'\mathrm{'+v+'}' if x[2] else v
    if k=='seq': return ''.join(latex(a) for a in x[1:])
    if k in ('sub','sup'): return '{'+latex(x[1])+'}'+('_' if k=='sub' else '^')+'{'+latex(x[2])+'}'
    if k=='frac': return r'\frac{'+latex(x[1])+'}{'+latex(x[2])+'}'
    if k=='bar': return r'\overline{'+latex(x[1])+'}'
    if k=='sqrt': return r'\sqrt{'+latex(x[1])+'}'
    if k=='par': return r'\left('+latex(x[1])+r'\right)'
    if k=='norm': return r'\left\Vert '+latex(x[1])+r'\right\Vert'
    if k=='sum': return r'\sum_{'+latex(x[2])+'}^{'+latex(x[3])+'}'+latex(x[1])
    if k=='lim': return r'\underset{'+latex(x[2])+'}{'+latex(x[1])+'}'
    raise ValueError(k)

def char(c,face=None):
    if face is None: face=8 if c.isdigit() else (3 if c.isalpha() and ord(c)<128 else 6)
    return bytes([2,0,128+face])+struct.pack('<H',ord(c))
def line(b=None): return b'\x01\x01' if b is None else b'\x01\x00'+b+b'\x00'
def tmpl(selector,content,variation=0): return bytes([3,0,selector,variation,0])+content+b'\x00'
def size(level): return bytes([min(10+level,12)])
def native(x,level=0):
    k=x[0]
    if k=='text': return b''.join(char(c,2 if x[2] else None) for c in x[1])
    if k=='seq': return b''.join(native(a,level) for a in x[1:])
    if k in ('sub','sup'):
        body=line(native(x[2],level+1))
        slots=body+line() if k=='sub' else line()+body
        return native(x[1],level)+tmpl(27 if k=='sub' else 28,size(level+1)+slots)+size(level)
    if k=='frac': return tmpl(11,line(native(x[1],level))+size(level)+line(native(x[2],level)))+size(level)
    if k=='bar': return tmpl(13,line(native(x[1],level)))+size(level)
    if k=='sqrt': return tmpl(10,line(native(x[1],level))+size(level+1)+line())+size(level)
    if k in ('par','norm'):
        fences=(char('(')+char(')')) if k=='par' else char(chr(0xEC09),22)+char(chr(0xEC0A),22)
        return tmpl(1 if k=='par' else 5,line(native(x[1],level))+fences,3)+size(level)
    if k=='sum':
        return tmpl(16,line(native(x[1],level))+size(level+1)+line(native(x[2],level+1))+line(native(x[3],level+1))+bytes([13])+char('∑'),0x43)+size(level)
    if k=='lim':
        return tmpl(23,line(native(x[1],level))+size(level+1)+line(native(x[2],level+1))+line(),1)+size(level)
    raise ValueError(k)

sample=(BASE/'spec/Equation Native.dat').read_bytes()
prefix=sample[28:28+184]
assert prefix[-1]==0 and prefix[:5]==bytes([5,1,0,5,0])
prefix=prefix.replace(b'Arial\x00',b'Times New Roman\x00')
prefix=prefix.replace(bytes.fromhex('12 00 08 21 0f'),bytes.fromhex('12 00 08 21 2f'))

def make_ole(expr,path):
    data=prefix+b'\x0a'+line(native(expr))+b'\x00'
    header=sample[:8]+struct.pack('<I',len(data))+bytes(16)
    st=pythoncom.StgCreateDocfile(str(path),0x1012,0)
    pythoncom.WriteClassStg(st,pythoncom.MakeIID('{0002CE03-0000-0000-C000-000000000046}'))
    for name,raw in [('Equation Native',header+data),('\x01CompObj',(BASE/'spec/_CompObj.dat').read_bytes()),('\x01Ole',(BASE/'spec/_Ole.dat').read_bytes()),('\x03ObjInfo',(BASE/'spec/_ObjInfo.dat').read_bytes())]:
        stream=st.CreateStream(name,0x1012,0,0);stream.Write(raw);stream.Commit(0);stream=None
    st.Commit(0);st=None
    return data

with ZipFile(SRC) as z: files={i.filename:z.read(i.filename) for i in z.infolist()}
root=E.fromstring(files['word/document.xml'])
rels=E.fromstring(files['word/_rels/document.xml.rels'])
ct=E.fromstring(files['[Content_Types].xml'])
ps=root.findall('.//w:p',NS)
manifest=[]
matplotlib.rcParams['mathtext.fontset']='stix'
matplotlib.rcParams['font.family']='STIXGeneral'

def relationship(rid,typ,target):
    e=E.SubElement(rels,'{http://schemas.openxmlformats.org/package/2006/relationships}Relationship')
    e.set('Id',rid);e.set('Type',NS['r']+'/'+typ);e.set('Target',target)

def equation_run(expr,pt=12):
    n=len(manifest)+1;stem=f'mathtype{n:02d}';tex=latex(expr)
    png=BASE/(stem+'.png');binary=BASE/(stem+'.bin')
    math_to_image('$'+tex+'$',str(png),prop=FontProperties(size=pt),dpi=360,format='png',color='black')
    im=Image.open(png);w,h=im.size;wp,hp=w/5,h/5
    data=make_ole(expr,binary)
    files['word/media/'+stem+'.png']=png.read_bytes();files['word/embeddings/'+stem+'.bin']=binary.read_bytes()
    ir=f'rIdMTImage{n}';er=f'rIdMTObject{n}'
    relationship(ir,'image','media/'+stem+'.png');relationship(er,'oleObject','embeddings/'+stem+'.bin')
    run=el('w:r');rp=el('w:rPr');rp.append(el('w:noProof'));run.append(rp)
    obj=el('w:object',**{'w:dxaOrig':round(wp*20),'w:dyaOrig':round(hp*20)})
    # VML previews are the display cache for native MathType OLE objects.
    st=el('v:shapetype',id='_x0000_t75',coordsize='21600,21600',**{'o:spt':'75','o:preferrelative':'t','path':'m@4@5l@4@11@9@11@9@5xe','filled':'f','stroked':'f'})
    st.append(el('v:stroke',joinstyle='miter'));st.append(el('v:path',**{'o:extrusionok':'f','gradientshapeok':'t','o:connecttype':'rect'}))
    obj.append(st)
    shapeid=f'_x0000_i{1200+n}'
    shape=el('v:shape',id=shapeid,type='#_x0000_t75',style=f'width:{wp:.2f}pt;height:{hp:.2f}pt',**{'o:ole':''})
    shape.append(el('v:imagedata',**{'r:id':ir,'o:title':tex}));obj.append(shape)
    obj.append(el('o:OLEObject',Type='Embed',ProgID='Equation.DSMT4',ShapeID=shapeid,DrawAspect='Content',ObjectID=f'_{1800000000+n}',**{'r:id':er}))
    run.append(obj)
    manifest.append({'id':n,'latex':tex,'native_bytes':len(data),'width_pt':wp,'height_pt':hp})
    return run

def text_run(s,bold=False):
    r=el('w:r');rp=el('w:rPr');rp.append(el('w:rFonts',**{'w:ascii':'Times New Roman','w:hAnsi':'Times New Roman','w:eastAsia':'宋体'}))
    rp.append(el('w:sz',**{'w:val':22}));rp.append(el('w:szCs',**{'w:val':22}))
    if bold: rp.append(el('w:b'))
    r.append(rp);t=el('w:t');t.set('{http://www.w3.org/XML/1998/namespace}space','preserve');t.text=s;r.append(t);return r
def para(*parts,keep=False):
    p=el('w:p');pr=el('w:pPr');pr.append(el('w:spacing',**{'w:after':80,'w:line':278,'w:lineRule':'auto'}))
    if keep: pr.append(el('w:keepNext'))
    pr.append(el('w:widowControl'));p.append(pr)
    for x in parts: p.append(text_run(x) if isinstance(x,str) else equation_run(x,11))
    return p
def display(i):
    p=para();pr=p.find(tag('w:pPr'));tabs=el('w:tabs')
    tabs.append(el('w:tab',**{'w:val':'center','w:pos':'4153'}));tabs.append(el('w:tab',**{'w:val':'right','w:pos':'8306'}));pr.append(tabs)
    for obj in [el('w:tab'),equation_run(EQ[i-1]),el('w:tab')]:
        if obj.tag==tag('w:r'): p.append(obj)
        else: r=el('w:r');r.append(obj);p.append(r)
    p.append(text_run(f'({i})'));return p
def replace(old,nodes):
    parent=old.getparent();idx=parent.index(old);parent.remove(old)
    for n,node in enumerate(nodes): parent.insert(idx+n,node)

replace(ps[5],[para('本项目要求实际采样空间间隔及分析站位间距原则上不大于 0.5 mm。'),
    para('采样率应满足：',keep=True),display(1),para('其中：',keep=True),
    para('•  ',fs,'：采样率，Hz；'),para('•  ',vmax,'：实际最大 TCP 速度，mm/s；'),
    para('•  ',ds,'：允许的最大采样空间间隔，mm。'),
    para('同时应满足测量系统动态带宽要求。'),
    para('小圆、短线段及转角附近应通过加密采样验证峰值捕获能力。预试验将采样或分析间距减半后，主要偏差结果的变化应不大于 0.02 mm；否则应进一步加密或更换测量系统。'),
    para('不得通过增加插值点代替真实采样。')])

symbols=[(seq('i=1,…,m'),'指令轨迹站位索引'),(T('m'),'有效分析站位数'),
    (seq('j=1,…,n'),'同方向重复运行索引'),(T('n'),'重复运行次数，本方案取 10'),
    (seq('k=1,…,',nj),'第 j 次运行的原始采样索引'),(nj,'第 j 次评价段的有效采样点数'),
    (qi,'第 i 个指令站位的三维位置向量'),(rij,'第 j 次实际轨迹与第 i 个法平面的对应交点'),
    (gi,'第 i 个站位的 n 个交点位置重心'),(ajk,'第 j 次第 k 个原始有效采样位置'),
    (norm(T('·')),'三维欧氏范数')]
table=ps[38].getparent().getparent().getparent()
assert table.tag==tag('w:tbl') and len(table.findall(tag('w:tr')))==1
pr=table.find(tag('w:tblPr'))
if pr is not None:
    width=pr.find(tag('w:tblW'))
    if width is not None: width.set(tag('w:w'),'8306');width.set(tag('w:type'),'dxa')
header=table.find(tag('w:tr'))
for expr,definition in symbols:
    row=el('w:tr');trpr=el('w:trPr');trpr.append(el('w:cantSplit'));row.append(trpr)
    for col,content in enumerate([expr,definition]):
        cell=el('w:tc');cp=el('w:tcPr');cp.append(el('w:tcW',**{'w:w':2500 if col==0 else 5806,'w:type':'dxa'}));cp.append(el('w:vAlign',**{'w:val':'center'}));cell.append(cp)
        cell.append(para(content));row.append(cell)
    table.append(row)
# Keep the short symbol table together, with its section title.
for p in table.findall('.//w:p',NS)[:-1]:
    ppr=p.find(tag('w:pPr'))
    if ppr is None: ppr=el('w:pPr');p.insert(0,ppr)
    if ppr.find(tag('w:keepNext')) is None: ppr.append(el('w:keepNext'))
ps[37].find(tag('w:pPr')).append(el('w:keepNext'))
replace(ps[40],[])

heading=deepcopy(ps[48])
for r in list(heading):
    if r.find('.//w:drawing',NS) is not None: heading.remove(r)
heading.find(tag('w:pPr')).append(el('w:keepNext'))
replace(ps[48],[para('第 i 个站位的交点位置重心为：',keep=True),display(2),
    para('位置轨迹准确度为：',keep=True),display(3),
    para('该指标先对同一站位的 10 次交点求重心，再沿轨迹取重心与指令站位之间距离的最大值。'),heading])
replace(ps[49],[para('交点到该站位重心的距离为：',keep=True),display(4),
    para('平均距离为：',keep=True),display(5),para('距离的样本标准差为：',keep=True),display(6),
    para('位置轨迹重复性为：',keep=True),display(7),
    para('RTp 用于反映重复运行的离散程度，应单独报告。本方案未设置独立 RTp 限值，不将其替代轨迹准确度要求。')])
replace(ps[51],[para('设 ',sjk,' 为依据程序顺序确定的当前有限指令线段或圆弧，则：',keep=True),display(8),
    para('第 j 次运行的最大偏差为：',keep=True),display(9),para('全部循环中的最大偏差为：',keep=True),display(10),
    para('转角过渡处可采用预先规定的相邻有限路径段并集作为参考。'),
    para('不得在整条组合轨迹中任意选择其他路径段来降低计算误差。直线应采用有限线段距离，不能使用无限延长线掩盖 P35 的超越。'),
    para('Emax 应覆盖全部有效评价采样点，包含高度偏差、转角及换向过程。')])

# Keep numbered section headings with their first paragraph/formula.
for p in root.findall('.//w:p',NS):
    text=''.join(p.xpath('.//w:t/text()',namespaces=NS))
    if re.match(r'^\d{1,2}(?:\.\d+)?\s',text):
        pr=p.find(tag('w:pPr'))
        if pr is None: pr=el('w:pPr');p.insert(0,pr)
        if pr.find(tag('w:keepNext')) is None: pr.append(el('w:keepNext'))
# Original document text and other package parts are retained.
for r in list(rels):
    if r.get('Id') in {'rId5','rId6','rId7','rId8','rId9'}: rels.remove(r)
for i in range(1,6): files.pop(f'word/media/image{i}.png')
ctns='{http://schemas.openxmlformats.org/package/2006/content-types}'
if not any(x.get('Extension')=='bin' for x in ct):
    E.SubElement(ct,ctns+'Default',Extension='bin',ContentType='application/vnd.openxmlformats-officedocument.oleObject')
files['word/document.xml']=E.tostring(root,xml_declaration=True,encoding='UTF-8',standalone=True)
files['word/_rels/document.xml.rels']=E.tostring(rels,xml_declaration=True,encoding='UTF-8',standalone=True)
files['[Content_Types].xml']=E.tostring(ct,xml_declaration=True,encoding='UTF-8',standalone=True)
with ZipFile(OUT,'w',ZIP_DEFLATED) as z:
    for path,data in files.items(): z.writestr(path,data)
(BASE/'conversion_manifest.json').write_text(json.dumps(manifest,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps({'output':str(OUT),'objects':len(manifest),'equations':10,'symbols':11},ensure_ascii=False))
