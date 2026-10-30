from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.style import WD_STYLE_TYPE
from pathlib import Path

OUT = Path(r'E:\试切实验\机器人试切实验_碳纤维复材毛刺分析报告.docx')
IMG_DIR = Path(r'C:\Users\86195\AppData\Local\Temp')
images = [
    (IMG_DIR/'codex-clipboard-8af5440c-59f0-46dc-a36d-90ae0e60cfd1.jpg', '图1  3k，不开', '连续锯齿状毛刺，局部有较大的纤维状突起。'),
    (IMG_DIR/'codex-clipboard-c819c119-e380-4809-b6b0-04805e27e1c9.jpg', '图2  3k，开', '边缘较平直，仅见少量短小毛刺。'),
    (IMG_DIR/'codex-clipboard-8acb49ba-9468-4fc4-b4e2-e725de306820.jpg', '图3  4k，不开', '边缘存在较密集细小毛刺和局部纤维毛刺。'),
    (IMG_DIR/'codex-clipboard-21878c1a-f9fe-4c1e-87bc-b7269d8965e5.jpg', '图4  4k，开', '边缘粗糙，可见纤维拔出和碎屑，属于异常偏差工况。'),
    (IMG_DIR/'codex-clipboard-081e4d19-b84b-46ef-b502-967052c9d62f.jpg', '图5  5k，不开', '上侧边缘有连续细小毛刺，下侧相对整齐。'),
    (IMG_DIR/'codex-clipboard-7b444b70-5598-4ace-83f4-290e65bdbb03.jpg', '图6  5k，开', '边缘最平整，毛刺少且短，当前照片中表现最好。'),
    (IMG_DIR/'codex-clipboard-47615a87-faf2-4bdc-bc6f-20b57d43b7a1.jpg', '图7  6k，不开', '两侧毛刺不规则，局部有较大突起和松散颗粒。'),
    (IMG_DIR/'codex-clipboard-8e6d909d-61a2-42d7-a998-b8eb925efaec.jpg', '图8  6k，开', '仍有少量细小毛刺，但整体边缘较整齐。'),
    (IMG_DIR/'codex-clipboard-e9eb3b8f-88b6-4732-93f6-981d3d90abb0.jpg', '图9  8k，不开', '毛刺连续且密集，局部突起较大，纤维撕裂最明显。'),
    (IMG_DIR/'codex-clipboard-e803d73a-8eeb-4f3d-97eb-f8aac434ed41.jpg', '图10  5k，开（重复照片）', '与图6内容相同或高度相似，作为同一工况的重复照片。'),
]
for p,_,_ in images:
    if not p.exists():
        raise FileNotFoundError(p)

def set_cell_shading(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = tcPr.find(qn('w:shd'))
    if shd is None:
        shd = OxmlElement('w:shd')
        tcPr.append(shd)
    shd.set(qn('w:fill'), fill)

def set_cell_borders(cell, color='D9D9D9', sz='6'):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    borders = tcPr.first_child_found_in('w:tcBorders')
    if borders is None:
        borders = OxmlElement('w:tcBorders')
        tcPr.append(borders)
    for edge in ('top','left','bottom','right','insideH','insideV'):
        tag = 'w:' + edge
        element = borders.find(qn(tag))
        if element is None:
            element = OxmlElement(tag)
            borders.append(element)
        element.set(qn('w:val'),'single')
        element.set(qn('w:sz'),sz)
        element.set(qn('w:space'),'0')
        element.set(qn('w:color'),color)

def set_cell_margins(cell, top=100, start=120, bottom=100, end=120):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = tcPr.first_child_found_in('w:tcMar')
    if tcMar is None:
        tcMar = OxmlElement('w:tcMar')
        tcPr.append(tcMar)
    for m, v in [('top',top),('start',start),('bottom',bottom),('end',end)]:
        node = tcMar.find(qn('w:'+m))
        if node is None:
            node = OxmlElement('w:'+m)
            tcMar.append(node)
        node.set(qn('w:w'), str(v))
        node.set(qn('w:type'),'dxa')

def prevent_row_split(row):
    trPr = row._tr.get_or_add_trPr()
    cant = OxmlElement('w:cantSplit')
    trPr.append(cant)

def set_run_font(run, name='Microsoft YaHei', size=10.5, bold=False, color='000000'):
    run.font.name = name
    run._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'), name)
    run._element.get_or_add_rPr().rFonts.set(qn('w:ascii'), name)
    run._element.get_or_add_rPr().rFonts.set(qn('w:hAnsi'), name)
    run.font.size = Pt(size)
    run.bold = bold
    run.font.color.rgb = RGBColor.from_string(color)

def set_para(paragraph, before=0, after=6, line=1.15, align=None):
    fmt = paragraph.paragraph_format
    fmt.space_before = Pt(before)
    fmt.space_after = Pt(after)
    fmt.line_spacing = line
    if align is not None:
        paragraph.alignment = align

def add_text(doc, text, size=10.5, bold=False, align=None, before=0, after=6):
    p = doc.add_paragraph()
    set_para(p, before, after, 1.15, align)
    r = p.add_run(text)
    set_run_font(r, size=size, bold=bold)
    return p

def add_heading(doc, text, level=1):
    p = doc.add_paragraph(style=f'Heading {level}')
    set_para(p, before=8 if level==1 else 5, after=5, line=1.1)
    r = p.add_run(text)
    set_run_font(r, size=14 if level==1 else 11.5, bold=True)
    return p

doc = Document()
sec = doc.sections[0]
sec.top_margin = Cm(1.8)
sec.bottom_margin = Cm(1.8)
sec.left_margin = Cm(1.9)
sec.right_margin = Cm(1.9)

# base styles
styles = doc.styles
for sty_name in ['Normal','Title','Heading 1','Heading 2']:
    st = styles[sty_name]
    st.font.name = 'Microsoft YaHei'
    st._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')
    st._element.rPr.rFonts.set(qn('w:ascii'), 'Microsoft YaHei')
    st._element.rPr.rFonts.set(qn('w:hAnsi'), 'Microsoft YaHei')
    st.font.color.rgb = RGBColor(0,0,0)
styles['Normal'].font.size = Pt(10.5)
styles['Title'].font.size = Pt(22)
styles['Title'].font.bold = True

# footer
footer = sec.footer
fp = footer.paragraphs[0]
fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
fr = fp.add_run('机器人试切实验  |  碳纤维复材周铣毛刺分析')
set_run_font(fr, size=8.5, color='666666')

# Cover
p = doc.add_paragraph(style='Title')
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
set_para(p, before=48, after=10, line=1.0)
for r in p.runs:
    set_run_font(r, size=22, bold=True)
# title run manually because style paragraph had no run
if not p.runs:
    r = p.add_run('机器人试切实验碳纤维复材周铣毛刺分析报告')
    set_run_font(r, size=22, bold=True)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
set_para(p, before=0, after=20, line=1.0)
r = p.add_run('超声刀柄开启状态与主轴转速对比')
set_run_font(r, size=13, color='404040')

info = doc.add_table(rows=3, cols=2)
info.alignment = WD_TABLE_ALIGNMENT.CENTER
info.autofit = False
for row in info.rows:
    prevent_row_split(row)
    for cell in row.cells:
        set_cell_borders(cell)
        set_cell_margins(cell, 100, 160, 100, 160)
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
info.cell(0,0).text='实验对象'
info.cell(0,1).text='碳纤维增强复合材料周铣边缘'
info.cell(1,0).text='主要变量'
info.cell(1,1).text='主轴转速、超声刀柄开关状态'
info.cell(2,0).text='照片数量'
info.cell(2,1).text='10 张照片，对应 9 个有效工况（5k，开重复 1 张）'
for i in range(3):
    info.cell(i,0).width = Cm(3.0)
    info.cell(i,1).width = Cm(10.5)
    set_cell_shading(info.cell(i,0), 'EAF2F8')
    for j in range(2):
        for para in info.cell(i,j).paragraphs:
            set_para(para, after=0, line=1.1)
            for rr in para.runs:
                set_run_font(rr, size=10.5, bold=(j==0))

add_text(doc, '报告结论', size=13, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, before=28, after=5)
add_text(doc, '在当前照片范围内，5k、超声开启时的边缘最平整、毛刺最少；3k 开和 6k 开次之。不开超声时，6k 和 8k 的毛刺明显加重，其中 8k 不开最差。超声总体有利于抑制纤维拔出，但 4k 开出现异常偏差，表明效果存在参数窗口。', size=11, after=0)

# Main report
add_heading(doc, '1 观察范围与评价方法', 1)
add_text(doc, '本报告根据照片中标注的主轴转速和超声刀柄开关状态，对切削边缘的毛刺、纤维拔出、碎屑和局部撕裂进行相对评价。照片未提供比例尺，因此等级仅用于工况之间的比较，不代表实际毛刺高度或宽度。', after=5)
add_text(doc, '照片中主要可见三类边缘缺陷：连续锯齿状毛刺、纤维端部拔出形成的毛刺，以及附着在切边上的松散碎屑。暂未观察到大面积贯穿性分层。', after=8)

add_heading(doc, '2 各工况照片分析', 1)
data = [
    ('3k，不开','连续锯齿状毛刺，局部有较大的纤维状突起','中等偏重'),
    ('3k，开','边缘较平直，仅见少量短小毛刺','较轻'),
    ('4k，不开','细小毛刺较密集，局部有纤维毛刺','中等'),
    ('4k，开','边缘粗糙，纤维拔出和碎屑较明显','较重'),
    ('5k，不开','上侧连续细小毛刺，下侧相对整齐','中等偏轻'),
    ('5k，开','边缘最平整，毛刺少且短','最轻'),
    ('6k，不开','毛刺不规则，局部有较大突起和松散颗粒','较重'),
    ('6k，开','少量细小毛刺，整体较整齐','轻至中等'),
    ('8k，不开','毛刺连续密集，突起较大，纤维撕裂明显','最重'),
]
t = doc.add_table(rows=1, cols=4)
t.alignment = WD_TABLE_ALIGNMENT.CENTER
t.autofit = False
headers = ['工况','边缘形貌','毛刺等级','综合判断']
for j,h in enumerate(headers):
    c=t.rows[0].cells[j]
    c.text=h
    set_cell_shading(c,'2F5597')
    set_cell_borders(c)
    set_cell_margins(c,110,120,110,120)
    c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
    for p in c.paragraphs:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_para(p,after=0,line=1.0)
        for r in p.runs: set_run_font(r,size=9.5,bold=True,color='FFFFFF')
prevent_row_split(t.rows[0])
for idx,row in enumerate(data, start=1):
    cells=t.add_row().cells
    for j,val in enumerate(row):
        cells[j].text=val
        set_cell_borders(cells[j])
        set_cell_margins(cells[j],100,120,100,120)
        cells[j].vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
        if idx%2==0: set_cell_shading(cells[j],'F5F7FA')
        for p in cells[j].paragraphs:
            set_para(p,after=0,line=1.05)
            if j in (0,2,3): p.alignment=WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs: set_run_font(r,size=9.2,bold=(j==0))
    prevent_row_split(t.rows[-1])
widths=[Cm(2.4),Cm(7.0),Cm(2.3),Cm(3.2)]
for row in t.rows:
    for j,w in enumerate(widths): row.cells[j].width=w

add_heading(doc, '3 转速与超声状态的影响', 1)
add_text(doc, '不开超声时，3k～5k 的毛刺处于中等水平；当转速提高到 6k 和 8k 后，边缘毛刺明显增多，且出现较大的纤维突起和松散颗粒。仅从本组照片看，转速升高并未带来持续的边缘质量改善。', after=5)
add_text(doc, '开启超声后，3k、5k 和 6k 的毛刺均明显减小，说明超声振动可能降低了切削过程中的纤维弯曲、拉扯和拔出。4k 开的照片却表现较粗糙，属于与整体趋势不一致的工况，可能与超声振幅匹配、刀具动态状态、进给参数、纤维方向或局部夹持条件有关。', after=5)
add_text(doc, '综合排序（由好到差）可暂定为：5k 开，其次为 3k 开和 6k 开；8k 不开最差。由于缺少 8k 开照片，尚不能判断超声对 8k 工况的改善幅度。', after=8)

add_heading(doc, '4 试验结论与后续建议', 1)
for txt in [
    '建议优先以 5k、超声开启作为后续试验的基准工况，并进行至少 3 次重复验证。',
    '对 4k、超声开启进行复测，同时检查超声振幅、刀柄谐振状态、刀具跳动、刀具磨损和工件夹持刚度。',
    '补充 8k、超声开启工况，以形成完整的转速—超声对照。',
    '后续采用固定拍摄角度、光照和放大倍率，并测量毛刺最大高度、单位长度毛刺数量、纤维拔出长度和分层宽度。',
    '若高转速条件下仍出现明显毛刺，应优先检查刀具跳动、进给量、每齿进给和切削方向，而不宜只通过继续提高主轴转速来改善。',
]:
    p=doc.add_paragraph(style='List Bullet')
    set_para(p, after=3, line=1.1)
    r=p.add_run(txt); set_run_font(r,size=10.5)

# Photo appendix

doc.add_page_break()
add_heading(doc, '附录 试切边缘照片', 1)
add_text(doc, '以下照片按附件顺序排列。图 6 和图 10 均标注为 5k，开，图 10 与图 6 内容重复，作为重复照片保留。', size=10, after=8)
photo_table = doc.add_table(rows=0, cols=2)
photo_table.alignment = WD_TABLE_ALIGNMENT.CENTER
photo_table.autofit = False
for i in range(0,len(images),2):
    row=photo_table.add_row()
    prevent_row_split(row)
    for j in range(2):
        cell=row.cells[j]
        set_cell_borders(cell)
        set_cell_margins(cell,120,100,140,100)
        cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
        if i+j < len(images):
            pth, cap, note = images[i+j]
            p=cell.paragraphs[0]
            p.alignment=WD_ALIGN_PARAGRAPH.CENTER
            set_para(p,after=3,line=1.0)
            run=p.add_run()
            run.add_picture(str(pth), width=Inches(3.02))
            cp=cell.add_paragraph()
            cp.alignment=WD_ALIGN_PARAGRAPH.CENTER
            set_para(cp,after=2,line=1.0)
            rr=cp.add_run(cap)
            set_run_font(rr,size=9.2,bold=True)
            np=cell.add_paragraph()
            np.alignment=WD_ALIGN_PARAGRAPH.LEFT
            set_para(np,after=0,line=1.0)
            nr=np.add_run(note)
            set_run_font(nr,size=8.5,color='555555')
        else:
            cell.text=''
    row.cells[0].width=Cm(8.1); row.cells[1].width=Cm(8.1)

# Core properties
props=doc.core_properties
props.title='机器人试切实验碳纤维复材周铣毛刺分析报告'
props.subject='超声刀柄与主轴转速对碳纤维复材周铣毛刺的照片分析'
props.author='Codex'
props.keywords='碳纤维复材, 周铣, 超声刀柄, 毛刺, 机器人试切'

doc.save(str(OUT))
print(OUT)
