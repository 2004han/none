# -*- coding: utf-8 -*-
"""生成《基于相空间重构的航空蒙皮边缘机器人铣削表面粗糙度在线监测》中文译文 DOCX。
保留原文结构：双栏、章节编号、公式编号 (1)-(34)、表 I-IV、图 1-14 标题位置。
"""
import os
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_LINE_SPACING
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT = r"E:\试切实验\译文_基于相空间重构的航空蒙皮边缘机器人铣削表面粗糙度在线监测.docx"

COL_W = 8.5  # cm, 双栏单栏宽
BODY_PT = 9.5
EQ_PT = 8.5
TAB_PT = 7.5
CAP_PT = 8

doc = Document()

# ---------- 基础样式 ----------
st = doc.styles['Normal']
st.font.name = 'Times New Roman'
st.font.size = Pt(BODY_PT)
st.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
pf = st.paragraph_format
pf.space_before = Pt(0)
pf.space_after = Pt(0)
pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
pf.first_line_indent = Pt(0)

sec = doc.sections[0]
sec.page_width = Cm(21.0)
sec.page_height = Cm(29.7)
sec.left_margin = Cm(1.8)
sec.right_margin = Cm(1.8)
sec.top_margin = Cm(1.7)
sec.bottom_margin = Cm(1.7)


def set_cols(section, n):
    cols = section._sectPr.find(qn('w:cols'))
    if cols is None:
        cols = OxmlElement('w:cols')
        section._sectPr.append(cols)
    cols.set(qn('w:num'), str(n))
    cols.set(qn('w:space'), '340')


def run_font(run, size=None, bold=False, italic=False, color=None, ea='宋体', latin='Times New Roman'):
    run.font.name = latin
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn('w:rFonts'))
    if rfonts is None:
        rfonts = OxmlElement('w:rFonts')
        rpr.append(rfonts)
    rfonts.set(qn('w:ascii'), latin)
    rfonts.set(qn('w:hAnsi'), latin)
    rfonts.set(qn('w:eastAsia'), ea)
    if size is not None:
        run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    if color is not None:
        run.font.color.rgb = color


def para(text='', size=BODY_PT, bold=False, italic=False, align=None, indent_first=0.0,
         space_before=0, space_after=0, color=None, keep=False):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    f = p.paragraph_format
    f.space_before = Pt(space_before)
    f.space_after = Pt(space_after)
    if indent_first:
        f.first_line_indent = Cm(indent_first)
    if keep:
        f.keep_with_next = True
    if text:
        r = p.add_run(text)
        run_font(r, size, bold, italic, color)
    return p


def h_section(text):
    p = para(text, size=10, bold=False, space_before=8, space_after=3, keep=True)
    for r in p.runs:
        r.font.small_caps = True
    return p


def h_sub(text):
    p = para(text, size=BODY_PT, bold=False, italic=True, space_before=4, space_after=2, keep=True)
    return p


def eq(body, num, size=EQ_PT):
    p = doc.add_paragraph()
    f = p.paragraph_format
    f.space_before = Pt(2)
    f.space_after = Pt(2)
    f.left_indent = Cm(0.25)
    f.tab_stops.add_tab_stop(Cm(COL_W - 0.35), WD_TAB_ALIGNMENT.RIGHT)
    r = p.add_run(body)
    run_font(r, size, italic=False)
    r2 = p.add_run('\t(%s)' % num)
    run_font(r2, size)
    return p


def eq_free(body, size=EQ_PT):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(2)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(body)
    run_font(r, size)
    return p


def table_caption(text):
    p = para(text, size=CAP_PT, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=6, space_after=2, keep=True)
    for r in p.runs:
        r.font.small_caps = True
    return p


def fig_caption(text):
    p = para(text, size=CAP_PT, align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=2, space_after=6)
    return p


def fig(n, caption, page):
    p = para('［图 %d 位置 —— 原图见 PDF 第 %d 页］' % (n, page), size=7, italic=True,
             align=WD_ALIGN_PARAGRAPH.CENTER, space_before=3, space_after=1,
             color=RGBColor(0x88, 0x88, 0x88))
    fig_caption('图 %d　%s' % (n, caption))


def fig_placeholder(note):
    p = para(note, size=7.5, italic=True, align=WD_ALIGN_PARAGRAPH.CENTER,
             space_before=2, space_after=1, color=RGBColor(0x80, 0x80, 0x80))
    return p


def _border(el, edge, sz, val='single'):
    b = OxmlElement('w:' + edge)
    b.set(qn('w:val'), val)
    b.set(qn('w:sz'), str(sz))
    b.set(qn('w:space'), '0')
    b.set(qn('w:color'), '000000')
    el.append(b)


def fix_table_layout(t, widths):
    """强制固定表格布局与列宽，避免 LibreOffice/Word 自动扩展列宽导致越栏。"""
    tbl = t._tbl
    tblPr = tbl.tblPr
    for tag in ('w:tblW', 'w:tblLayout'):
        el = tblPr.find(qn(tag))
        if el is not None:
            tblPr.remove(el)
    tblW = OxmlElement('w:tblW')
    tblW.set(qn('w:w'), str(int(sum(widths) * 567)))
    tblW.set(qn('w:type'), 'dxa')
    tblPr.append(tblW)
    layout = OxmlElement('w:tblLayout')
    layout.set(qn('w:type'), 'fixed')
    tblPr.append(layout)
    grid = tbl.find(qn('w:tblGrid'))
    if grid is not None:
        tbl.remove(grid)
    grid = OxmlElement('w:tblGrid')
    for w in widths:
        gc = OxmlElement('w:gridCol')
        gc.set(qn('w:w'), str(int(w * 567)))
        grid.append(gc)
    tblPr.addnext(grid)
    for j, w in enumerate(widths):
        try:
            t.columns[j].width = Cm(w)
        except Exception:
            pass
    for row in t.rows:
        for j, w in enumerate(widths):
            if j < len(row.cells):
                row.cells[j].width = Cm(w)


def make_table(data, widths, header_rows=1, align_center=True, first_col_left=True, bold_rows=()):
    rows = len(data)
    cols = len(data[0])
    t = doc.add_table(rows=rows, cols=cols)
    t.style = 'Table Grid'
    t.alignment = WD_TABLE_ALIGNMENT.CENTER

    # 表格边框：仅上下横线 + 表头下横线
    tblPr = t._tbl.tblPr
    for old in tblPr.findall(qn('w:tblBorders')):
        tblPr.remove(old)
    borders = OxmlElement('w:tblBorders')
    _border(borders, 'top', 10)
    _border(borders, 'bottom', 10)
    _border(borders, 'left', 0, 'none')
    _border(borders, 'right', 0, 'none')
    _border(borders, 'insideH', 0, 'none')
    _border(borders, 'insideV', 0, 'none')
    tblPr.append(borders)

    for i, row in enumerate(data):
        for j, val in enumerate(row):
            cell = t.cell(i, j)
            cell.width = Cm(widths[j])
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1)
            p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            if align_center and not (first_col_left and j == 0):
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            r = p.add_run(str(val))
            run_font(r, TAB_PT, bold=(i in bold_rows))
    # 表头下横线
    for i in range(header_rows):
        for j in range(cols):
            tcPr = t.cell(i, j)._tc.get_or_add_tcPr()
            tcb = OxmlElement('w:tcBorders')
            _border(tcb, 'bottom', 6)
            tcPr.append(tcb)
    # 行不跨页
    for row in t.rows:
        trPr = row._tr.get_or_add_trPr()
        c = OxmlElement('w:cantSplit')
        trPr.append(c)
    fix_table_layout(t, widths)
    return t


def bullet(text, size=BODY_PT):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.left_indent = Cm(0.7)
    p.paragraph_format.first_line_indent = Cm(-0.35)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r = p.add_run(text)
    run_font(r, size)
    return p


def numbered(num, text, size=BODY_PT):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.left_indent = Cm(0.85)
    p.paragraph_format.first_line_indent = Cm(-0.85)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r = p.add_run(num + '\t')
    run_font(r, size)
    p.paragraph_format.tab_stops.add_tab_stop(Cm(0.85))
    r2 = p.add_run(text)
    run_font(r2, size)
    return p


# =====================================================================
# 标题区（通栏）
# =====================================================================
p = para('基于相空间重构的航空蒙皮边缘', size=15, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=0)
p = para('机器人铣削表面粗糙度在线监测', size=15, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=4)
para('Phase Space Reconstruction-Based Online Roughness Monitoring in Robotic Milling of Aircraft Skin Edge',
     size=9.5, italic=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=8,
     color=RGBColor(0x44, 0x44, 0x44))

para('张涵宇，徐　蔚（Member, IEEE），彭庆宇，王东方，蒋　诚，李文龙（Member, IEEE）',
     size=11, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
para('Han-yu Zhang, Wei Xu, Qing-yu Peng, Dong-fang Wang, Cheng Jiang, Wen-long Li',
     size=9, italic=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)
para('（华中科技大学 智能制造装备与技术全国重点实验室，武汉 430074）',
     size=8.5, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)

p = para('收稿日期：2025 年 1 月 4 日；修回：2025 年 3 月 26 日、2025 年 5 月 26 日；录用：2025 年 8 月 20 日；'
         '出版：2025 年 10 月 3 日；当前版本：2026 年 4 月 20 日。推荐：技术编辑 A. Visioli、高级编辑 M. Indri。'
         '本工作部分受国家自然科学基金（批准号 52205524、52575573、52188102）、中国博士后科学基金（2022M721247）'
         '以及智能制造装备与技术全国重点实验室开放课题（IMETKF2025001）资助。（通信作者：徐蔚。）'
         '作者单位：华中科技大学智能制造装备与技术全国重点实验室，武汉 430074，中国。',
         size=7.5, space_after=2,
         color=RGBColor(0x33, 0x33, 0x33))

# ---------- 通栏结束，进入双栏 ----------
sec2 = doc.add_section(WD_SECTION.CONTINUOUS)
set_cols(sec2, 2)
# 关键：python-docx 的 add_section 会克隆节属性，需逐节重新写入页面几何，避免首节沿用默认页宽
for _s in doc.sections:
    _s.page_width = Cm(21.0)
    _s.page_height = Cm(29.7)
    _s.left_margin = Cm(1.8)
    _s.right_margin = Cm(1.8)
    _s.top_margin = Cm(1.7)
    _s.bottom_margin = Cm(1.7)
set_cols(doc.sections[0], 1)
set_cols(doc.sections[1], 2)

# =====================================================================
# 摘要
# =====================================================================
p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(4)
p.paragraph_format.space_after = Pt(3)
r = p.add_run('摘要——')
run_font(r, 8.5, bold=True)
r = p.add_run('在机器人铣削中，实时监测表面粗糙度对维持航空蒙皮边缘的质量至关重要。然而，由于铣削过程工况复杂，'
              '现有技术在精度与鲁棒性方面往往难以胜任。本文提出了一种面向航空蒙皮边缘机器人铣削表面粗糙度实时预测的新方法。'
              '该方法首先利用频响函数对振动信号进行动态调整，以保持预测精度；随后采用相空间重构，将振动信号从时域映射到'
              '与机器人系统微分同胚的相空间，从而揭示该系统的动力学特征；进而使用格拉姆角求和场（GASF）算法从相空间轨迹中'
              '提取特征，并用卷积神经网络（CNN）进行分类。基于实际实验构建了振动—粗糙度数据集，用于训练所提模型。'
              '加工系统的可预测性得到了定量证明。此外，模型的实用性通过在线实验得到验证。基于测试集的预测结果证实，'
              '与现有方法相比，该方法在预测精度与鲁棒性方面均有显著提升。')
run_font(r, 8.5)

p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(3)
p.paragraph_format.space_after = Pt(6)
r = p.add_run('关键词——')
run_font(r, 8.5, bold=True)
r = p.add_run('航空蒙皮零件，在线粗糙度监测，相空间重构，机器人铣削。')
run_font(r, 8.5, italic=True)

# ---------- 术语表 ----------
table_caption('术语表（Nomenclature）')
nomen = [
    ('q(t)', '位移响应矩阵'),
    ('F(t)', '激振矩阵'),
    ('M', '系统质量矩阵'),
    ('C', '系统阻尼矩阵'),
    ('K', '系统刚度矩阵'),
    ('q(ω)', 'q(t) 的傅里叶变换表达式'),
    ('F(ω)', 'F(t) 的傅里叶变换表达式'),
    ('H(ω)=[H_ij(ω)]', '系统位移频响函数（FRF）'),
    ('H_a(ω)', '加速度频响函数'),
    ('k = 1, 2, …, K', '沿进给方向的离散点序号'),
    ('H^R_ij(ω)', 'H_ij(ω) 的实部'),
    ('H^I_ij(ω)', 'H_ij(ω) 的虚部'),
    ('ω_m', '系统模态频率'),
    ('ξ_m', '系统模态阻尼'),
    ('ε', '模态参数的变化率'),
    ('Δ_k', '相邻离散点间距'),
    ('x_t', '沿进给方向的坐标'),
    ('L', '蒙皮零件铣削区域总长度'),
    ('a(ω)', '加速度 q̈(t) 的傅里叶变换表达式'),
    ('H_a(ω)⁺', 'H_a(ω) 的 Moore–Penrose 伪逆'),
    ('F⁻¹{·}', '傅里叶逆变换'),
    ('s_i', '相空间向量'),
    ('s_η(i)', 's_i 的最近邻'),
    ('n', '振动采样点数'),
    ('m', '嵌入维数'),
    ('τ', '延迟时间'),
    ('S', '相空间轨迹'),
    ('I(P,Q)', '互信息'),
    ('τ_s', '最优延迟时间'),
    ('m_s', '最优嵌入维数'),
    ('R_m(i)', 's_i 与 s_η(i) 之间的欧氏距离'),
    ('a(i,m)', 's_i 与 s_η(i) 之间的距离比'),
    ('E(m)', '整个相空间 S 内的距离比'),
    ('E_1(m)', 'E(m) 的变化率'),
    ('δ', 'E_1(m) 导数的上限'),
    ('s̃_i', 's_i 的角度表示'),
    ('s_0', '相空间中的参考向量'),
    ('S̃', 'S 的角度表示'),
]

def make_list_table(rows, widths):
    t = doc.add_table(rows=len(rows), cols=2)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    tblPr = t._tbl.tblPr
    borders = OxmlElement('w:tblBorders')
    for edge in ('top', 'bottom', 'left', 'right', 'insideH', 'insideV'):
        _border(borders, edge, 0, 'none')
    tblPr.append(borders)
    for i, (a, b) in enumerate(rows):
        c0, c1 = t.cell(i, 0), t.cell(i, 1)
        c0.width = Cm(widths[0])
        c1.width = Cm(widths[1])
        p0 = c0.paragraphs[0]
        p0.paragraph_format.space_before = Pt(0.5)
        p0.paragraph_format.space_after = Pt(0.5)
        r = p0.add_run(a)
        run_font(r, 8)
        p1 = c1.paragraphs[0]
        p1.paragraph_format.space_before = Pt(0.5)
        p1.paragraph_format.space_after = Pt(0.5)
        r = p1.add_run(b)
        run_font(r, 8)
    fix_table_layout(t, widths)
    return t

make_list_table(nomen, [2.6, 5.6])

# =====================================================================
# I. 引言
# =====================================================================
h_section('I.　引　言')
para('由于灵活性与成本效益，机器人正越来越多地用于加工大型航空蒙皮零件[1]–[3]。一个著名的例子是 Fraunhofer IFAM '
     '为制造空客 A320 垂尾而开发的机器人铣削系统[4]，如图 1 所示。然而，机器人手臂与柔性蒙皮零件的低刚度和低稳定性'
     '所带来的振动问题[5]，对航空蒙皮零件的机器人铣削质量产生了负面影响。蒙皮边缘的表面质量对于维持飞机的气动性能、'
     '隐身性能与结构完整性至关重要。表面过于粗糙会破坏蒙皮件之间的紧密贴合、削弱表面密封效果，并降低飞机的整体性能。'
     '因此，预测表面粗糙度，尤其是在机器人铣削过程中实时预测，对于保障高质量、高效率的航空制造至关重要。',
     indent_first=0.5)
fig(1, 'Fraunhofer IFAM 为制造空客 A320 垂尾而开发的机器人铣削系统[4]。', 1)

para('表面粗糙度通常由人工在已加工表面选取特定点进行离线测量[6]。这种方法不仅耗时，而且容易产生误差与不一致，'
     '降低了整体效率。虽然统计模型[7]与数学模型[8]、[9]常被用于离线预测粗糙度，但这些方法通常无法捕捉加工过程中'
     '发生的动态变化[10]。因此，它们的效率、精度与可靠性都受到限制。相比之下，在线粗糙度监测能够在机器人铣削过程中'
     '实时检测加工异常，从而提高生产良率与整体工艺效率。', indent_first=0.5)

h_sub('A. 相关工作')
para('在线粗糙度监测一般通过振动测量间接实现，可分为非接触式与接触式两类。非接触式方法主要包括声发射[11]、[12]、'
     '激光多普勒测振[13]、[14]以及基于结构光系统的三维（3-D）点跟踪[15]、[16]。声发射由于受工业环境噪声干扰，'
     '信噪比偏低。激光多普勒测振对低频噪声的测量精度有限[16]，且只能捕捉法向振动。三维点跟踪与结构光系统虽然'
     '测量范围更大，但精度较低、信噪比更低[16]。接触式方法（如加速度计与测力仪）无法直接测量加工区域的振动，'
     '而是在加工系统的其他位置（如夹具[17]、主轴[18]、工件[19]）间接测量。在铣刀进给过程中，安装在固定点的传感器'
     '无法分辨振动变化究竟是由加工位置改变引起，还是由表面粗糙度变化引起。Liu 等[20]将切削位置与切削力、加速度'
     '监测信号进行了关联。Yao 等[21]、[22]建立了位置相关的幅值预处理函数来调整振动信号的幅值。然而，由于振动的'
     '非线性传递特性，振动传播不仅影响幅值，还会在振动信号中引入相移，而上述方法并未对此加以补偿。',
     indent_first=0.5)

para('基于信号的预测模型通过时域[26]、[27]与时频域[23]–[25]分析，从采集的振动信号中提取特征，主要包括短时傅里叶'
     '变换[23]、小波变换[24]、希尔伯特—黄变换[25]、格拉姆角场[26]以及相空间重构[27]。近年来，基于人工智能的方法'
     '因其灵活性、实用性以及更强的特征提取能力而受到关注[28]。例如，Zhao 等[29]开发了长短期记忆（LSTM）网络，'
     '用于从振动加速度信号估计材料去除率。Cao 等[30]建立了稀疏贝叶斯回归模型，通过分析实时铣削力、切削参数与'
     '位置来监测加工误差。Lu 等[31]使用卷积神经网络（CNN）分类器将振动信号与已加工表面图像联系起来，'
     'Li 等[32]则提出了一种融合加工参数与铣削力信号的深度学习模型以提高预测精度。He 等[33]应用对称点模式（SDP）'
     '—CNN 框架，从振动信号中准确预测磨削质量。现有方法依赖单通道信号或多通道信号的简单组合，并采用特征提取器'
     '来实现预测。这些方法对噪声高度敏感，且缺乏对信号中所蕴含非线性特征的全面分析。', indent_first=0.5)

h_sub('B. 贡献')
para('本文针对航空蒙皮边缘机器人铣削的粗糙度预测提出了一套完整框架，包括基于 FRF 的振动信号校正方法与'
     '相空间重构—格拉姆角求和场—卷积神经网络（PSR-GASF-CNN）粗糙度监测模型。振动信号校正方法利用离散化铣削'
     '分段的频响函数（FRF），对安装在固定点的加速度计所采集的振动信号进行补偿。通过 PSR 方法将时域振动信号映射为'
     '相空间中的轨迹，并采用 GASF 从这些轨迹中提取特征。随后由 CNN 分类器完成识别与分类，实现表面粗糙度的精确'
     '预测。本方法的主要贡献归纳如下。', indent_first=0.5)

numbered('1)', '更高的精度与鲁棒性：将振动信号校正方法引入粗糙度预测，以避免刀具路径位移引起的精度退化；同时，'
              'GASF 方法编码了时间序列中任意两个元素之间的全局关系，这种高维嵌入提升了精度与鲁棒性。')
numbered('2)', '重构的动力学特征：相空间与机器人铣削系统微分同胚。这意味着机器人铣削系统的动力学特征可由有限的'
              '观测振动信号重构出来，从而便于 CNN 提取高层特征。')
numbered('3)', '振动信号可预测性得到验证：机器人铣削系统是复杂的非线性系统，关于其振动信号可预测性的研究尚少。'
              '本文定量证明了振动信号的可预测性，从而为所提预测模型提供了支撑。')

para('本文其余部分安排如下。第 II 节介绍振动信号的校正方法；第 III 节描述 PSR-GASF-CNN 模型；第 IV 节开展'
     '实际应用实验，构建振动—粗糙度数据集，说明振动信号的可预测性，并分析监测模型的有效性；最后，第 V 节总结全文。',
     indent_first=0.5)

# =====================================================================
# II. 基于 FRF 的振动信号校正方法
# =====================================================================
h_section('II.　基于 FRF 的振动信号校正方法')
para('在蒙皮零件的机器人铣削过程中，固定位置的传感器无法适应加工点的变化。为克服这一难题，本文提出一种振动信号'
     '校正方法：将蒙皮边缘离散化，分别测量各离散点与加速度计之间的 FRF，从而对振动信号施加位置相关的校正。'
     '该方法有助于实现对粗糙度的精确监测。', indent_first=0.5)

h_sub('A. 振动传递模型')
para('在大型航空蒙皮零件的机器人铣削中，为适应零件的尺寸与复杂曲率，常采用磁铁、真空吸盘等夹持方式。然而，'
     '这些方式并不满足精确动力学理论分析所需的边界条件[34]。为绕开为粗糙度监测建立精确振动传递模型的困难，'
     '本节引入基于振动 FRF 的振动传递模型。该途径为理论方法提供了一种替代方案，能够在航空蒙皮零件的机器人铣削中'
     '更可靠地获取振动。', indent_first=0.5)

para('图 2 给出了蒙皮零件边缘铣削的振动传递模型。对于航空蒙皮零件，x 向与 y 向刚度远大于 z 向刚度；但基于磁铁与'
     '真空吸盘的夹持方式导致夹持刚度较低。在实际加工过程中，x、y、z 三个方向均存在显著振动，因此传递模型需要考虑'
     '三个方向的振动。', indent_first=0.5)
fig(2, '蒙皮零件边缘铣削的振动传递模型。', 3)

para('在蒙皮零件边缘的机器人铣削过程中，铣刀的位置随时间变化。然而，在一个采样周期内，铣刀的运动距离远小于工件'
     '尺寸。因此，在每个采样周期内可认为激励点不变；又因加速度传感器的位置固定，振动模型可以简化。该模型可转化为'
     '单激励点、单响应点的三自由度振动模型。模型的系统动力学方程为',
     indent_first=0.5)
eq('M q̈(t) + C q̇(t) + K q(t) = F(t)', '1')
para('式中，M、C、K 分别为系统的质量矩阵、阻尼矩阵与刚度矩阵；q(t)=[x(t) y(t) z(t)]ᵀ 与 F(t)=[f_x(t) f_y(t) f_z(t)]ᵀ '
     '分别表示时域的位移响应矩阵与激振矩阵。对式 (1) 作傅里叶变换可得',
     indent_first=0.5)
eq('(−ω²M + ωjC + K) q(ω) = F(ω)', '2')
para('式中，q(ω) 与 F(ω) 分别为 q(t) 与 F(t) 的傅里叶变换表达式。定义 H(ω) = (−ω²M + ωjC + K)⁻¹ 为系统的位移 FRF，'
     '则式 (2) 可改写为', indent_first=0.5)
eq('q(ω) = H(ω) F(ω)', '3')
para('H(ω) 矩阵的展开形式为', indent_first=0.5)
eq('H(ω) = [H_pq(ω)]，p, q ∈ {x, y, z}（3 × 3 矩阵）', '4')
para('其中对角元素表示动态系统的直接 FRF，非对角元素表示交叉 FRF。由于 x、y、z 坐标的正交性，直接 FRF 通常远大于'
     '交叉 FRF，因此在常规情况下 H(ω) 可近似为对角矩阵。然而，在实际模态试验中我们发现，交叉 FRF 显著到不可忽略'
     '的程度。这意味着磁力吸盘夹持方式引起的非线性使得 H(ω) 不再是对角矩阵。', indent_first=0.5)

h_sub('B. 位置相关校正')
para('在蒙皮边缘铣削过程中，位置相关校正方法对于降低铣削位置变化对模型预测精度的影响至关重要。尽管工件与夹持'
     '系统存在固有的非线性，但其仍保持了近似连续性。这意味着，通过对加工区域离散化并在离散点上进行模态锤击试验，'
     '可以近似得到蒙皮零件整个加工区域的 FRF。', indent_first=0.5)
fig(3, '航空蒙皮零件的离散空间模型。', 4)

para('图 3 给出了蒙皮零件的离散空间模型。当铣刀沿 x_t 方向进给时，式 (1) 可改写为', indent_first=0.5)
eq('M_k q̈(t) + C_k q̇(t) + K_k q(t) = F_k(t)', '5')
para('式中 k = 1, 2, …, K 表示沿进给方向的离散点序号。式 (3) 可改写为', indent_first=0.5)
eq('q(ω) = H_k(ω) F_k(ω)', '6')
para('为确定离散点，分别利用 H(ω) 的实部 H^R_ij(ω) 与虚部 H^I_ij(ω) 辨识系统模态频率与模态阻尼，即',
     indent_first=0.5)
eq('H_ij(ω) = H^R_ij(ω) + i H^I_ij(ω)', '7')
para('式中 H_ij(ω) 表示第 i 个自由度激励、第 j 个自由度响应之间的 FRF。系统模态频率 ω_m 由下式确定：',
     indent_first=0.5)
eq('dH^I(ω)/dω |_(ω=ω_m) = 0', '8')
para('系统模态阻尼 ξ_m 由下式确定：', indent_first=0.5)
eq('ξ_m = |ω_a − ω_b| / ω_m', '9')
para('式中 ω_a 与 ω_b 满足：', indent_first=0.5)
eq('dH^R(ω)/dω |_(ω=ω_a, ω_b) = 0', '10')
para('因此，可通过限制模态频率与模态阻尼的变化来确定离散点：', indent_first=0.5)
eq('|ω_m^k−ω_m^(k+1)|/ω_m^k ≤ ε，或 |ξ_m^k−ξ_m^(k+1)|/ξ_m^k ≤ ε', '11', size=7.5)
para('式中 ε 表示相邻离散点之间模态参数的变化率，实际实施中通常取 0.1。定义相邻离散点的间距为 Δ_k：',
     indent_first=0.5)
eq('Δ_k = x_(k+1) − x_k', '12')
para('蒙皮零件的 H(ω) 由下式给出：', indent_first=0.5)
eq('H(ω)|_(x_t) = H_k(ω)，x_t ∈ [Σ_(i<k) Δ_i, Σ_(i≤k) Δ_i)，k = 1, …, K', '13', size=7.2)
para('式中 x_t 表示沿进给方向的坐标，从蒙皮零件的初始切削位置起算（见图 2）；L 表示蒙皮零件铣削区域的总长度。',
     indent_first=0.5)

para('在标准模态锤击试验中通常使用加速度计，此时测得的是加速度 FRF H_a(ω)：', indent_first=0.5)
eq('a(ω) = −ω² q(ω) = H_a(ω) F(ω)', '14')
para('式中 a(ω) 为 q̈(t) 的傅里叶变换表达式。联立式 (3) 与式 (14)，加速度 FRF H_a(ω) 与位移 FRF H(ω) 的关系为',
     indent_first=0.5)
eq('H_a(ω) = −ω² H(ω)', '15')
para('由于前述夹持方式引起的非线性，式 (15) 不再能被视作对角矩阵，且难以证明其可逆性。因此，采用 Moore–Penrose '
     '伪逆来获得激振信号的最小二乘估计：', indent_first=0.5)
eq('F(ω) = H_a(ω)⁺ a(ω)', '16')
para('式中 H_a(ω)⁺ 为 H_a(ω) 的 Moore–Penrose 伪逆。边缘铣削过程中任意位置的时域激振信号可由下式获得：',
     indent_first=0.5)
eq('F(t) = F⁻¹{F(ω)}', '17')
para('式中 F⁻¹{·} 为傅里叶逆变换。至此，在固定位置采集的振动信号已被转换为实际加工区域的激振信号，'
     '完成了振动信号的校正。', indent_first=0.5)

# =====================================================================
# III. 在线粗糙度监测模型
# =====================================================================
h_section('III.　在线粗糙度监测模型')
para('由于微分同胚性，相空间相比时域与频域能更直接地呈现机器人加工系统的非线性动力学特征，有利于提取具有高'
     '辨识度的特征[35]。因此，建立基于相空间轨迹的粗糙度监测模型，对于提升模型的预测性能至关重要。',
     indent_first=0.5)

h_sub('A. 相空间重构')
para('机器人—刀具—工件系统是一个确定性的复杂非线性动力系统，其中 {f(1), f(2), …, f(n)} 表示一组观测值。'
     '变量 n 表示振动采样点数，在本文中即为第 II-B 节得到的激振序列。时域信号如图 4(a) 所示。',
     indent_first=0.5)
fig(4, '机器人铣削的振动信号。(a) 时域表示。(b) 相空间轨迹。', 5)
para('设 m 为嵌入维数、τ 为延迟时间，则一个 m 维相空间向量 s_i 可构造为', indent_first=0.5)
eq('s_i = [f(i), f(i+τ), …, f(i+(m−1)τ)]ᵀ', '18')
para('时间序列可变换为相空间 S：', indent_first=0.5)
eq('S = {s_1, s_2, …, s_(n−(m−1)τ)}', '19')
para('Takens 定理证明：通过选择合适的嵌入维数 m 与延迟时间 τ，重构相空间可以与原始动力系统微分同胚。'
     '由于动力系统与该相空间在拓扑上一致，因此可以通过相空间预测机器人—刀具—工件系统的行为。',
     indent_first=0.5)
para('定义两个时间序列 P 与 Q 为', indent_first=0.5)
eq('P = {p_1, p_2, …, p_n} = {f(1), f(2), …, f(n)}；Q = {q_1, q_2, …, q_n} = {f(1+τ), f(2+τ), …, f(n+τ)}', '20')
para('式中 P 表示一个时间序列，Q 为延迟 τ 后的时间序列。于是 P 与 Q 的互信息为', indent_first=0.5)
eq('I(P,Q) = H(Q) + H(P) − H(Q,P) = Σ_i Σ_j P_pq log₂ [P_pq / (P_p P_q)]', '21', size=7.5)
para('式中 H(Q)、H(P) 分别为 P、Q 的信息熵，H(Q,P) 为二者的联合熵。I(P,Q) 增大，表明 P 与 Q 之间的关联强度增强，'
     '反之则减弱。Fraser 与 Swinney[36] 认为，延迟时间的选择应使相关性最小。因此，选取使互信息最小的最优延迟时间 '
     'τ_s：', indent_first=0.5)
eq('τ_s = argmin { I(P,Q) : 1 ≤ τ ≤ τ_max }', '22')
para('式中 τ_max 为 τ 的最大值，以保证相空间中相邻点的相关性较低。', indent_first=0.5)
para('设 s^m_i 为 m 维相空间 S 中的一点，s^m_η(i) 为其最近邻之一，则', indent_first=0.5)
eq('η(i) = argmin { ‖s^m_i − s^m_j‖ : j = 1, 2, …, n−(m−1)τ_s, j ≠ i }', '23')
para('用欧氏距离度量两个相点之间的距离（其中 s^m_i ∈ S^m，s^m_η(i) ∈ S^m）：', indent_first=0.5)
eq('R_m(i) = ‖s^m_i − s^m_η(i)‖', '24')
para('当相空间维数由 m 增至 m+1 时，点 s^(m+1)_i 与 s^(m+1)_η(i) 之间的距离为', indent_first=0.5)
eq('R_(m+1)(i) = ‖s^(m+1)_i − s^(m+1)_η(i)‖', '25')
para('定义 s_i 与 s_η(i) 在不同维数空间中的距离比为', indent_first=0.5)
eq('a(i,m) = R_(m+1)(i) / R_m(i),  i = 1, 2, …, n−(m−1)τ_s', '26')
para('整个相空间 S 内的距离比表示为', indent_first=0.5)
eq('E(m) = (1/(n−mτ_s)) Σ_(i=1..n−mτ_s) a(i,m)', '27')
para('显然，E(m) 仅依赖维数 m。定义', indent_first=0.5)
eq('E_1(m) = E(m+1) / E(m)', '28')
para('对于确定性系统，当嵌入维数大于某一特定值 m_s 时，相空间结构被完全打开，能够充分反映系统吸引子[37]。'
     '这意味着，当 m 大于该特定值 m_s 后，E_1(m) 不再变化。通常我们取 m_s 作为系统的最优嵌入维数',
     indent_first=0.5)
eq('m_s = min { m : dE_1(m)/dm = 0 }', '29')
para('实际处理中信号存在随机噪声，E_1(m) 往往并非严格不变。因此，假设一个小量 δ，当 E_1(m) 小于 δ 时停止迭代，'
     '并根据实际噪声情况调整 δ 的取值：', indent_first=0.5)
eq('m_s = min { m : |dE_1(m)/dm| ≤ δ, δ ≪ 1 }', '30')

h_sub('B. PSR-GASF-CNN 粗糙度预测模型')
para('GASF 变换提供了一种将时间序列转换为二维特征矩阵的有效方法，能够在原始时间序列的任意两点之间进行特征提取。'
     '借助 CNN 从二维矩阵中提取特征的强大能力，GASF 有助于建立振动信号与表面粗糙度之间的映射关系。GASF 的核心'
     '在于将序列转换为角度表示：', indent_first=0.5)
eq('s̃_i = cos θ_i = (s_i · s_0) / (|s_i| |s_0|)', '31')
para('式中 s_i 为重构相空间轨迹 S 中的一个元素，s_0 表示相空间中的参考向量，s̃_i 为角度序列 '
     'S̃ = {s̃_1, s̃_2, …, s̃_(n−(m−1)τ)} 中的一个元素。一般而言，参考向量 s_0 可以任意选取；基于不同参考向量 s_0 '
     '的角度映射本质上是相同的，只是提供了不同形式的特征表示。图 5 演示了利用 GASF 对向量序列进行图像编码的过程。',
     indent_first=0.5)
fig(5, 'GASF 方法的基本原理。', 5)
para('值得注意的是，原始 GASF 变换是为标量时间序列设计的，以采样值与最大值的比值作为输入。这一做法无法直接'
     '应用于向量序列（相空间轨迹）。因此，我们采用相空间向量与参考向量之间的内积作为输入，实现了一种合理的推广，'
     '从而拓展了 GASF 的适用范围。', indent_first=0.5)
para('将向量序列转换为角度序列 S̃ 后，GASF 矩阵通过计算角度和的余弦得到：', indent_first=0.5)
eq('GASF = [cos(θ_i + θ_j)]', '32')
para('式中 i, j = 1, 2, 3, …, n−(m−1)τ。考虑到三角函数的计算开销较大，为提高计算效率，将上式变换为矩阵运算形式：',
     indent_first=0.5)
eq('GASF = [s̃_i s̃_j − √(1−s̃_i²) √(1−s̃_j²)]', '33', size=7.8)
eq_free('= S̃ᵀ·S̃ − √(I − (S̃ᵀ)²) · √(I − S̃²)', size=7.8)
para('GASF 矩阵所呈现的纹理与形状特征反映了相空间点的分布，能够体现加工系统的非线性特征。利用卷积神经网络从'
     'GASF 矩阵中提取特征，并拟合振动信号与机器人铣削表面粗糙度之间的非线性关系。', indent_first=0.5)
para('结合振动信号校正与相空间特征提取，所提粗糙度监测方法可由图 6 描述。该方法通过离线锤击试验获得工件的位置'
     '相关 FRF；同时通过离线实验构建的数据集完成 CNN 分类器的训练。在线监测时，由加速度传感器与机器人控制器分别'
     '获取振动信号 q̈(t) 与工件的实时切削位置 x_t。根据式 (17) 计算补偿后的振动激振信号 F(t)，并执行相空间重构'
     '得到相空间轨迹 S；根据式 (33) 计算 GASF 矩阵，将其输入 CNN 分类器，从而实现表面粗糙度的在线监测。',
     indent_first=0.5)
fig(6, '粗糙度监测方法的架构。', 6)

# =====================================================================
# IV. 模型训练与实际实验
# =====================================================================
h_section('IV.　模型训练与实际实验')
para('本节介绍为构建振动—粗糙度数据集并训练粗糙度监测模型而设计的实际实验方案；定量证明振动信号的可预测性；'
     '与同类方法比较，说明本文方法具有更高的精度与更强的鲁棒性；并通过在线监测试验验证预测模型的实用性。'
     '下述实验所用的全部程序代码（除另有说明）均以 MATLAB R2023a 编写，运行于配备 Intel i7-8750H CPU、'
     '16 GB DDR4 内存与 NVIDIA GTX1060 GPU 的计算机上。', indent_first=0.5)

h_sub('A. 基于实际实验的数据集构建')
para('1) 实际铣削实验设置：表面粗糙度预测模型将基于由实际蒙皮零件铣削实验构建的数据集进行训练。本文所选工件为'
     '尺寸 200 × 200 × 2 mm 的 6061 铝合金蒙皮零件。图 7 给出了实验装置的总体情况，铣削平台采用 ABB IRB-6660 '
     '机器人，末端执行器为 EBLTS TS120F 高速电主轴。加速度信号由 KSI-138M100 三轴加速度计监测，并通过商业软件 '
     'DewesoftX 与 SIRIUS X 数据采集模块采集。工件借助强磁铁固定在台钳夹具上，悬伸长度为 20 mm。对于不同实验组，'
     '强磁铁的位置保持恒定以减小误差。加速度计通过强磁底座安装在蒙皮零件的几何中心。边缘铣削采用三刃整体硬质合金'
     '平底立铣刀，螺旋角 38°，直径 10 mm。', indent_first=0.5)
para('采用符合 ISO 21920-2:2021 的算术平均粗糙度 Ra 作为评价表面粗糙度的定量指标。表面粗糙度使用接触式表面'
     '粗糙度仪（TR200）测量。为避免加工毛刺的影响，测量沿侧面中线、垂直于厚度方向进行。粗糙度测量装置如图 7(c) '
     '所示；测量方法遵循国际标准 ISO 21920-3:2021。', indent_first=0.5)
fig(7, '蒙皮零件的实际铣削实验。(a) 机器人铣削系统。(b) 振动信号采集。(c) 粗糙度测量。', 6)

para('2) 机器人铣削数据集构建：在开展正式实验之前，先在图 7 所示的硬件平台上进行预研。预实验中实施多参数组合：'
     '主轴转速取三个水平（4000 r/min、6000 r/min、8000 r/min），进给速度分为从 1 mm/s 到 8 mm/s 的八个梯度'
     '（间隔 1 mm/s），径向切深（RDC）固定为 1 mm。轴向切深（ADC）受蒙皮零件几何形状限制固定为 2.0 mm。'
     '为提高数据可靠性，沿蒙皮边缘区域选取 10 个离散测点，并计算每种参数组合下表面粗糙度的算术平均值。',
     indent_first=0.5)
para('如表 I 所示，通过在每齿进给量与平均表面粗糙度之间建立 Pearson 相关模型，可观察到显著的正相关'
     '（双尾检验显著性水平 p = 0.003），拒绝原假设的概率低于 0.3%。分析表明，提高主轴转速可有效降低每齿进给量，'
     '从而改善加工表面质量。基于此结论，正式实验阶段将主轴转速固定为 8000 r/min。如表 II 所示，进给速度与 RDC '
     '被选为构建机器人铣削数据集的自变量，ADC 固定为 2.0 mm（受蒙皮零件几何形状限制）。采用全因子试验设计来确定'
     '进给速度与 RDC 的水平。', indent_first=0.5)

table_caption('表 I　Pearson 相关性分析')
make_table([
    ['参数', 'Pearson R', 'P 值', '置信区间', '样本数'],
    ['数值', '0.744', '0.003', '[0.360, 0.933]', '24'],
], widths=[1.9, 1.6, 1.2, 2.4, 1.4])

para('为验证数据集设计的合理性，开展了颤振稳定性分析[38]。在加工轨迹的起点、中点与终点选取三个机器人位姿，'
     '其对应的稳定性叶瓣图如图 8(a)–(c) 所示。横轴表示主轴转速，纵轴表示 ADC。不同 RDC 对应的稳定性叶瓣以'
     '不同颜色的曲线区分，图例中给出了明确标注。实际实验参数以蓝色标记表示，因为主轴转速与 ADC 分别固定为 '
     '8000 r/min 与 2.0 mm。总之，数据集中的所有加工参数均位于稳定叶瓣之内，从而保证了实际实验过程中的加工稳定性。',
     indent_first=0.5)
fig(8, '机器人加工系统在不同位姿下的稳定性叶瓣图。(a) 加工轨迹起点。(b) 加工轨迹中点。(c) 加工轨迹终点。', 7)
para('参数范围覆盖了航空蒙皮铣削应用中常见的典型取值。针对每一种加工参数，沿蒙皮工件边缘测量 60 个离散点的'
     '表面粗糙度并采集相应的振动信号，构建包含 3600 个样本、3 个通道的振动—粗糙度数据集。表面粗糙度值依据 '
     'ISO 3:1973 的 R10 优先数系并结合 ISO 21920 中定义的表面粗糙度分级，被划分为五个离散类别，如表 III 所示。',
     indent_first=0.5)

# 表 II：60 组全因子
table_caption('表 II　实际铣削实验的加工参数')
hdr = ['No', '进给 mm/s', 'RDC mm'] * 3
data2 = [hdr]
combos = []
for rdc in [0.5, 1.0, 1.5, 2.0, 2.5, 3.0]:
    for feed in range(1, 11):
        combos.append((feed, rdc))
for i in range(20):
    row = []
    for b in range(3):
        idx = b * 20 + i
        f, r = combos[idx]
        row += [str(idx + 1), str(f), ('%.1f' % r)]
    data2.append(row)
make_table(data2, widths=[0.75, 1.15, 0.85] * 3, bold_rows=(0,))

# 表 III
table_caption('表 III　表面粗糙度分类')
make_table([
    ['Ra (μm)', '[0.8, 1.6)', '[1.6, 2.0)', '[2.0, 2.5)', '[2.5, 3.2)', '[3.2, 6.3)'],
    ['', '类别 1', '类别 2', '类别 3', '类别 4', '类别 5'],
], widths=[1.55, 1.35, 1.3, 1.3, 1.4, 1.6], first_col_left=True)

h_sub('B. 表面粗糙度预测模型的训练')
para('1) 振动信号校正：通过冲击力锤激励的模态锤击试验，可以确定蒙皮零件边缘不同位置的 FRF。图 9 给出了模态'
     '锤击试验的装置配置。实验采用 KSI-728A002 中型冲击力锤，灵敏度 2.52 mV/N，冲击力量程 0–2 kN，锤头质量 '
     '100 g、直径 5 mm。每个位置敲击三次并对结果取平均。FRF 数据由商业软件 DewesoftX 与 SIRIUS X 数据采集模块'
     '采集。', indent_first=0.5)
para('在铣削过程中，由于铣刀的切削作用，蒙皮零件上形成一个沿进给方向推进的缺口。刀具对工件的切削过程就发生在'
     '该缺口处，从而对工件施加铣削力。如图 9(b) 所示，在蒙皮边缘加工出专用缺口，以便在该位置用冲击力锤实现三轴'
     '（X、Y、Z 方向）激励。这一配置能更好地模拟实际铣削工况。加速度计安装在蒙皮零件的几何中心，与实际加工实验'
     '一致。采用均匀离散方法，将蒙皮边缘沿进给方向等分为 11 个点，相邻点间距 20 mm。', indent_first=0.5)
fig(9, '模态锤击试验装置。(a) 模态试验实验。(b) 蒙皮边缘的离散化。', 8)
para('利用实测的工件 FRF，各方向上相邻点的首阶模态频率变化率如图 10(a) 所示。在蒙皮零件的端部与中部区域，'
     '由于距离磁铁较远，模态频率变化率较高；相反，在磁铁附近（如第 3 号与第 8 号点附近），稳定性增强，'
     '变化率较低。蒙皮表现出足够的连续性，即使采用简单的均匀离散方法，任意相邻点之间的模态频率变化率仍低于 0.1，'
     '满足 ε 取 0.1 时的式 (11)。在 0.1 的允许容差内，离散 FRF 可近似代表蒙皮的全边缘 FRF。', indent_first=0.5)
para('为验证振动信号校正模型的有效性，选取了图 10(b) 所示的一段振动信号，图中标注了工艺参数，对应第 8 号点与'
     '第 10 号点之间的加工区段。随着刀具逐渐远离加速度计，振动幅值显著减小。如图 10(c) 所示，校正后不同位置的'
     '振动幅值趋于一致，说明校正有效。', indent_first=0.5)
fig(10, '振动信号校正。(a) 相邻点间首阶模态频率的变化率。(b) 采集的振动加速度信号。(c) 补偿后的振动力信号。', 8)

para('2) 相空间重构与可预测性证明：预设的最大延迟时间 τ_max 与最大嵌入维数 m_max 分别设为 30 与 50。依据式 (21)，'
     '遍历 τ 的取值范围并计算 X、Y、Z 三通道信号的互信息，计算结果如图 11(a) 所示。曲线先下降后上升，在 τ = 5 '
     '附近出现局部极小。依据式 (23)–(28)，计算 X、Y、Z 通道信号的距离比 E_1，如图 11(b) 所示。E_1 随 m 增大而'
     '增大，但增长率逐渐放缓并趋于不变。取 δ 为 0.005，计算导数 dE_1/dm 如图 11(c) 所示。可以清楚看到，当嵌入'
     '维数 m 大于 15 时，E_1 基本保持不变，且 dE_1/dm 小于 δ，满足要求。', indent_first=0.5)
para('从每种加工参数组合中选取多组信号，分别按式 (22) 与式 (30) 计算三个通道的最优延迟时间 τ_s 与最优嵌入维数 '
     'm_s。这些计算结果的均值与标准差如图 11(d) 所示。可以观察到，三个通道数据的 τ_s 与 m_s 相近，表明三个通道'
     '表现出相似的非线性特征。因此，将最优延迟时间 τ_s 设为 5，最优嵌入维数 m_s 设为 15。', indent_first=0.5)
fig(11, '关键参数的选择。(a) 互信息。(b) 距离比。(c) 距离比的导数。(d) 延迟时间 τ_s 与嵌入维数 m。', 8)
para('从每种加工参数组合中选取多组信号，计算 X、Y、Z 通道信号的最大 Lyapunov 指数（LLE）与 Kolmogorov 熵（KE），'
     '二者的均值与标准差如图 12 所示。LLE 表示非线性系统相空间轨迹随时间收敛或发散的指数平均速率。三个通道'
     '信号的 LLE 均为负值且接近零，说明机器人铣削系统是渐近稳定的。这种稳定性意味着相空间轨迹最终将收敛到某个'
     '吸引子或不动点。同时，信号的正 KE 表明系统存在一定程度的混沌，但 KE 接近零表明混沌程度很低。', indent_first=0.5)
fig(12, 'X、Y、Z 通道信号的可预测性分析参数。(a) 最大 Lyapunov 指数。(b) Kolmogorov 熵。', 9)
para('综上所述，机器人—刀具—工件系统表现出低混沌特性，使其动力学行为在长时间内可预测。所采集的信号可用于'
     '预测系统未来的行为。信号相空间轨迹形成有序的闭合环，这些轨迹的特征可借助 GASF-CNN 识别，从而实现表面'
     '粗糙度的预测。这验证了所提粗糙度预测模型的可预测性。', indent_first=0.5)

para('3) 粗糙度预测模型的训练：特征提取完成后，最终数据集包含 3600 个样本。训练集与测试集按 8:2 划分，'
     '训练集 2880 个样本，测试集 720 个样本。预测模型在一台配备 Intel i7-9700（3.00 GHz）CPU、64 GB DDR4 内存'
     '与 24 GB 显存的 NVIDIA RTX3090 GPU 的计算机上训练。为提高预测精度，采用贝叶斯优化对初始学习率、学习率'
     '下降因子等超参数进行微调。经过 20 次迭代，训练完成的神经网络在测试集上的分类准确率为 96.7%。',
     indent_first=0.5)

table_caption('表 IV　对比模型的预测结果')
make_table([
    ['对比模型', 'SNR +∞', 'SNR 10', 'SNR 5', 'SNR 0'],
    ['PSR-GASF-CNN（本文模型）', '96.7%', '95.3%', '93.2%', '90.1%'],
    ['PSR-CNN', '88.2%', '85.5%', '82.7%', '79.5%'],
    ['GASF-CNN', '94.6%', '90.4%', '87.9%', '84.7%'],
    ['CNN', '82.1%', '76.9%', '65.3%', '55.2%'],
    ['Zhao 模型', '91.1%', '88.8%', '83.5%', '79.3%'],
    ['Lu 模型', '89.3%', '84.2%', '78.2%', '72.6%'],
    ['He 模型', '91.2%', '87.6%', '80.3%', '75.6%'],
], widths=[3.5, 1.25, 1.25, 1.25, 1.25], bold_rows=(1,))

para('混淆矩阵如图 13(a) 所示。样本主要聚集在 Ra 1.6 μm 至 Ra 3.2 μm 范围内（占数据集的 88%），误分类有限'
     '（错误率 3.3%），说明它们具有相似的特征。小于 Ra 1.6 μm 与大于 Ra 3.2 μm 的样本较少，而这些样本的分类'
     '准确率接近 100%，说明粗糙度预测模型在处理极端类别时表现良好。利用 PCA 算法对训练集的分类性能进行可视化，'
     '如图 13(b) 所示。模型在多数类别上表现良好，样本在二维空间中形成明显的簇。尽管不同簇之间存在一定重叠，'
     '但仍可观察到较为清晰的边界。这与实际观察一致：粗糙度值略有差异的样本表现出相似的振动信号特征。类别 1 与'
     '类别 2 样本之间区分度有限，源于铣削工艺的约束：Ra 1.6 以下的大多数样本都紧邻 Ra 1.6，只有极少数样本低于 '
     'Ra 1.3。', indent_first=0.5)
fig(13, '粗糙度预测模型的训练结果。(a) 混淆矩阵图。(b) 训练集的主成分分析（PCA）可视化。', 9)
para('为验证模型的鲁棒性，在实验数据中引入高斯噪声，并用信噪比（SNR）表示噪声强度：', indent_first=0.5)
eq('SNR = 10 lg (P_signal / P_noise)', '34')
para('式中 P_signal 与 P_noise 分别表示信号与噪声的功率。SNR 越低，噪声水平越高。', indent_first=0.5)
para('针对本文所提的 PSR-GASF-CNN 模型，我们开展了系统的消融研究，涉及三种架构变体：PSR-CNN、GASF-CNN 与 '
     'CNN 基线；并纳入三种 SOTA 预测模型（Zhao 模型[29]、Lu 模型[31] 与 He 模型[33]）进行对比分析。全部七个'
     '模型均采用相同方法在本文构建的数据集上训练，其预测性能在表 IV 中进行定量比较。PSR-GASF-CNN 模型的预测'
     '准确率超过 96.7%，且鲁棒性强、对噪声干扰的敏感性很低。即使在 SNR 为 0 dB 的强噪声环境下，模型仍保持 90% '
     '的分类准确率。模型的高精度与强鲁棒性可由以下两方面解释。', indent_first=0.5)
numbered('1)', '基于 Takens 嵌入定理，相空间与机器人铣削系统微分同胚。这意味着机器人铣削系统的动力学特征可通过'
              '振动信号重构出来，便于 CNN 提取高层特征，从而提高预测模型的精度与鲁棒性。')
numbered('2)', 'GASF 方法不仅捕捉相邻元素之间的局部变化，还编码了时间序列中任意两个元素之间的全局时序关系。'
              '这种高维嵌入提升了分类精度，并对工程应用中常见的随机高斯噪声表现出很强的免疫力。')

h_sub('C. 在线监测试验')
para('为验证所提粗糙度监测模型的适用性，利用第 IV-A 节所述的机器人加工系统开展了实际在线监测试验。采用变进给'
     '速度的加工策略可以产生不同的表面粗糙度水平，有助于分析模型对粗糙度变化的敏感性。RDC 与 ADC 固定为 2.0 mm。'
     '该策略中，A–B–C 段的进给速度设为 3 mm/s，C–D–E 段设为 1 mm/s，如图 14(a) 所示。由于结构不稳定，边缘区域'
     '（即点 A 之前与点 E 之后）被排除在分析之外。', indent_first=0.5)
fig(14, '实际在线监测试验。(a) 实验装置。(b) 在线粗糙度监测程序。', 10)
para('应用上一节训练得到的预测模型，即可对蒙皮零件的表面粗糙度进行在线监测。预测结果如图 14(b) 所示，'
     '图中标注了点 A 至点 E。在预测结果中，B–C 段位于两块磁铁之间，夹持刚度较高，因此粗糙度低于 A–B 段。'
     '在 C 点，模型准确响应了进给速度突变所引起的粗糙度变化。D 点更靠近磁铁，夹持刚度最佳，因此该点表现出最好的'
     '加工质量。', indent_first=0.5)
para('在每个特征点附近选取十个离散点并测量其表面粗糙度。结果表明，模型在低速段具有较高的预测精度（接近 100%），'
     '而在变速加工段精度较低（仅为 60%）。整体预测精度为 90%，低于测试集的结果。这一差异可归因于实际加工过程'
     '固有的随机性与不稳定性，它影响了模型的预测性能。总体而言，所提模型能够实现实时的表面粗糙度监测，'
     '及时发现航空蒙皮边缘机器人铣削的异常。', indent_first=0.5)

# =====================================================================
# V. 结论
# =====================================================================
h_section('V.　结　论')
para('本文提出了一种基于 PSR-GASF-CNN 的粗糙度监测模型，用于航空蒙皮边缘铣削中表面粗糙度的在线评估。该模型'
     '采用位置相关的 FRF 对振动信号进行校正，有效抑制了刀具沿工件进给运动的影响。该方法通过延迟嵌入将振动信号'
     '重构到相空间中；与 CNN 分类器结合时，通过 GASF 编码相空间轨迹的全局关系，在测试集上取得了超过 96.7% 的'
     '准确率。模型表现出很强的鲁棒性，在高噪声环境下仍保持 90% 以上的准确率。此外，通过 LLE 与 KE 分析验证了'
     '振动信号的可预测性。当前模型面向矩形平面工件，未来研究将把该框架推广至自由曲面工件。总体而言，'
     '该方案成功满足了航空蒙皮在线粗糙度监测的需求，并为工业实施提供了实质性指导。', indent_first=0.5)

# =====================================================================
# 参考文献
# =====================================================================
h_section('参考文献（References）')
refs = [
 '[1] C. Jiang et al., “A novel dual-robot accurate calibration method using convex optimization and lie derivative,” IEEE Trans. Robot., vol. 40, pp. 960–977, Jan. 2024, doi: 10.1109/TRO.2023.3344025.',
 '[2] G. Wang, W. Li, C. Jiang, and H. Ding, “Machining allowance calculation for robotic edge milling an aircraft skin considering the deformation of assembly process,” IEEE/ASME Trans. Mechatron., vol. 27, no. 5, pp. 3350–3361, Oct. 2022, doi: 10.1109/TMECH.2021.3131309.',
 '[3] Q. Peng, W. Li, C. Jiang, Y. Cheng, D. Wang, and W. Xu, “A novel path error compensation method for robotic milling by using hybrid temporal network,” IEEE/ASME Trans. Mechatron., doi: 10.1109/TMECH.2024.3479410.',
 '[4] C. Moeller et al., “Real time pose control of an industrial robotic system for machining of large scale components in aerospace industry using laser tracker system,” SAE Int. J. Aerosp., vol. 10, no. 2, pp. 100–108, Sep. 2017, doi: 10.4271/2017-01-2165.',
 '[5] A. Verl et al., “Robots in machining,” CIRP Ann., vol. 68, pp. 799–822, Jun. 2019, doi: 10.1016/j.cirp.2019.05.009.',
 '[6] S. C. Du, D. L. Huang, and H. Wang, “An adaptive support vector machine-based workpiece surface classification system using high-definition metrology,” IEEE Trans. Instrum. Meas., vol. 64, no. 10, pp. 2590–2604, Oct. 2015, doi: 10.1109/TIM.2015.2418684.',
 '[7] P. Xu et al., “Influence of process parameters and robot postures on surface quality in robotic machining,” Int. J. Adv. Manuf. Technol., vol. 124, no. 7/8, pp. 2545–2561, Feb. 2023, doi: 10.1007/s00170-022-10640-2.',
 '[8] Y. Hao and Y. Liu, “Analysis of milling surface roughness prediction for thin-walled parts with curved surface,” Int. J. Adv. Manuf. Technol., vol. 93, no. 5–8, pp. 2289–2297, Nov. 2017, doi: 10.1007/s00170-017-0615-4.',
 '[9] D. Vakondios et al., “Influence of milling strategy on the surface roughness in ball end milling of the aluminum alloy Al7075-T6,” Measurement, vol. 45, no. 6, pp. 1480–1488, Jul. 2012, doi: 10.1016/j.measurement.2012.03.001.',
 '[10] W. Wang et al., “A state-of-the-art review on robotic milling of complex parts with high efficiency and precision,” Robot. Comput.-Integr. Manuf., vol. 79, Aug. 2022, Art. no. 102436, doi: 10.1016/j.rcim.2022.102436.',
 '[11] G. Onwuka and K. Abou-El-Hossein, “Surface roughness in ultra-high precision grinding of BK7,” Procedia CIRP, vol. 45, pp. 143–146, Jun. 2016, doi: 10.1016/j.procir.2016.03.023.',
 '[12] A. Sio-Sever et al., “Non-invasive estimation of machining parameters during end-milling operations based on acoustic emission,” Sensors, vol. 20, no. 18, Sep. 2020, Art. no. 5326, doi: 10.3390/s20185326.',
 '[13] H. Shi, J. Shi, Y. Yang, S. Tang, C. Su, and L. Liu, “Spatial vibration and displacement measurement through single-point continuous and interval scanning laser Doppler vibrometer,” IEEE Trans. Instrum. Meas., vol. 72, Jan. 2023, Art. no. 7001011, doi: 10.1109/TIM.2022.3225055.',
 '[14] Y. Qu, T. Wang, and Z. Zhu, “Vision-aided laser Doppler vibrometry for remote automatic voice detection,” IEEE/ASME Trans. Mechatron., vol. 16, no. 6, pp. 1110–1119, Dec. 2011, doi: 10.1109/TMECH.2010.2077678.',
 '[15] L. Su, W. Xu, Y.-P. Wang, and W.-L. Li, “Robust point cloud registration in robotic inspection with locally consistent Gaussian mixture model,” IEEE Trans. Instrum. Meas., vol. 73, Oct. 2024, Art. no. 5038712, doi: 10.1109/TIM.2024.3480198.',
 '[16] C. Warren et al., “Comparison of FRF measurements and mode shapes determined using optically image based, laser, and accelerometer measurements,” Mech. Syst. Signal Process., vol. 25, no. 6, pp. 2191–2202, Aug. 2011, doi: 10.1016/j.ymssp.2011.01.018.',
 '[17] W. Sun, M. Luo, and D. Zhang, “Machining vibration monitoring based on dynamic clamping force measuring in thin-walled components milling,” Int. J. Adv. Manuf. Technol., vol. 107, no. 5/6, pp. 2211–2226, Mar. 2020, doi: 10.1007/s00170-020-05153-9.',
 '[18] B. Cardoz et al., “Random forests-based classification of tool wear using vibration signals and wear area estimation from tool image data,” Int. J. Adv. Manuf. Technol., vol. 126, no. 7/8, pp. 3069–3081, Jun. 2023, doi: 10.1007/s00170-023-11173-y.',
 '[19] Q. Zhu, B. Sun, Y. Zhou, W. Sun, and J. Xiang, “Sample augmentation for intelligent milling tool wear condition monitoring using numerical simulation and generative adversarial network,” IEEE Trans. Instrum. Meas., vol. 70, May 2021, Art. no. 3516610, doi: 10.1109/TIM.2021.3077995.',
 '[20] D. Liu et al., “Position-oriented process monitoring in milling of thin-walled parts,” J. Manuf. Syst., vol. 60, pp. 360–372, Jul. 2021, doi: 10.1016/j.jmsy.2021.06.010.',
 '[21] Z. Yao et al., “Position dependent vibration evaluation in milling of thin-walled part based on single-point monitoring,” Measurement, vol. 171, Feb. 2021, Art. no. 108810, doi: 10.1016/j.measurement.2020.108810.',
 '[22] Z. Yao et al., “Position-dependent milling process monitoring and surface roughness prediction for complex thin-walled blade component,” Mech. Syst. Signal Process., vol. 198, Sep. 2023, Art. no. 110439, doi: 10.1016/j.ymssp.2023.110439.',
 '[23] J. Burriel-Valencia, R. Puche-Panadero, J. Martinez-Roman, A. Sapena-Bano, and M. Pineda-Sanchez, “Short-frequency Fourier transform for fault diagnosis of induction machines working in transient regime,” IEEE Trans. Instrum. Meas., vol. 66, no. 3, pp. 432–440, Mar. 2017, doi: 10.1109/TIM.2016.2647458.',
 '[24] Y. Jin, C. Shan, Y. Wu, Y. Xia, Y. Zhang, and L. Zeng, “Fault diagnosis of hydraulic seal wear and internal leakage using wavelets and wavelet neural network,” IEEE Trans. Instrum. Meas., vol. 68, no. 4, pp. 1026–1034, Apr. 2019, doi: 10.1109/TIM.2018.2863418.',
 '[25] A. Soualhi, K. Medjaher, and N. Zerhouni, “Bearing health monitoring based on Hilbert–Huang transform, support vector machine, and regression,” IEEE Trans. Instrum. Meas., vol. 64, no. 1, pp. 52–62, Jan. 2015, doi: 10.1109/TIM.2014.2330494.',
 '[26] S. R. Fahim et al., “Self-attention convolutional neural network with time series imaging-based feature extraction for transmission line fault detection and classification,” Elect. Power Syst. Res., vol. 187, Oct. 2020, Art. no. 106437, doi: 10.1016/j.epsr.2020.106437.',
 '[27] W. X. Wang, Y. C. Lai, and C. Grebogi, “Data based identification and prediction of nonlinear and complex dynamical systems,” Phys. Rep., vol. 644, pp. 1–76, Jul. 2016, doi: 10.1016/j.physrep.2016.06.004.',
 '[28] H. Ding, R. X. Gao, A. J. Isaksson, R. G. Landers, T. Parisini, and Y. Yuan, “State of AI-based monitoring in smart manufacturing and introduction to focused section,” IEEE/ASME Trans. Mechatron., vol. 25, no. 5, pp. 2143–2154, Oct. 2020, doi: 10.1109/TMECH.2020.3022983.',
 '[29] X. Zhao, H. Lu, W. Yu, B. Tao, and H. Ding, “Robotic grinding process monitoring by vibration signal based on LSTM method,” IEEE Trans. Instrum. Meas., vol. 71, Oct. 2022, Art. no. 400911, doi: 10.1109/TIM.2022.3210978.',
 '[30] L. Cao, X. M. Zhang, T. Huang, and H. Ding, “Online monitoring machining errors of thin-walled workpiece: A knowledge embedded sparse Bayesian regression approach,” IEEE/ASME Trans. Mechatron., vol. 24, no. 3, pp. 1259–1270, Jun. 2019, doi: 10.1109/TMECH.2019.2912195.',
 '[31] H. Lu, X. Zhao, B. Tao, and Z. Yi, “Online process monitoring based on vibration-surface quality map for robotic grinding,” IEEE/ASME Trans. Mechatron., vol. 25, no. 6, pp. 2882–2892, Dec. 2020, doi: 10.1109/TMECH.2020.2996939.',
 '[32] S. Li, X. Lin, H. Shi, Y. Shi, and K. Zhu, “Physics-guided deep learning method for tool condition monitoring in smart machining system,” IEEE/ASME Trans. Mechatron., vol. 29, no. 3, pp. 2327–2337, Jun. 2024, doi: 10.1109/TMECH.2023.3311435.',
 '[33] W. He et al., “Unified diagnostic and matching framework of fault and quality for robotic grinding system,” IEEE Trans. Instrum. Meas., vol. 73, Mar. 2024, Art. no. 3513612, doi: 10.1109/TIM.2024.3376009.',
 '[34] D. M. Shi, T. Huang, X. M. Zhang, and H. Ding, “Online monitoring dynamic characteristics in thin-walled structure milling: A physics-constrained Bayesian updating approach,” IEEE/ASME Trans. Mechatron., vol. 28, no. 6, pp. 3422–3433, Dec. 2023, doi: 10.1109/TMECH.2023.3268180.',
 '[35] Y. Zhao, Z. Huang, and X. Wang, “A review of specific emitter identification based on phase space reconstruction,” J. Radars, vol. 12, no. 4, pp. 713–737, Aug. 2023, doi: 10.12000/JR23057.',
 '[36] A. M. Fraser and H. L. Swinney, “Independent coordinates for strange attractors from mutual information,” Phys. Rev. A, vol. 33, no. 2, pp. 1134–1140, Feb. 1986, doi: 10.1103/PhysRevA.33.1134.',
 '[37] L. Cao, “Practical method for determining the minimum embedding dimension of a scalar time series,” Physica D, Nonlinear Phenomena, vol. 110, no. 1/2, pp. 43–50, Dec. 1997, doi: 10.1016/S0167-2789(97)00118-8.',
 '[38] Y. Ding et al., “A full-discretization method for prediction of milling stability,” Int. J. Mach. Tools Manuf., vol. 50, pp. 502–509, Feb. 2010, doi: 10.1016/j.ijmachtools.2010.01.003.',
]
for r in refs:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(1.5)
    p.paragraph_format.left_indent = Cm(0.6)
    p.paragraph_format.first_line_indent = Cm(-0.6)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    rr = p.add_run(r)
    run_font(rr, 7.5)

# ---------- 作者简介 ----------
h_section('作者简介（Author Biographies）')
bios = [
 ('张涵宇（Han-yu Zhang）', '2022 年获华中科技大学机械设计制造及其自动化工学学士学位，现于该校攻读机械工程博士学位。'
  '研究方向为机器人加工与测量。'),
 ('徐蔚（Wei Xu, Member, IEEE）', '2013 年、2016 年分别获中国矿业大学（北京）流体机械与工程专业工学学士与工程硕士学位，'
  '2021 年获德国莱布尼茨汉诺威大学测量工程博士学位。现任华中科技大学讲师。研究方向为机器人测量与制造。'),
 ('彭庆宇（Qing-yu Peng）', '2020 年获武汉理工大学轮机工程工学学士学位，现于华中科技大学攻读机械工程博士学位。'
  '研究方向为机器人加工与机器人测量。'),
 ('王东方（Dong-fang Wang）', '2020 年获武汉理工大学机械设计制造及其自动化工学学士学位，现于华中科技大学攻读机械工程'
  '博士学位。研究方向为机器人加工。'),
 ('蒋诚（Cheng Jiang）', '2016 年、2024 年分别获华中科技大学机械设计制造及其自动化工学学士学位与机械电子工程'
  '博士学位。现为华中科技大学与吉利汽车研究院联合培养项目博士后。研究方向为机器人自主执行与操作。'),
 ('李文龙（Wen-long Li, Member, IEEE）', '2004 年获西安交通大学机械工程及自动化工学学士学位，2010 年获华中科技大学'
  '机械电子工程博士学位。现任华中科技大学教授。研究方向为机器人加工与三维光学测量。'),
]
for name, txt in bios:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r = p.add_run(name + '：')
    run_font(r, 8, bold=True)
    r = p.add_run(txt)
    run_font(r, 8)

p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(8)
r = p.add_run('说明：本译文仅供个人学习与研究使用，不用于再出版或再分发。原文版权归 © 2025 IEEE 所有；'
              '原文图表未复制，正文中的图表以其题注位置标示，阅读时请对照原文 PDF。公式按原文编号 (1)–(34) 保留。')
run_font(r, 7, italic=True, color=RGBColor(0x66, 0x66, 0x66))

doc.save(OUT)
print('saved', OUT, os.path.getsize(OUT))
