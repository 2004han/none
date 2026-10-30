# -*- coding: utf-8 -*-
"""生成《文献精读报告》DOCX。"""
import os
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT = r"E:\试切实验\精读报告_PSR-GASF-CNN蒙皮边缘粗糙度在线监测.docx"
ACCENT = RGBColor(0x1F, 0x3B, 0x63)
GREY = RGBColor(0x66, 0x66, 0x66)

doc = Document()
st = doc.styles['Normal']
st.font.name = 'Times New Roman'
st.font.size = Pt(10.5)
st.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
st.paragraph_format.space_before = Pt(0)
st.paragraph_format.space_after = Pt(3)
st.paragraph_format.line_spacing = 1.28

sec = doc.sections[0]
sec.page_width = Cm(21.0)
sec.page_height = Cm(29.7)
sec.left_margin = Cm(2.2)
sec.right_margin = Cm(2.2)
sec.top_margin = Cm(2.2)
sec.bottom_margin = Cm(2.0)

USABLE = 16.6  # cm


def rfont(run, size=None, bold=False, italic=False, color=None, ea='宋体'):
    run.font.name = 'Times New Roman'
    rpr = run._element.get_or_add_rPr()
    rf = rpr.find(qn('w:rFonts'))
    if rf is None:
        rf = OxmlElement('w:rFonts')
        rpr.append(rf)
    rf.set(qn('w:ascii'), 'Times New Roman')
    rf.set(qn('w:hAnsi'), 'Times New Roman')
    rf.set(qn('w:eastAsia'), ea)
    if size is not None:
        run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    if color is not None:
        run.font.color.rgb = color


def shade(cell, hexcolor):
    tcPr = cell._tc.get_or_add_tcPr()
    sh = OxmlElement('w:shd')
    sh.set(qn('w:val'), 'clear')
    sh.set(qn('w:color'), 'auto')
    sh.set(qn('w:fill'), hexcolor)
    tcPr.append(sh)


def tborder(t, color='BFBFBF', sz=4):
    tblPr = t._tbl.tblPr
    for old in tblPr.findall(qn('w:tblBorders')):
        tblPr.remove(old)
    b = OxmlElement('w:tblBorders')
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        e = OxmlElement('w:' + edge)
        e.set(qn('w:val'), 'single')
        e.set(qn('w:sz'), str(sz))
        e.set(qn('w:space'), '0')
        e.set(qn('w:color'), color)
        b.append(e)
    tblPr.append(b)


def h1(text, num=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.keep_with_next = True
    if num:
        r = p.add_run(num + '　')
        rfont(r, 14, bold=True, color=ACCENT, ea='黑体')
    r = p.add_run(text)
    rfont(r, 14, bold=True, color=ACCENT, ea='黑体')
    # 下边框
    pPr = p._p.get_or_add_pPr()
    pbdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:space'), '2')
    bottom.set(qn('w:color'), '1F3B63')
    pbdr.append(bottom)
    pPr.append(pbdr)
    return p


def h2(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(9)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    rfont(r, 11.5, bold=True, color=ACCENT, ea='黑体')
    return p


def h3(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    rfont(r, 10.5, bold=True)
    return p


def body(text, indent=True, size=10.5, space_after=4):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_after = Pt(space_after)
    if indent:
        p.paragraph_format.first_line_indent = Cm(0.74)
    r = p.add_run(text)
    rfont(r, size)
    return p


def rich(parts, indent=True, size=10.5, space_after=4):
    """parts: list of (text, bold)"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_after = Pt(space_after)
    if indent:
        p.paragraph_format.first_line_indent = Cm(0.74)
    for text, bold in parts:
        r = p.add_run(text)
        rfont(r, size, bold=bold)
    return p


def bullet(text, size=10.5, level=0):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_after = Pt(2.5)
    p.paragraph_format.left_indent = Cm(0.74 + 0.6 * level)
    p.paragraph_format.first_line_indent = Cm(-0.4)
    r = p.add_run('▪ ' if level == 0 else '· ')
    rfont(r, size, color=ACCENT)
    r = p.add_run(text)
    rfont(r, size)
    return p


def note(text, size=9.5):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.left_indent = Cm(0.74)
    p.paragraph_format.right_indent = Cm(0.74)
    r = p.add_run(text)
    rfont(r, size, italic=True, color=GREY)
    return p


def caption(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    rfont(r, 9, bold=True, color=GREY, ea='黑体')
    return p


def make_table(data, widths, header=True, font=9, align_center=None,
               header_fill='1F3B63', zebra=True, first_bold=False):
    t = doc.add_table(rows=len(data), cols=len(data[0]))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    tborder(t)
    for i, row in enumerate(data):
        for j, val in enumerate(row):
            cell = t.cell(i, j)
            cell.width = Cm(widths[j])
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.line_spacing = 1.05
            if align_center is not None:
                c = align_center[j] if isinstance(align_center, (list, tuple)) else align_center
            else:
                c = (j > 0)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c else WD_ALIGN_PARAGRAPH.LEFT
            r = p.add_run(str(val))
            if i == 0 and header:
                rfont(r, font, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF), ea='黑体')
                shade(cell, header_fill)
            else:
                rfont(r, font, bold=(first_bold and j == 0))
                if zebra and i % 2 == 0:
                    shade(cell, 'F2F5FA')
    # 固定布局
    tbl = t._tbl
    tblPr = tbl.tblPr
    tblW = OxmlElement('w:tblW')
    tblW.set(qn('w:w'), str(int(sum(widths) * 567)))
    tblW.set(qn('w:type'), 'dxa')
    tblPr.append(tblW)
    lay = OxmlElement('w:tblLayout')
    lay.set(qn('w:type'), 'fixed')
    tblPr.append(lay)
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
            row.cells[j].width = Cm(w)
        trPr = row._tr.get_or_add_trPr()
        trPr.append(OxmlElement('w:cantSplit'))
    return t


def spacer(pt=4):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = Pt(pt)
    return p


# ============================ 封面标题 ============================
p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(6)
p.paragraph_format.space_after = Pt(2)
r = p.add_run('文 献 精 读 报 告')
rfont(r, 12, bold=True, color=GREY, ea='黑体')
p.alignment = WD_ALIGN_PARAGRAPH.CENTER

p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(6)
r = p.add_run('基于相空间重构的航空蒙皮边缘机器人铣削\n表面粗糙度在线监测')
rfont(r, 19, bold=True, color=ACCENT, ea='黑体')
p.alignment = WD_ALIGN_PARAGRAPH.CENTER

p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(12)
r = p.add_run('Phase Space Reconstruction-Based Online Roughness Monitoring in Robotic Milling of Aircraft Skin Edge')
rfont(r, 10, italic=True, color=GREY)
p.alignment = WD_ALIGN_PARAGRAPH.CENTER

make_table([
    ['项目', '内容'],
    ['作者', '张涵宇、徐蔚（通信作者）、彭庆宇、王东方、蒋诚、李文龙'],
    ['单位', '华中科技大学 智能制造装备与技术全国重点实验室'],
    ['期刊', 'IEEE/ASME Transactions on Mechatronics（TMECH）'],
    ['卷期页码', 'Vol. 31, No. 2, April 2026, pp. 1741–1752'],
    ['DOI', '10.1109/TMECH.2025.3603285'],
    ['时间线', '投稿 2025-01-04｜修回 2025-03-26 / 2025-05-26｜录用 2025-08-20｜上线 2025-10-03'],
    ['基金', '国家自然科学基金 52205524、52575573、52188102；中国博士后科学基金 2022M721247；IMETKF2025001'],
    ['文献类型', '工程方法类（信号处理 + 深度学习 + 机器人铣削实验验证）'],
    ['阅读深度', '全文精读（12 页，4 表 14 图，表格以矢量图形式嵌入，已逐格核读）'],
], widths=[2.6, 14.0], font=9, align_center=[False, False])

spacer(6)
caption('总体评级（5 分制）')
make_table([
    ['维度', '评分', '一句话理由'],
    ['问题价值', '★★★★☆', '蒙皮边缘在线粗糙度监测是真实且紧迫的产线需求，位置效应是公认痛点'],
    ['方法创新', '★★★☆☆', 'FRF 反卷积校正 + 相空间成像的组合有新意，但 PSR、GASF、CNN 均为成熟工具'],
    ['实验扎实度', '★★★★☆', '自适应工况、稳定性叶瓣、ISO 标准、在线验证齐备，工程可信度高'],
    ['统计严谨性', '★★☆☆☆', '随机划分疑似数据泄漏，只报总体准确率，缺分组验证与置信度'],
    ['可复现性', '★★☆☆☆', '采样率、窗长、CNN 结构、通道融合方式全部缺失'],
    ['写作质量', '★★★★☆', 'gap—贡献一一对应，主动补齐审稿人会问的实验'],
], widths=[3.0, 2.2, 11.4], font=9, align_center=[False, True, False])

# ============================ 1 ============================
h1('一句话概括与核心结论', '一')

body('这篇文章解决的问题可以浓缩成一句：加速度计固定在工件中心、刀具沿蒙皮边缘走刀，'
     '此过程中"信号变了"到底是位置变了还是粗糙度变了？作者的做法是用离线实测的分段频响函数（FRF）做反卷积，'
     '把固定点振动还原成"当前切削点的激振力"，再经相空间重构（PSR）与格拉姆角求和场（GASF）成像，'
     '最后用 CNN 做 5 级粗糙度分类。')

rich([('核心数据：', True),
      ('6061 铝蒙皮 200×200×2 mm，ABB IRB 6660，φ10 三刃铣刀，8000 r/min，60 组全因子（进给 1–10 mm/s × RDC 0.5–3.0 mm），'
       '3600 样本 × 3 通道。测试集准确率 96.7%，SNR 0 dB 下仍有 90.1%，在线变进给实验整体 90%、变速段仅 60%。', False)])

note('读这篇文章最值得记住的一点：作者把"传感器位置固定 vs 加工点移动"这个所有单点监测方法都会遇到的'
     '物理难题，用一条反卷积公式（式 16）正面解决了，而不是靠"多布几个传感器"或"拿位置当输入特征"绕过去。'
     '这是它区别于 Yao 等[21][22]（只修幅值）的实质进步。')

# ============================ 2 ============================
h1('问题链：作者如何一步步立论', '二')

body('全文的论证是一套标准的"缺口—补丁"结构，三条 gap 对应三个贡献，逻辑链清晰，值得借鉴：')
make_table([
    ['环节', '内容', '对应贡献'],
    ['工程需求', '蒙皮边缘粗糙度影响气动、隐身与装配密封，需要在线而非离线打点抽检', '—'],
    ['手段选择', '声发射信噪比低、激光多普勒只能测法向且低频精度差、结构光精度低 → 选接触式三轴加速度计', '—'],
    ['缺口 1', '传感器固定、刀具移动：信号变化混淆了"位置效应"与"粗糙度效应"。'
               '文献[20] 只做位置关联，[21][22] 只修幅值，未补偿相位漂移', '贡献① FRF 校正'],
    ['缺口 2', '现有特征提取（STFT/小波/HHT/GAF）对噪声敏感，缺乏对信号内在非线性的刻画', '贡献② PSR + GASF'],
    ['缺口 3', '机器人铣削振动的"可预测性"缺乏定量证据', '贡献③ LLE + KE 证明'],
], widths=[2.2, 10.6, 3.8], font=9, align_center=[False, False, True])

# ============================ 3 ============================
h1('方法拆解', '三')

h2('3.1　第一步：基于 FRF 的振动信号校正（第 II 节）')

body('作者放弃了理论动力学建模，改以"实测 FRF"充当传递关系的替身，理由是磁铁/真空吸盘夹持不满足理论边界条件。'
     '三个关键推导：')

rich([('① 采样周期内激励点不变。', True),
      ('刀具在一个采样周期内的位移远小于工件尺寸，因此"单激励点—单响应点"成立，可写成三自由度方程 '
       'M q̈ + C q̇ + K q = F(t)，傅里叶变换后得 (−ω²M + jωC + K)q(ω) = F(ω)，'
       '定义位移 FRF 为 H(ω) = (−ω²M + jωC + K)⁻¹。', False)], indent=False)

rich([('② 一个反教科书的实测发现。', True),
      ('理论上 x/y/z 正交时 H 应为对角阵，但作者实测发现交叉 FRF 大到不可忽略——磁力夹持引入的非线性'
       '使 H 不再是对角阵。因此后续不能简单求逆，必须用 Moore–Penrose 伪逆。这是全文可信度最高的细节。', False)], indent=False)

rich([('③ 位置离散化 + 伪逆反解。', True),
      ('把边缘按 Δk 分段，用"相邻点模态频率/阻尼相对变化 ≤ ε（取 0.1）"确定分段；'
       '模态频率取 dHᴵ/dω = 0，阻尼用半功率带宽 ξ_m = |ω_a − ω_b|/ω_m。'
       '最终校正公式为 F(ω) = H_a(ω)⁺ a(ω)，F(t) = F⁻¹{F(ω)}，'
       '即用该刀位对应的加速度 FRF 伪逆，把固定点测到的加速度反解成切削点激振力谱。', False)], indent=False)

note('物理解释：这一步本质是"位置相关的逆滤波"。传递路径效应被除掉，留下的是激励源信息，而激励源才与材料去除、'
     '粗糙度直接相关。这也是全文的支点——若这一步不成立，后面的特征工程就失去了物理依据。')

h2('3.2　第二步：相空间重构（第 III-A 节）')

body('延迟嵌入构造 s_i = [f(i), f(i+τ), …, f(i+(m−1)τ)]ᵀ，得到轨迹集 S。理论依据是 Takens 嵌入定理：'
     '选择合适的 m 与 τ，重构相空间与原动力系统微分同胚（拓扑等价），因此有限个观测量即可重构系统动力学。')
body('两个参数分别用成熟方法确定：延迟时间 τ 用 Fraser–Swinney 互信息最小值（式 21、22）；'
     '嵌入维数 m 用 Cao 法（式 23–30），即最近邻距离比 E₁(m) 的导数降到 δ 以下时停止。'
     '实测 τ_max = 30、m_max = 50，得到 τ_s ≈ 4–5、m_s ≈ 11–12，最终取 τ = 5、m = 15，δ = 0.005。')
note('值得注意的小问题：Fig. 11(d) 显示三通道 m_s 均值约 11.5，作者却取 15，文中未说明"向上取整/留安全余量"'
     '的理由；Fig. 11(a) 的互信息出现了负值（−4 到 1），而真实互信息恒非负，说明直方图估计存在偏差或样本不足，'
     '作者对此没有讨论。')

h2('3.3　第三步：GASF 的"向量版"推广（第 III-B 节）')

body('原始 GASF 只适用于标量序列（用 x/max x 归一化后取 arccos）。本文把标量换成"相空间向量与参考向量的夹角余弦"：'
     's̃_i = cos θ_i = (s_i · s_0)/(|s_i||s_0|)，然后 GASF = [cos(θ_i + θ_j)]，'
     '并用三角恒等式化为矩阵运算 GASF = S̃ᵀS̃ − √(I − (S̃ᵀ)²)·√(I − S̃²) 以避免逐元素三角函数、提升速度。')

rich([('这里有两处必须打折扣：', True),
      ('（a）信息塌缩——s̃_i 只依赖 s_i 与 s_0 的夹角，对 |s_i| 完全不变（尺度不变），'
       '即 m 维向量被压成一个标量序列 {θ_i} 后再做 GASF，"高维嵌入的优势"部分被抵消；'
       '（b）"参考向量可任取、本质相同"不成立——cos(θ_i+θ_j) 明显依赖 s_0 的选择，不同 s_0 得到的是不同图像。'
       '实践上 CNN 能容忍，但作为理论陈述过强。', False)])

h2('3.4　分类器与训练')

body('CNN + 贝叶斯超参优化（初始学习率、学习率下降因子等），20 次迭代。训练环境 i7-9700 / 64 GB / RTX 3090。'
     '注意：网络层级结构、输入图像尺寸、样本窗长、采样率全部未给出，这在复现层面是硬伤（见第七节）。')

# ============================ 4 ============================
h1('实验设计与数据集', '四')

make_table([
    ['项目', '具体内容'],
    ['工件', '6061 铝合金蒙皮 200 × 200 × 2 mm，悬伸 20 mm，台钳 + 强磁铁夹持'],
    ['装备', 'ABB IRB 6660 机器人 + EBLTS TS120F 高速电主轴；KSI-138M100 三轴加速度计（磁座固定在工件几何中心）；'
             'DewesoftX + SIRIUS X 采集'],
    ['刀具', '三刃整体硬质合金平底立铣刀，螺旋角 38°，φ10 mm，边缘铣削'],
    ['粗糙度', 'Ra（ISO 21920-2:2021），TR200 接触式轮廓仪，沿侧面中线、垂直厚度方向测（避开毛刺），'
              '流程遵循 ISO 21920-3:2021'],
    ['预实验', '3 转速（4000/6000/8000 r/min）× 8 进给（1–8 mm/s），RDC 固定 1 mm、ADC 固定 2.0 mm，每组 10 点取均值'],
    ['相关性', 'Pearson R = 0.744，p = 0.003（双尾），95% CI [0.360, 0.933]，n = 24 → 转速提至 8000 r/min'],
    ['正式实验', '全因子 60 组：进给 1–10 mm/s（10 水平）× RDC {0.5, 1.0, 1.5, 2.0, 2.5, 3.0} mm（6 水平），'
                'ADC 固定 2.0 mm；每组沿边缘测 60 点 ⟹ 3600 样本 × 3 通道'],
    ['稳定性', '全离散法在轨迹起点/中点/终点三个位姿画稳定性叶瓣（6 条曲线对应 6 个 RDC），实验点（8000 r/min, 2.0 mm）'
              '落在稳定域内，排除颤振'],
    ['模态测试', '边缘 11 等分（20 mm 间距），力锤 KSI-728A002（2.52 mV/N，0–2 kN，锤头 100 g/φ5 mm），每点敲 3 次取平均；'
                '专门铣出缺口以实现三轴激励'],
], widths=[2.4, 14.2], font=9, align_center=[False, False])

spacer(4)
caption('表 A　粗糙度分级（ISO 3:1973 R10 优先数系 + ISO 21920 分级）')
make_table([
    ['Ra (μm)', '[0.8, 1.6)', '[1.6, 2.0)', '[2.0, 2.5)', '[2.5, 3.2)', '[3.2, 6.3)'],
    ['类别', 'Class 1', 'Class 2', 'Class 3', 'Class 4', 'Class 5'],
], widths=[2.6, 2.8, 2.8, 2.8, 2.8, 2.8], font=9)

note('分级边界采用 R10 优先数系而非等距分箱，是本文工程严谨性的一个细节，评审接受度更高——这一点在自己的'
     '论文里可以直接照抄。同时要注意：88% 的样本落在 Class 2/3/4（Ra 1.6–3.2），极端类样本极少。')

caption('表 B　可预测性论证的关键结果')
make_table([
    ['指标', 'X 通道', 'Y 通道', 'Z 通道', '解读'],
    ['最大 Lyapunov 指数 LLE', '≈ −2.1×10⁻³', '≈ −1.6×10⁻³', '≈ −1.8×10⁻³',
     '均为负且接近 0 → 系统渐近稳定，轨迹收敛到吸引子'],
    ['Kolmogorov 熵 KE', '≈ 4.5×10⁻⁴', '≈ 4.4×10⁻⁴', '≈ 4.5×10⁻⁴',
     '为正但接近 0 → 存在混沌但程度很低'],
], widths=[4.2, 2.8, 2.8, 2.8, 4.0], font=9, align_center=[False, True, True, True, False])

# ============================ 5 ============================
h1('结果与解读', '五')

h2('5.1　训练效果')

body('训练 2880 / 测试 720（8:2），测试集准确率 96.7%。混淆矩阵显示：88% 的样本集中在 Ra 1.6–3.2 μm，'
     '该区间错误率 3.3%；<1.6 与 >3.2 的样本很少但准确率接近 100%。PCA 可视化中多数类成簇，Class 1/2 重叠——'
     '因为 Ra 1.6 以下的样本大多紧贴 1.6，低于 1.3 的极少，属于工艺约束而非模型缺陷。')

h2('5.2　消融与横向对比（Table IV，测试集准确率 %）')

make_table([
    ['模型', 'SNR ∞', 'SNR 10 dB', 'SNR 5 dB', 'SNR 0 dB', '∞→0 dB 跌幅'],
    ['PSR-GASF-CNN（本文）', '96.7', '95.3', '93.2', '90.1', '−6.6'],
    ['GASF-CNN', '94.6', '90.4', '87.9', '84.7', '−9.9'],
    ['He 模型（SDP-CNN）[33]', '91.2', '87.6', '80.3', '75.6', '−15.6'],
    ['Zhao 模型（LSTM）[29]', '91.1', '88.8', '83.5', '79.3', '−11.8'],
    ['Lu 模型 [31]', '89.3', '84.2', '78.2', '72.6', '−16.7'],
    ['PSR-CNN', '88.2', '85.5', '82.7', '79.5', '−8.7'],
    ['CNN 基线', '82.1', '76.9', '65.3', '55.2', '−26.9'],
], widths=[5.2, 2.1, 2.3, 2.3, 2.3, 2.4], font=9, align_center=[False, True, True, True, True, True],
    first_bold=False)

spacer(4)
rich([('论文没明说、但表里写得出来的三件事：', True)])
bullet('GASF 才是主力：82.1 → 94.6（+12.5），而 PSR 单独只贡献 +6.1。'
       'PSR 的真正价值体现在强噪声下——GASF-CNN → PSR-GASF-CNN 在 0 dB 时多 5.4 个百分点，干净条件下只多 2.1。'
       '也就是说，论文把"微分同胚"讲成头号卖点，与数据略有错位。')
bullet('成像化是必要不充分条件：CNN 基线在 0 dB 掉到 55.2%（5 分类），说明原始振动直接喂 CNN 几乎不可行。')
bullet('抗噪性比精度更有说服力：本文从 ∞ 到 0 dB 只掉 6.6 个百分点，Lu 模型掉 16.7、He 模型掉 15.6。'
       '这个"退化斜率"其实比 96.7% 这个绝对数字更能说明方法的稳健性。')

h2('5.3　在线验证（Fig. 14）')

body('变进给策略：A–B–C 段 3 mm/s，C–D–E 段 1 mm/s（RDC = ADC = 2.0 mm），排除首尾不稳定区。'
     '结果符合工艺直觉：B–C 位于两磁铁之间、夹持刚度更高 → Ra 更低；C 点准确响应进给突变；'
     'D 点最靠近磁铁、夹持最优 → 质量最好。')
rich([('精度：', True), ('低速段接近 100%，变速段仅 60%，整体 90%。作者归因于"实际加工的随机性与不稳定性"。', False)])

# ============================ 6 ============================
h1('贡献评估：哪些是真新', '六')

make_table([
    ['声称的贡献', '我的评价'],
    ['① FRF 位置相关振动校正', '最实在的创新。把"位移—响应"反解成"切削点激励"，并明确指出磁力夹持下 H 非对角、'
                        '必须用伪逆。相比[21][22] 只修幅值，这里幅值与相位一起修。'],
    ['② PSR + GASF + CNN 框架', '组合创新。PSR 与 GASF 均为成熟工具（GASF 源自 Wang & Oates 2015 的时序成像），'
                          '把 GASF 从标量序列推广到相空间向量是合理但不算深刻的一步。'],
    ['③ LLE/KE 定量证明可预测性', '有价值但偏装饰：只说明"低混沌、可预测"，并未反过来指导模型设计或超参选择。'],
    ['④ 工程严谨性（交叉 FRF、稳定性叶瓣、R10 分级、ISO 标准）', '不算创新，但在同类论文中少见，是最值得学的地方。'],
], widths=[5.6, 11.0], font=9, align_center=[False, False])

note('一句话定性：这是一篇"接口工程"做得很好、理论增量有限的 TMECH 级方法论文。它的价值在于把"位置效应补偿"'
     '和"时序成像"两条线缝在一起，并用真实机器人铣削实验闭环验证。')

# ============================ 7 ============================
h1('关键局限与可质疑之处（按严重度排序）', '七')

h3('7.1　【最严重】数据划分方式可能导致乐观偏差')
body('3600 个样本来自 60 个加工行程，每个行程取 60 个相邻点。8:2 随机划分会把同一行程的样本同时放进训练集和测试集'
     '——相邻点工况几乎相同、信号高度相关。这会让 96.7% 明显偏乐观。')
rich([('佐证：', True), ('在线实验（真正没见过的工况）整体掉到 90%，变速段只有 60%。', False)])
rich([('正确做法：', True), ('按加工参数组合分组划分（leave-one-condition-out），或至少按行程划分。'
                        '论文没有交代划分是否分组，这是最大的方法学漏洞，也是你最容易规避的坑。', False)])

h3('7.2　只做分类、不做回归，且类别严重不平衡')
bullet('输出是 5 个等级而非 Ra 数值，工艺闭环补偿的粒度太粗。')
bullet('88% 样本落在 Class 2/3/4，论文只报总体准确率，没有 macro-F1、没有每类召回、没有平衡准确率。'
       '极端类"接近 100%"是因为样本少且极易分。')
bullet('Class 1 与 Class 2 混淆本质是工艺上造不出更细样本，但这意味着模型在"高质量区"的判别力没有被真正检验。')

h3('7.3　核心模块 FRF 校正没有做消融')
body('Table IV 只有 PSR-CNN / GASF-CNN / CNN，偏偏没有"不做 FRF 校正"的对照组。'
     '被列为第一贡献的校正模块，其增益只有 Fig. 10(c)"幅值变一致了"这种定性证据。'
     '而"校正后幅值一致"是弱证据——任意逐段归一化都能做到，却会毁掉与粗糙度相关的真实幅值。'
     '缺的关键实验是：把反解出的 F(t) 与测力仪实测切削力对比（哪怕只在一个工况上）。')

h3('7.4　FRF 反卷积的病态性与未讨论的正则化')
body('H_a(ω)⁺ 在反共振点附近奇异，噪声会被放大。全文没有条件数分析、没有奇异值谱、'
     '没有截断 SVD 或 Tikhonov 正则化、没有校正前后信噪比的量化。')

h3('7.5　GASF 推广处的两处不严格')
body('（a）尺度不变导致信息塌缩：s̃_i 对 |s_i| 完全不敏感，m 维向量被压成标量角度序列再成像；'
     '（b）"参考向量可任取、本质相同"不成立，cos(θ_i+θ_j) 依赖 s_0。')

h3('7.6　可预测性论证的量纲问题')
body('LLE 的单位是"每步（per sample）"，但采样率全文未给。−2.1×10⁻³/样本 对应多长的收敛时间无法判断；'
     'KE 同理。这使第三条贡献的可信度打了折扣。')

h3('7.7　噪声鲁棒性测试过于理想')
body('只用加性高斯白噪声 + SNR（式 34）。真实铣削噪声含主轴谐波、冲击、非平稳成分，高斯白噪声是最容易"扛"的一种。'
     '0 dB 下 90.1% 的说服力有限。')

h3('7.8　"在线"的证据不足')
bullet('没有推理延迟、窗长、更新频率、单次决策耗时。')
bullet('Fig. 14(b) 的预测曲线在 0–40 s 内出现大量类别跳变（5→4→3→2 反复切换），'
       '说明没有做时间平滑/后处理，实时输出不稳定。')
bullet('验证点只有 A–E 五个特征点各取 10 点，样本量很小。')

h3('7.9　泛化边界很窄（作者自己也承认）')
body('单一材料（6061）、单刀具（φ10 三刃 38°）、单机器人、平板直边。结论段明确说"目前只做矩形平面件，'
     '未来扩展到自由曲面"。对航空蒙皮的真实几何（双曲率、变厚度、铆接区）差距不小。')

h3('7.10　其他细节瑕疵')
bullet('式 (11) 的分段判据在实现上退化为"均匀 20 mm 分段 + 事后验证 ε<0.1"，方法与实践不完全一致；')
bullet('"prediction" 一词被用来指代"同步分类"，严格说是辨识而非预测；')
bullet('缺少特征可视化/敏感性分析（GASF 图像里哪块纹理起作用），可解释性弱；')
bullet('Fig. 11(a) 互信息出现负值，与互信息恒非负矛盾，作者未讨论；')
bullet('参考文献 [3][29][30] 存在卷号/页码不全的排版错误。')

# ============================ 8 ============================
h1('复现所需但论文缺失的信息', '八')

make_table([
    ['缺失项', '影响'],
    ['采样频率、抗混叠滤波设置', 'τ = 5 无法换算成物理时间，LLE/KE 无法解读'],
    ['单样本窗长 n（多少采样点构成一张 GASF 图）', 'GASF 矩阵尺寸 = (n − (m−1)τ)²，完全无法复现'],
    ['CNN 结构（层数、通道、卷积核、池化、全连接维度、参数量）', '无法复现 96.7%'],
    ['三通道如何融合（3 张 GASF 堆成 3 通道？拼接？分别推理再投票？）', '影响结论可解释性'],
    ['GASF 矩阵是否裁剪/缩放后输入', '影响可复现性'],
    ['训练轮数、batch size、优化器、正则化、早停策略', '影响可复现性'],
    ['随机划分还是分组划分', '决定结论有效性（见 7.1）'],
    ['每类样本数（五类的 count）', '无法判断不平衡程度'],
    ['校正模块的消融实验', '第一贡献缺证据'],
    ['推理耗时 / 在线更新频率', '"在线"主张缺支撑'],
], widths=[8.6, 8.0], font=9, align_center=[False, False])

# ============================ 9 ============================
h1('写作层面值得学的地方', '九')

bullet('三条 gap、三条贡献，一一对应，读者不用猜逻辑。')
bullet('主动补审稿人会问的实验：稳定性叶瓣证明没颤振、Pearson 相关证明参数选择合理、PCA 证明类可分、'
       '混淆矩阵给出错误分布。')
bullet('承认实测反常：明说"理论上 H 应对角，实测交叉项不可忽略"，比强行套理想模型更可信。')
bullet('数据诚实：在线实验只有 60% 的变速段精度也照实报，并解释原因。')
bullet('国际标准落地：Ra 用 ISO 21920-2/-3:2021，分级用 ISO 3:1973 R10 优先数系——等级边界不是拍脑袋定的。')

# ============================ 10 ============================
h1('能借到你自己试切实验里的点', '十')

note('工作区里的《试切数据.docx》记录：超声频率 24 kHz、切削速度 5 mm/s、主轴 5000 r/min、'
     '磁铁个数 4/5/6，Ra 在 6.3–16.7 μm 之间。与本文的对应关系很清楚——你变的正是"夹持刚度"这个变量。')

body('1. 把"磁铁个数"做成刚度变量。本文最有价值的实测结论是"磁铁附近的模态频率变化率低、夹持刚度高、'
     '粗糙度更低"（Fig. 10a）。你已经在变磁铁个数，缺的正是那一步：补一组锤击模态测试，'
     '画出"首阶模态频率/阻尼 vs 磁铁个数"的曲线，把"磁铁数 → 刚度 → Ra"这条链做实，'
     '比单纯报 Ra 数字有说服力得多。')
body('2. 同一条件多点测量。论文每组沿边缘测 60 点做统计；你目前每个磁铁数只有 3 个位置，'
     '且离散度极大（5 块磁铁时 7.2–16.7 μm）。这个离散度要么说明测点选取敏感，要么说明过程本身不稳，'
     '值得先排查再谈建模。')
body('3. 数据按工况分组划分，不要随机划分——这是本文最该被批评、也是你最容易避免的坑。')
body('4. 类别边界用优先数系（R10/R20）而不是等距分箱，评审接受度更高。')
body('5. 若要做在线监测，先确定采样率与窗长，并报告推理延迟，否则"在线"二字站不住。')
body('6. 若考虑位置效应，最省事的起点不是全套 FRF 反卷积，而是"刀位相关幅值归一化 + 相位补偿"两步走，'
     '看增益再决定是否上伪逆。')
rich([('7. 一个必须提醒的边界：', True),
      ('本文的分级只到 Ra 6.3 μm，而你的数据 6.3–16.7 μm 完全在其类别范围之外，'
       '模型不可能直接迁移。要么重新定义分级区间，要么先解决"为什么你的 Ra 高一个数量级"'
       '（刀具、参数、测量方式或单位）这个更根本的问题。', False)])

# ============================ 11 ============================
h1('延伸阅读（顺藤摸瓜）', '十一')

make_table([
    ['方向', '文献', '为什么要读'],
    ['数据驱动非线性系统辨识', '[27] Wang, Lai, Grebogi, Phys. Rep. 644 (2016)', 'PSR 路线的总纲与综述'],
    ['嵌入维数确定', '[37] L. Cao, Physica D 110 (1997)', '本文 m 的确定直接来自此文'],
    ['延迟时间确定', '[36] Fraser & Swinney, Phys. Rev. A 33 (1986)', '互信息法原始出处'],
    ['GASF 原始文献', 'Wang & Oates, AAAI Workshop 2015（本文未直接引用）', '核对"向量版推广"是否与原定义冲突'],
    ['最直接竞争者', '[21][22] Yao 等（Measurement 2021；MSSP 2023）', '位置相关监测的幅值预处理函数值得精读'],
    ['颤振稳定性', '[38] Ding 等全离散法, IJMTM 50 (2010)', '数据集合理性验证的方法来源'],
], widths=[3.4, 7.2, 6.0], font=9, align_center=[False, False, False])

# ============================ 附录 ============================
h1('附录：关键参数速查')

make_table([
    ['类别', '参数'],
    ['工件', '6061 铝，200×200×2 mm，悬伸 20 mm'],
    ['工艺', '主轴 8000 r/min；进给 1–10 mm/s；RDC 0.5–3.0 mm（6 水平）；ADC 2.0 mm'],
    ['刀具', 'φ10 mm 三刃硬质合金平底立铣刀，螺旋角 38°'],
    ['模态', '边缘 11 点等分，间距 20 mm，ε = 0.1'],
    ['PSR', 'τ = 5，m = 15，δ = 0.005，τ_max = 30，m_max = 50'],
    ['数据集', '3600 样本 × 3 通道；训练/测试 2880/720'],
    ['分级', '[0.8,1.6) [1.6,2.0) [2.0,2.5) [2.5,3.2) [3.2,6.3) μm'],
    ['精度', '96.7%（测试集）｜90.1%（0 dB）｜90%（在线整体）｜60%（在线变速段）'],
], widths=[2.6, 14.0], font=9, align_center=[False, False])

spacer(10)
p = doc.add_paragraph()
r = p.add_run('本报告为个人研读笔记，评分与批评意见为阅读者判断，不代表期刊或作者立场。'
              '页码与图表编号均对应 IEEE/ASME TMECH, vol. 31, no. 2, pp. 1741–1752 (2026) 原文。')
rfont(r, 8.5, italic=True, color=GREY)

doc.save(OUT)
print('saved', OUT, os.path.getsize(OUT))
