from pathlib import Path
import base64, html

out = Path(r'E:\试切实验\机器人试切实验_碳纤维复材毛刺分析报告.html')
img_dir = Path(r'C:\Users\86195\AppData\Local\Temp')
photos = [
    ('codex-clipboard-8af5440c-59f0-46dc-a36d-90ae0e60cfd1.jpg','图1  3k，不开','连续锯齿状毛刺，局部有较大的纤维状突起。'),
    ('codex-clipboard-c819c119-e380-4809-b6b0-04805e27e1c9.jpg','图2  3k，开','边缘较平直，仅见少量短小毛刺。'),
    ('codex-clipboard-8acb49ba-9468-4fc4-b4e2-e725de306820.jpg','图3  4k，不开','边缘存在较密集细小毛刺和局部纤维毛刺。'),
    ('codex-clipboard-21878c1a-f9fe-4c1e-87bc-b7269d8965e5.jpg','图4  4k，开','边缘粗糙，可见纤维拔出和碎屑，属于异常偏差工况。'),
    ('codex-clipboard-081e4d19-b84b-46ef-b502-967052c9d62f.jpg','图5  5k，不开','上侧边缘有连续细小毛刺，下侧相对整齐。'),
    ('codex-clipboard-7b444b70-5598-4ace-83f4-290e65bdbb03.jpg','图6  5k，开','边缘最平整，毛刺少且短，当前照片中表现最好。'),
    ('codex-clipboard-47615a87-faf2-4bdc-bc6f-20b57d43b7a1.jpg','图7  6k，不开','两侧毛刺不规则，局部有较大突起和松散颗粒。'),
    ('codex-clipboard-8e6d909d-61a2-42d7-a998-b8eb925efaec.jpg','图8  6k，开','仍有少量细小毛刺，但整体边缘较整齐。'),
    ('codex-clipboard-e9eb3b8f-88b6-4732-93f6-981d3d90abb0.jpg','图9  8k，不开','毛刺连续且密集，局部突起较大，纤维撕裂最明显。'),
    ('codex-clipboard-e803d73a-8eeb-4f3d-97eb-f8aac434ed41.jpg','图10  5k，开（重复照片）','与图6内容相同或高度相似，作为同一工况的重复照片。'),
]
for fn,_,_ in photos:
    if not (img_dir/fn).exists():
        raise FileNotFoundError(img_dir/fn)

def data_uri(path):
    return 'data:image/jpeg;base64,' + base64.b64encode(path.read_bytes()).decode('ascii')

rows = [
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

def esc(s): return html.escape(s, quote=True)

table_rows = ''.join(f'<tr><td>{esc(a)}</td><td>{esc(b)}</td><td class="level">{esc(c)}</td></tr>' for a,b,c in rows)
photo_cards = []
for fn, cap, note in photos:
    uri = data_uri(img_dir/fn)
    photo_cards.append(f'''<figure class="photo-card"><img src="{uri}" alt="{esc(cap)}"><figcaption><strong>{esc(cap)}</strong><br><span>{esc(note)}</span></figcaption></figure>''')

html_doc = f'''<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>机器人试切实验碳纤维复材周铣毛刺分析报告</title>
<style>
:root {{ --blue:#2f5597; --light:#eaf2f8; --border:#d9d9d9; --muted:#555; }}
* {{ box-sizing:border-box; }}
body {{ margin:0; background:#f2f4f7; color:#111; font-family:"Microsoft YaHei","Noto Sans CJK SC",Arial,sans-serif; line-height:1.65; }}
.page {{ width:210mm; min-height:297mm; margin:14mm auto; padding:18mm 18mm 16mm; background:#fff; box-shadow:0 1px 10px rgba(0,0,0,.08); }}
h1 {{ margin:0 0 8px; text-align:center; font-size:28px; line-height:1.25; }}
.subtitle {{ text-align:center; color:#444; font-size:17px; margin-bottom:24px; }}
h2 {{ margin:24px 0 8px; font-size:20px; border-left:4px solid var(--blue); padding-left:10px; }}
p {{ margin:7px 0; font-size:14px; }}
.meta {{ width:100%; border-collapse:collapse; margin:10px auto 22px; }}
.meta td {{ border:1px solid var(--border); padding:9px 12px; vertical-align:middle; font-size:14px; }}
.meta td:first-child {{ width:24%; background:var(--light); font-weight:700; }}
.summary {{ font-size:15px; }}
table.analysis {{ width:100%; border-collapse:collapse; margin:10px 0 14px; font-size:13px; }}
.analysis th {{ background:var(--blue); color:#fff; padding:9px 8px; text-align:center; }}
.analysis td {{ border:1px solid var(--border); padding:8px; vertical-align:middle; }}
.analysis tr:nth-child(even) td {{ background:#f5f7fa; }}
.analysis td:first-child, .analysis .level {{ text-align:center; font-weight:600; }}
ul {{ margin:6px 0 10px 24px; padding:0; font-size:14px; }}
li {{ margin:4px 0; }}
.photo-grid {{ display:grid; grid-template-columns:1fr 1fr; gap:14px; margin-top:12px; }}
.photo-card {{ margin:0; padding:8px; border:1px solid var(--border); background:#fff; break-inside:avoid; }}
.photo-card img {{ display:block; width:100%; height:auto; max-height:118mm; object-fit:contain; background:#f6f6f6; }}
.photo-card figcaption {{ margin-top:7px; font-size:12px; line-height:1.45; color:var(--muted); }}
.photo-card figcaption strong {{ color:#111; }}
.note {{ color:var(--muted); font-size:13px; }}
@media print {{
  body {{ background:#fff; }}
  .page {{ width:auto; min-height:auto; margin:0; padding:12mm 14mm; box-shadow:none; }}
  .page + .page {{ page-break-before:always; }}
  h2 {{ break-after:avoid; }}
  table, figure {{ break-inside:avoid; }}
}}
@media (max-width:900px) {{
  body {{ background:#fff; }}
  .page {{ width:100%; min-height:auto; margin:0; padding:24px; box-shadow:none; }}
  .photo-grid {{ grid-template-columns:1fr; }}
}}
</style>
</head>
<body>
<section class="page">
  <h1>机器人试切实验碳纤维复材周铣毛刺分析报告</h1>
  <div class="subtitle">超声刀柄开启状态与主轴转速对比</div>
  <table class="meta">
    <tr><td>实验对象</td><td>碳纤维增强复合材料周铣边缘</td></tr>
    <tr><td>主要变量</td><td>主轴转速、超声刀柄开关状态</td></tr>
    <tr><td>照片数量</td><td>10 张照片，对应 9 个有效工况（5k，开重复 1 张）</td></tr>
  </table>
  <h2>报告结论</h2>
  <p class="summary">在当前照片范围内，5k、超声开启时的边缘最平整、毛刺最少；3k 开和 6k 开次之。不开超声时，6k 和 8k 的毛刺明显加重，其中 8k 不开最差。超声总体有利于抑制纤维拔出，但 4k 开出现异常偏差，表明效果存在参数窗口。</p>
  <h2>1 观察范围与评价方法</h2>
  <p>本报告根据照片中标注的主轴转速和超声刀柄开关状态，对切削边缘的毛刺、纤维拔出、碎屑和局部撕裂进行相对评价。照片未提供比例尺，因此等级仅用于工况之间的比较，不代表实际毛刺高度或宽度。</p>
  <p>照片中主要可见三类边缘缺陷：连续锯齿状毛刺、纤维端部拔出形成的毛刺，以及附着在切边上的松散碎屑。暂未观察到大面积贯穿性分层。</p>
  <h2>2 各工况照片分析</h2>
  <table class="analysis">
    <thead><tr><th>工况</th><th>边缘形貌</th><th>毛刺等级</th></tr></thead>
    <tbody>{table_rows}</tbody>
  </table>
</section>
<section class="page">
  <h2>3 转速与超声状态的影响</h2>
  <p>不开超声时，3k～5k 的毛刺处于中等水平；当转速提高到 6k 和 8k 后，边缘毛刺明显增多，且出现较大的纤维突起和松散颗粒。仅从本组照片看，转速升高并未带来持续的边缘质量改善。</p>
  <p>开启超声后，3k、5k 和 6k 的毛刺均明显减小，说明超声振动可能降低了切削过程中的纤维弯曲、拉扯和拔出。4k 开的照片却表现较粗糙，属于与整体趋势不一致的工况，可能与超声振幅匹配、刀具动态状态、进给参数、纤维方向或局部夹持条件有关。</p>
  <p>综合排序（由好到差）可暂定为：5k 开，其次为 3k 开和 6k 开；8k 不开最差。由于缺少 8k 开照片，尚不能判断超声对 8k 工况的改善幅度。</p>
  <h2>4 试验结论与后续建议</h2>
  <ul>
    <li>建议优先以 5k、超声开启作为后续试验的基准工况，并进行至少 3 次重复验证。</li>
    <li>对 4k、超声开启进行复测，同时检查超声振幅、刀柄谐振状态、刀具跳动、刀具磨损和工件夹持刚度。</li>
    <li>补充 8k、超声开启工况，以形成完整的转速—超声对照。</li>
    <li>后续采用固定拍摄角度、光照和放大倍率，并测量毛刺最大高度、单位长度毛刺数量、纤维拔出长度和分层宽度。</li>
    <li>若高转速条件下仍出现明显毛刺，应优先检查刀具跳动、进给量、每齿进给和切削方向，而不宜只通过继续提高主轴转速来改善。</li>
  </ul>
  <h2>附录 试切边缘照片</h2>
  <p class="note">以下照片按附件顺序排列。图 6 和图 10 均标注为 5k，开，图 10 与图 6 内容重复，作为重复照片保留。</p>
  <div class="photo-grid">{''.join(photo_cards)}</div>
</section>
</body>
</html>
'''
out.write_text(html_doc, encoding='utf-8')
print(out)
