from pathlib import Path
from html.parser import HTMLParser
p=Path(r'E:\试切实验\机器人试切实验_碳纤维复材毛刺分析报告.html')
class P(HTMLParser):
    def __init__(self): super().__init__(); self.img=0; self.h=[]
    def handle_starttag(self,tag,attrs):
        if tag=='img': self.img+=1
        if tag in ('h1','h2'): self.h.append(tag)
q=P(); q.feed(p.read_text(encoding='utf-8'))
text=p.read_text(encoding='utf-8')
print('exists',p.exists(),'size',p.stat().st_size,'img_tags',q.img,'headings',len(q.h),'data_images',text.count('data:image/jpeg;base64,'),'utf8_ok',True)
