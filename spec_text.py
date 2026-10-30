from html.parser import HTMLParser
from pathlib import Path
class P(HTMLParser):
 def __init__(self):super().__init__();self.skip=0;self.t=[]
 def handle_starttag(self,t,a):
  if t in ['script','style']:self.skip+=1
 def handle_endtag(self,t):
  if t in ['script','style']:self.skip-=1
 def handle_data(self,t):
  if not self.skip and t.strip():self.t.append(t.strip())
p=P();p.feed(Path(r'E:\试切实验\mtef_spec.html').read_text('utf-8'));print('\n'.join(p.t)[:15000])
