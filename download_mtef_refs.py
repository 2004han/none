import requests,concurrent.futures,zipfile,io
from pathlib import Path
base=Path(r'E:\试切实验\测量系统公式转换\参考工具');base.mkdir(exist_ok=True)
repos=['jure/mathtype_to_mathml','AndyQsmart/MTEF-py','zhexiao/mtef-go']
def f(repo):
 try:
  info=requests.get('https://api.github.com/repos/'+repo,timeout=20).json();branch=info.get('default_branch','master')
  r=requests.get('https://codeload.github.com/'+repo+'/zip/refs/heads/'+branch,timeout=40);r.raise_for_status()
  z=zipfile.ZipFile(io.BytesIO(r.content));z.extractall(base)
  print(repo,[(n,z.getinfo(n).file_size) for n in z.namelist() if n.endswith(('.docx','.bin','.pdf','.html','.md'))][:30])
 except Exception as e:print(repo,e)
with concurrent.futures.ThreadPoolExecutor() as ex:list(ex.map(f,repos))
