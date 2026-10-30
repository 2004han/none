import urllib.request,json,concurrent.futures
urls=['https://api.github.com/search/repositories?q=mtef','https://docs.wiris.com/mathtype/en/mathtype-sdk/mtef5.html','https://docs.wiris.com/en/mathtype/mathtype_desktop/mathtype-sdk/mtef5.html']
def f(u):
 try:
  r=urllib.request.urlopen(u,timeout=20);b=r.read();print(u,r.status,r.url,len(b));
  if 'github' in u:
   data=json.loads(b);print([(x['full_name'],x['description']) for x in data.get('items',[])][:15])
  else:
   from pathlib import Path
   (Path(r'E:\试切实验\测量系统公式转换')/(str(urls.index(u))+'.html')).write_bytes(b)
 except Exception as e:print(u,str(e))
with concurrent.futures.ThreadPoolExecutor() as x:list(x.map(f,urls))

