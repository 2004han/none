import sys,types,importlib,io,olefile
from pathlib import Path
base=Path(r'E:\试切实验\测量系统公式转换\参考工具\MTEF-py-main')
p=types.ModuleType('mteflib');p.__path__=[str(base)];sys.modules['mteflib']=p
from mteflib.mtef import MTEF
f=Path(r'E:\试切实验\测量系统公式转换\参考工具\mathtype_to_mathml-master\spec\fixtures\input\299.bin')
o=olefile.OleFileIO(f);raw=o.openstream('Equation Native').read()[28:]
m=MTEF();m.reader=io.BytesIO(raw)
for name in ['readLine','readChar','readTMPL']:
 fn=getattr(m,name)
 def wrap(obj,fn=fn,name=name):
  pos=m.reader.tell()-1;result=fn(obj);print(name,pos,vars(obj));return result
 setattr(m,name,wrap)
m.readRecord();m.makeAST();print(m.Translate() if hasattr(m,'Translate') else 'parsed')
print('RAW',raw.hex(' '))
