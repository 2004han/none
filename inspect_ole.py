from pathlib import Path
import olefile
b=Path(r'E:\试切实验\测量系统公式转换\参考工具\mathtype_to_mathml-master\spec\fixtures\input')
for fn in ['equation10.bin','equation1.bin','299.bin','629.bin']:
 o=olefile.OleFileIO(b/fn);print(fn,'CLSID',o.root.clsid)
 for p in o.listdir():
  d=o.openstream(p).read();print(p,len(d),d[:90].hex(' '))
