from docx import Document
from zipfile import ZipFile
from pathlib import Path
p=Path(r'E:\试切实验\机器人试切实验_碳纤维复材毛刺分析报告.docx')
d=Document(str(p))
print('exists',p.exists(),'size',p.stat().st_size)
print('tables',len(d.tables),'inline_shapes',len(d.inline_shapes),'paragraphs',len(d.paragraphs))
with ZipFile(p) as z:
    print('zip_ok', z.testzip(), 'media', len([x for x in z.namelist() if x.startswith('word/media/')]))
