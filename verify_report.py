from docx import Document
from zipfile import ZipFile
from pathlib import Path
p=Path(r'E:\试切实验\机器人试切实验_碳纤维复材毛刺分析报告.docx')
doc=Document(str(p))
print('paragraphs', len(doc.paragraphs), 'tables', len(doc.tables), 'sections', len(doc.sections))
print('inline_shapes', len(doc.inline_shapes))
print('title', doc.core_properties.title)
with ZipFile(p) as z:
    names=z.namelist()
    media=[n for n in names if n.startswith('word/media/')]
    print('media', len(media), media[:3], media[-2:])
    print('zip_ok', z.testzip())
print('qa_dir_exists', Path(r'E:\试切实验\qa_render').exists())
