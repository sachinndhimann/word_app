from traceback import print_tb
from docx import Document 
from utils import CreateDocument
file=r"C:\Users\sachi\Downloads\Complete Document.docx"
#document = Document ()
document = Document(file)
sections=document.sections
count=0
master_dict={} 
tr_levels=[]
for index,paragraph in enumerate (document.paragraphs):
    result=paragraph.text.replace(" ","")
    if result:
        tr_levels.append({index:paragraph.style.name})

#xit(1)

pairs=[]
pair=[]
for index, object in enumerate (tr_levels):
    key=list(object.keys())[0]
    print(object[key])
    if index > 0 and "heading" in object[key].lower():
        pairs.append (pair)
        pair=[]
        pair.append (key)
    elif object==tr_levels[-1]:
        pair.append (key)
        pairs.append (pair)
    else:
        pair.append (key)

def get_style(tr_levels,index):
    for level in tr_levels:
        key=list(level.keys())[0]
        if key==index:
            return level [key]


#print(pairs)
file=r"C:\Users\sachi\Downloads\Blank.docx"

for index, pair in enumerate (pairs) :
    doc= Document (file)
    for data in pair:
        #print(data,document.paragraphs[data].text)
        doc.add_paragraph(document .paragraphs [data]. text, style=get_style(tr_levels,data)) 
        doc.save(r"C:\Users\sachi\Downloads\split_{name}.docx".format(name=index))