import os
from docx import Document
class DocMerge():
    def __init__(self,list_of_documents,filename):
        self.documents=list_of_documents
        self.filename=filename
        self.merged_doc=Document(r"C:\Users\sachi\Downloads\Merged.docx")
    

    def merge(self):
        for path in self.documents:
            doc= Document(path)
            result=doc.paragraphs 
            for value in result:
                self.merged_doc.add_paragraph(value.text,style=value.style.name)
            self.merged_doc.save(r"C:\Users\sachi\Downloads\{}_Merged.docx".format(self.filename))


obj=DocMerge([r"C:\Users\sachi\Downloads\split_0.docx",r"C:\Users\sachi\Downloads\split_1.docx"],"legal")
obj.merge()