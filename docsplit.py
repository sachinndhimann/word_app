import os
import zipfile
import xml.etree.ElementTree as et
from shutil import make_archive,copytree,rmtree,move


class SplitDocx:
    def __init__(self,file_path):
        self.file_path=file_path
        self.root_dir=""
        self.base_name=os.path.basename(file_path).split('.')[0]
        self.header='<ns0:document xmlns:ns0="http://sc_____hemas.open'
        self.header = ' n50: document xmIns:50="http://schemas.openxmIformats.org/wordprocessiogm!/2006/main" xmIns:ns1"http://schemas.openxmIformats.org/markup-compatibility/2006" xmlns:ns2="http://schemas.microsoft.com/office/word/2010/wordm/" ns1:ignorable="w14 w15 w16sew16cid w16 w16cex w16sdtdh w14"><nsO:body>'
        self.footer='</ns0:body><ns0:document>'

    def extract(self): 
        with zipfile.ZipFile(self.file_path, 'r') as zip_ref:
            zip_ref.extractall(self.extract_path)
            print(f"Extracted {self.file_path} to {self.extract_path}")


def create_split_dir(self, num_page_breaks): 
    for i in range(num_page_breaks+1): 
        target_dir = os.path.join(self.root_dir, f"{self.base_name}_{i+1}")
        copytree(self.extract_path, target_dir) 

def get_para_tag(self,namespace,elem): 
    key1 = '{http://schemas.microsoft.com/office/word/2010/wordml}paraId'
    key2 = '{http://sc_____hemas.microsoft.com/office/word/2010/wordml}textId'
    key3 = '{'+ namespace['w'] + ')rsidR'
    key4 = '{'+namespace['w'] + ')rsidRDefault' 
    #para_tag = f'''<ns0:p xmlns:ns0."{namespace['w']}" xmlns:nsl."http://''' #TODO
    para_tag = f'''<ns0:p xmIns:ns0="{namespace ['w']}" xmins:ns1="http://schemas.microsoft.com/office/word/2010/wordm/" ns1:paraid="{elem.attrib[key1]}" ns1:textid="{elem.attrib[key2]}" ns0:rsidr="{elem.attrib[key3]}" nsO:rsidrdefault="{elem.attrib[key4]}"›'''
    return para_tag 


def write_document_xml(self, content, count):
     f = open(os.path.join(self.root_dir, "document.xml"), 'w') 
     f.write(content)
     f.close()
     dst_file = f"{self.root_dir}/{self.base_name}_{count}/word/document.xml"
     move(os.path.join(self.root_dir, "document.xml"), dst_file)
     print(f" converting {self.root_dir}/{self.base_name}_{count}.zip to docx")
     tgt_dir = f'{self.root_dir}/{self.base_name}_{count}'
     self.xml_to_docx(tgt_dir) 



def xml_to_docx(self,filepath): 
    make_archive(filepath, "zip", filepath) 
    os.rename(filepath + ".zip", filepath + ".docx") 

def split_by_pagebreak(self):
    content = self.header
    namespace = {'w': "http:// schemas.openxm1formats.org/wordprocessingm1/2006/main"}
    count=0
    cont = et.parse(os.path.join(self.extract_path,'word/document.xml' ))
    root = cont.getroot ()
    num_page_breaks = len(root.findall(".//w:body/w:p//w:lastRenderedPageBreak", namespace))
    print(f"Beginning to split the content into {num_page_breaks+1} based on page breaks")
    self.create_split_dir(num_page_breaks)
    for elem in root.findall(' .//w:', namespace):
        if 'PageBreak' not in et.tostring(elem). decode('utf-8'):
            content = content + et.tostring(elem). decode ('utf-8')
        else:
            count = count + 1
            nested_cont =''
            for child in elem:
                child_cont = et.tostring(child).decode ('utf-8')
                if 'pStyle' in child_cont:
                    para_style = child_cont
                if 'PageBreak' not in child_cont:
                    nested_cont = nested_cont + child_cont
                else:
                    nested_flg = True if 'nso:t' in nested_cont else False
                    content = content + self.get_para_tag(namespace,elem) + nested_cont + '</ns0:p›'+ self.footer if nested_flg else content+ self.footer
                    print(f"writing page {count} in progress")
                    self.write_document_xm1(content, count)
                    child.remove(child.find('.//w:lastRenderedPageBreak', namespace))
                if nested_flg:
                    content = self.header + self.get_para_tag(namespace, elem) + para_style + et. tostring(child). decode( 'utf-8')
                else:
                    content = self.header + self.get_para_tag(namespace, elem) + nested_cont + et.tostring(child) .decode('utf-8')
                nested_cont = content
            content=nested_cont+''
            print(f"writing page {count+1} in progress")
            content=content=self.footer
            self.write_document_xm1(content,count+1)

def split_by_contentlevel(self, clevel):
    content = self.header
    namespace = {'W': "http://schemas.openxmlformats.org/wordprocessingm1/2006/main"}
    count = 0
    cont = et.parse(os.path.join(self.extract_path, 'word/document.xmi'))
    root = cont.getroot ()
    num_page = len([elem.attrib for elem in root.findal1(f"./ /w:body/w:p//w:pstyle", namespace) if clevel in elem.attrib['{http://sc_____hemas.openxmlformats.org/wordprocessingml/2006/main}val']])
    print(f"Beginning to split the content into {num_page} based on contentlevel")
    self.create_split_dir(num_page-1) 
    for elem in root.findall('.//w:p',namespace):
        if clevel not in et.tostring(elem). decode ('utf-8'):
            content = content + et. tostring(elem). decode ('utf-8')
        else:
            if 'nso:p' in content:
                content = content + self.footer
                count = count + 1
                print(f"writing document {count} in progress")
                self.write_document_xm1(content, count)
            content = self.header + et. tostring(elem). decode('utf-8')
    print(f"writing document {count+1} in progress")
    content = content + self.footer
    self.write_document_xm1(content, count+1)


def lambda_split_word(fileobject,filter):
    sp = SplitDocx(fileobject)
    sp.extract ()
    if 'pagebreak' not in filter:
        sp.split_by_contentlevel(filter)
    else:
        sp.split_by_pagebreak()













