from bs4 import BeautifulSoup
import re 
import pandas as pd
from tkinter import filedialog as fd

def findOccurrences(s, ch):
    return [i for i, letter in enumerate(s) if letter == ch]

def takeParentList(xsd_element):
    lista_parent=[]
    parent_element=xsd_element['name']
    while parent_element!='':
        try:
            parent_element=xsd_element.find_parent(re.compile('(xsd:element.*)'))['name']
            xsd_element=xsd_element.find_parent(re.compile('(xsd:element.*)'))
            lista_parent.append(parent_element)
        except TypeError:
            parent_element=''
    
    return lista_parent
            

input_file=fd.askopenfilename()  
with open(input_file, 'r',encoding='utf-8') as file:
    parser=BeautifulSoup(file, 'html.parser')

path=input_file[:findOccurrences(input_file,'.')[-1]] + '_output.xlsx'  
writer = pd.ExcelWriter(path, engine = 'xlsxwriter')

len_parent_prev=0
pattern_PozUsz=re.compile('(PozycjaUszczegolawiajaca.*)')
raport=''
complexType=parser.find_all(re.compile('(xsd:complextype.*)'), attrs={"name":True})
col=['desc','tag_name','parent', 'level']
level=0
df=pd.DataFrame(columns=col)

for raport in complexType:
    lsl=[]
    raport_name=raport['name']
    complexType_name=raport['name']
    element = raport.find_all(re.compile('(xsd:element.*)'))
    
    for i in element:
        tag_name = i['name']
        try:
            parent=i.find_parent(re.compile('(xsd:element.*)'))['name']
        except TypeError:
            parent=complexType_name      
        
        
        len_parent=len(takeParentList(i))
        
        try:
            desc=i.find('xsd:documentation').string
        except AttributeError:
            desc=''
        
        
        if pattern_PozUsz.match(tag_name):
            lsl.append(['',tag_name,parent,len_parent])
        else:
            lsl.append([desc,tag_name,parent,len_parent])
    
        try:
            if i['type']=='dtsf:TPozycjaSprawozdania' or i['type']=='dtsf:TPozycjaSprawozdaniaTys':
                parent=tag_name
                tag_name='PozycjaUszczegolawiajaca'  
                len_parent=len_parent_prev+1
                lsl.append(['',tag_name,parent,len_parent])

            if i['type']=='dtsf:TPozycjaSprawozdania' or i['type']=='dtsf:TKwotyPozycjiSprawozdania' or i['type'] == 'dtsf:TKwotyPozycjiSprawozdaniaTys' or i['type']=='dtsf:TPozycjaSprawozdaniaTys':
                lsl.append(['','Podpozycja',tag_name + ' - ' + parent,len_parent+1])
        except KeyError:
            pass
        
        len_parent_prev=len_parent
        df=pd.DataFrame(lsl, columns=col)   
        df.to_excel(writer,sheet_name=raport_name[:30])

writer.save()
writer.close()