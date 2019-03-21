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


def createChildPath(ls_parent,parent_additional,tag_name,raport_name,txt):
    list_parent=''
    child_path=''
    
    for i_parent in ls_parent:
        list_parent= '->' + i_parent + list_parent
        
    if tag_name=='PozycjaUszczegolawiajaca':
        list_parent=  list_parent + '->' +  parent_additional
        
    child_path=raport_name  + list_parent + '->' +  tag_name + '->' + txt
    
    return child_path
            
input_file=fd.askopenfilename()  
with open(input_file, 'r',encoding='utf-8') as file:
    parser=BeautifulSoup(file, 'html.parser')

path=input_file[:findOccurrences(input_file,'.')[-1]] + '_output.xlsx'  
writer = pd.ExcelWriter(path, engine = 'xlsxwriter')

len_parent_prev=0
pattern_PozUsz=re.compile('(PozycjaUszczegolawiajaca.*)')
raport=''
complexType=parser.find_all(re.compile('(xsd:complextype.*)'), attrs={"name":True})
col=['desc','tag_name','parent', 'level', 'path_to_child']
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
        
        ls_parent=takeParentList(i)
        len_parent=len(ls_parent)
        try:
            desc=i.find('xsd:documentation').string
        except AttributeError:
            desc=''
            
        if tag_name=='XXII':
            print("Stop1!")    
            
        if str(tag_name)=='XXII':
            print("Stop2!")      
    
        child_path=createChildPath(ls_parent,parent, tag_name,raport_name,'{0}')
        
        #Pozycja_Uszczegolawiajaca_X
        if pattern_PozUsz.match(tag_name):
            lsl.append(['',tag_name,parent,len_parent,child_path])
        else:
            lsl.append([desc,tag_name,parent,len_parent,child_path])
    
        #Pozycja_Uszczegolawiajaca    
        
        try:
            if i['type']=='dtsf:TPozycjaSprawozdania' or i['type']=='dtsf:TPozycjaSprawozdaniaTys':
                parent=tag_name
                tag_name='PozycjaUszczegolawiajaca'  
                len_parent=len_parent_prev+1
                child_path=createChildPath(ls_parent,parent, tag_name,raport_name,'{0}')
                
                lsl.append(['',tag_name,parent,len_parent,child_path])

        #Podpozycja    
            if i['type']=='dtsf:TPozycjaSprawozdania' or i['type']=='dtsf:TKwotyPozycjiSprawozdania' or i['type'] == 'dtsf:TKwotyPozycjiSprawozdaniaTys' or i['type']=='dtsf:TPozycjaSprawozdaniaTys':
                child_path=createChildPath(ls_parent,parent, tag_name,raport_name,'{0}->Podpozycja->{0}')
                lsl.append(['','Podpozycja',tag_name + ' - ' + parent,len_parent+1,child_path])
                
        except KeyError:
            pass
    
        
        len_parent_prev=len_parent
        df=pd.DataFrame(lsl, columns=col)   
        df.to_excel(writer,sheet_name=raport_name[:30])

writer.save()
writer.close()