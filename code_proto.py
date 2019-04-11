import spacy
import pandas as pd
import numpy as np
import re
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
import os
import sys, getopt
from io import StringIO 
import docx2txt
import sys
#get_ipython().config.get('IPKernelApp', {})['parent_appname'] = ""

nlp1 = spacy.load('C:/Users/a630934/AppData/Local/Continuum/anaconda3/Lib/site-packages/fr_core_news_md/fr_core_news_md-2.0.0')
nlp = spacy.load('C:/Users/a630934/AppData/Local/Continuum/anaconda3/Lib/site-packages/fr_core_news_sm/fr_core_news_sm-2.0.0')

def convert(fname, pages=None):
    if not pages:
        pagenums = set()
    else:
        pagenums = set(pages)

    output = StringIO()
    manager = PDFResourceManager()
    converter = TextConverter(manager, output, laparams=LAParams())
    interpreter = PDFPageInterpreter(manager, converter)

    infile = open(fname, 'rb')
    for page in PDFPage.get_pages(infile, pagenums):
        interpreter.process_page(page)
    infile.close()
    converter.close()
    text = output.getvalue()
    output.close
    return text 

regex = r"[a-zA-Z0-9_\.]+@[a-zA-Z0-9_\-]+\.[a-zA-Z]+"
regex2='[0-9]{10}'
regex2bis='\d{2} \d{2} \d{2} \d{2} \d{2}'
regex3='\d{4}-\d{4}-\d{4}-\d{4}'
regex4='\d{3}-\d{3}-\d{3}-\d{3}-\d{3}'
regex5= '\d{1} \d{2} \d{2} \d{2} \d{3} \d{3} \d{2}'
regex6='[0-9]{14}'
def get_personal_information(commentaire,nlp):
    nom_list=[]
    company_list=[]
    location_list=[]
    number_list=[]
    number_list_bis=[]
    mail_list=[]
    banc_list=[]
    sec_list=[]
    siret_list=[]
    
    doc = nlp(commentaire)
    text=[]
    label=[]
    for ent in doc.ents:
        #print(ent.text, ent.start_char, ent.end_char, ent.label_)
        text.append(ent.text)
        label.append(ent.label_)
    keys = label
    values = text
    dictionary = dict(zip(keys, values))
    if 'PER' in keys:
        nom_list.append(dictionary['PER'])
    else:
        nom_list.append(np.nan)

    if 'LOC' in keys:
        location_list.append(dictionary['LOC'])
    else:
        location_list.append(np.nan)

    if 'ORG' in keys:
        company_list.append(dictionary['ORG'])
    else:
        company_list.append(np.nan)
     
    if len(re.findall(regex2bis, commentaire ))>0:
        number_list_bis.append(re.findall(regex2bis, commentaire )[0])
    else:
        number_list_bis.append(np.nan)
    
    if len(re.findall(regex, commentaire ))>0:
        mail_list.append(re.findall(regex, commentaire )[0])
    else:
        mail_list.append(np.nan)
        
    if len(re.findall(regex3, commentaire ))>0:
        banc_list.append(re.findall(regex3, commentaire )[0])
    else:
        banc_list.append(np.nan)
        
    if len(re.findall(regex5, commentaire ))>0:
        sec_list.append(re.findall(regex5, commentaire )[0])
    else:
        sec_list.append(np.nan)  
        
    if len(re.findall(regex6, commentaire ))>0:
        siret_list.append(re.findall(regex6, commentaire )[0])
        if len(re.findall(regex2, commentaire ))>0:
            if  str(re.findall(regex2, commentaire )[0]) in  str(re.findall(regex6, commentaire )[0]):
                number_list.append(np.nan)
            else:
                 number_list.append(re.findall(regex2, commentaire )[0])
        else:
            number_list.append(np.nan) 
    else:
        siret_list.append(np.nan)     
        if len(re.findall(regex2, commentaire ))>0:
                 number_list.append(re.findall(regex2, commentaire )[0])
        else:
            number_list.append(np.nan) 
    
    
    return [nom_list,company_list,location_list,number_list,number_list_bis,mail_list,banc_list,sec_list,siret_list]




def extract_info(data,file_path,column):
    g_file_path=[]
    g_nom_list=[]
    g_company_list=[]
    g_location_list=[]
    g_number_list=[]
    g_number_list_bis=[]
    g_mail_list=[]
    g_banc_list=[]
    g_sec_list=[]
    g_siret_list=[]
    for i in range(data.shape[0]):
        g_file_path.append(file_path)
        a=get_personal_information(data[column][i],nlp1)
        g_nom_list.append(a[0][0])
        g_company_list.append(a[1][0])
        g_location_list.append(a[2][0])
        g_number_list.append(a[3][0])
        g_number_list_bis.append(a[4][0])
        g_mail_list.append(a[5][0])
        g_banc_list.append(a[6][0])
        g_sec_list.append(a[7][0])
        g_siret_list.append(a[8][0])

    list_of_tuples = list(zip(g_file_path,g_nom_list, g_location_list, g_company_list, g_number_list,g_number_list_bis, g_mail_list,g_banc_list,g_sec_list,g_siret_list))
    df = pd.DataFrame(list_of_tuples, columns = ['file','Name', 'Location', 'Company', 'Phone_number','Phone_number_bis', 'e-mail','credit_card','SS_number','Siret'])
    df=df.dropna(axis=0, how='all',subset=['Name','Phone_number','Phone_number_bis', 'e-mail','credit_card','SS_number','Siret'])
    return df 



def get_file_to_process(pdfDir):
    list_extention=['pdf','csv','docx','xlsx']
    list_to_process=[]
    for root, dirs, files in os.walk(pdfDir, topdown=False):
        for name in files:
            if name.split(".")[-1] in list_extention:
                list_to_process.append(os.path.join(root, name))
                #print(os.path.join(root, name).split(".")[-1])
    return list_to_process


def data_personal(pdfDir,outputDir):
    data_glob=[]
    list_file=[]
    list_to_process=get_file_to_process(pdfDir)
    nb_file=len(list_to_process)
    i=0
    for pdf in list_to_process: #iterate through pdfs in pdf directory
            file=pdf.replace('\\','/').split('/')[-1]
            print("processing file:", file)
            i+=1
            print("processing file:" +' ' +str(i) +' '+ "from:"+' ' +str(nb_file))
            fileExtension = pdf.split(".")[-1]
            if fileExtension == "pdf":
                #pdfFilename = pdfDir + pdf
                text = convert(pdf) #get string of text content of pdf
                text=text.replace('\n',' ')
                data= [x for x in text.split('.'+' ') if x != ""]
                data=pd.DataFrame(data)
                extraction=extract_info(data,pdf,0)
            elif fileExtension == "docx":
                #pdfFilename = pdfDir + pdf
                text = docx2txt.process(pdf) #get string of text content of pdf
                text=text.replace('\n',' ')
                data= [x for x in text.split('.'+' ') if x != ""]
                data=pd.DataFrame(data)
                extraction=extract_info(data,pdf,0)
            elif fileExtension == "xlsx":
                #pdfFilename = pdf
                data=list(pd.read_excel(pdf, sheet_name=None).values())[0]
                extraction=extract_info(data,pdf,'Sage')
            else:
                #pdfFilename = pdfDir + pdf
                data=pd.read_csv(pdf, sep=';',encoding='latin-1')
                extraction=extract_info(data,pdf,'Commentaire Fin')
            if extraction.shape[0]>0:
                list_file.append(pdf)
                data_glob.append(extraction)
    data_glob=pd.concat(data_glob)
    data_glob.to_csv(os.path.join(outputDir,'personaldata.csv'))
    textFile = open(os.path.join(outputDir,'file_with_personal_information.txt'), "w") #make text file
    for name in list_file:
        textFile.write(name.replace('/','\\')+'\n') #write text to text file
    textFile.close()
    return list_file, data_glob

data_personal(sys.argv[1],sys.argv[2])