# -*- coding: utf-8 -*-
"""
Created on Wed Feb 12 09:37:17 2020

@author: GM
"""
import numpy as np
import pandas as pd
import os
import glob
from  tkinter import *
from tkinter import filedialog
from openpyxl import load_workbook

liste_ident=[]
liste_identhdv=[]
liste_identhdvptv=[]
liste_hdv=[]   #création liste avant la boucle pour qu'elle s'incremente
liste_hdvptv=[]
liste_hdvptvdf =[]
liste_plan=[]
liste_index=[] 
root = Tk()
root.filedirectory =  filedialog.askdirectory(title = "Sélectionner le dossier où se trouvent les fichiers")
root.withdraw()
os.chdir(root.filedirectory)
files_with_html = glob.glob("*.html") #list of html files
#if not os.path.exists ('results.xlsx'):
#    file_path = 'results.xlsx'
#else:
#    file_path = 'results.xlsx'

for files in files_with_html:
    identhdv=[]
    identhdvptv=[]
    data=pd.read_html(files, flavor='bs4')#lecture de toutes les df de la page html
    #recuperation identité dossier
    identdf = pd.DataFrame(data[1])#selection de la df cible hdv
    ident=identdf.iloc[1:,0:]
    ident.columns=['Patient','Date analyse','Plan','Machine','Opérateur']
    #recuperation données um et mf et HI
    plandf = pd.DataFrame(data[5])#selection de la df cible hdv
    plan=plandf.iloc[1:,1:5]
    plan.columns=['Dmax CE (%)','UM','MF','HI']
    #recuperation données hdv PTV
    df_7 = data[7]
    HI = df_7[df_7[0].str.contains('CTV dosi', na=False, case=False) |
                        df_7[0].str.contains('CTV_dosi', na=False, case=False)][5].values[0]
    plan['HI'] =  HI
        
    hdvptv = pd.DataFrame(data[9])#selection de la df cible hdv
    hdvptv.columns=['Cible','Volume (cc)','Dmax (%)','D99% (%)','D95% (%)','D90% (%)','Dmoy (Gy)','D50% (Gy)','Dmin (Gy)','test']#rename des headers
     #rajout de l'ident dans les listes hdv et index dosi
    for i in range(len(hdvptv)):#iteration de la ligne d'idendité sur le nombre de ligne de le la df hdv
        identhdvptv.append(ident)
        identthdvptv=pd.concat(identhdvptv) 
    #recuperation données hdv oar
    hdvdf = pd.DataFrame(data[11])#selection de la df cible hdv
    hdv=hdvdf.iloc[1:,:6]
    hdv=hdv.replace({4: ' %'}, {4: ''}, regex=True).replace({4: ' Gy'}, {4: ''}, regex=True).replace({4: ' cc'}, {4: ''}, regex=True).replace({1: ' cc'}, {1: ''}, regex=True)#remplacement de certains characteres dans la df
    hdv.columns=['OAR','Volume (cc)','Dose','Contrainte','Valeur','referentiel']#rename des headers
    #rajout de l'ident dans les listes hdv et index dosi
    for i in range(len(hdv)):#iteration de la ligne d'idendité sur le nombre de ligne de le la df hdv
        identhdv.append(ident)
        identthdv=pd.concat(identhdv)
        
    hdv.columns=['OAR','Volume (cc)','V ir','Contrainte','Valeur','referentiel']#rename des headers
    #rajout de l'ident dans les listes hdv et index dosi
    for i in range(len(hdv)):#iteration de la ligne d'idendité sur le nombre de ligne de le la df hdv
        identhdv.append(ident)
        identthdv=pd.concat(identhdv)  
        
    #iterations append
    liste_ident.append(ident)
    liste_identhdv.append(identthdv)
    liste_identhdvptv.append(identthdvptv)
    liste_hdv.append(hdv)
    liste_hdvptv.append(hdvptv)
    liste_hdvptvdf.append(hdvptv)
    liste_plan.append(plan)
    
    hdviter=pd.concat(liste_hdv)#concatenation des df dans 1 grande
    hdvptviter=pd.concat(liste_hdvptv)
    identhdviter=pd.concat(liste_identhdv)
    identhdvptviter=pd.concat(liste_identhdvptv)    
    identiter=pd.concat(liste_ident)
    planiter=pd.concat(liste_plan)    
    
        
with pd.ExcelWriter("Data.xlsx", engine='openpyxl') as writer:           
    identhdviter.to_excel(writer,sheet_name='Analyse HDV OARs',startrow=0, startcol=0,index=False)
    hdviter.to_excel(writer,sheet_name='Analyse HDV OARs',startrow=0, startcol=5,index=False)
    identhdvptviter.to_excel(writer,sheet_name='Analyse HDV PTVs',startrow=0, startcol=0,index=False)
    hdvptviter.to_excel(writer,sheet_name='Analyse HDV PTVs',startrow=0, startcol=5,index=False)
    identiter.to_excel(writer,sheet_name='Analyse index dosi',startrow=0, startcol=0,index=False)
    planiter.to_excel(writer,sheet_name='Analyse index dosi',startrow=0, startcol=5,index=False)   
    workbook=writer.book


