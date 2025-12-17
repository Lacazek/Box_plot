# -*- coding: utf-8 -*-
"""
Created on Wed Nov 12 15:32:02 2025

@author: lacazek
"""

import pandas as pd
import matplotlib.pyplot as plt
from tkinter import Tk, filedialog, simpledialog, messagebox
import os
from collections import defaultdict
import re
import numpy as np
import seaborn as sns
from matplotlib.ticker import MultipleLocator
import matplotlib.patches as mpatches

#variables
yname = "Dose [Gy]"
xname = {"Référence","Evaluation"}
data_ref = "ref"
#data_filter = {"0 UH","-100 UH", "-200 UH", "-300 UH", "-400 UH", "-500 UH", "-600 UH", "-700 UH"}
data_filter = {"TTT replan","0 UH","-100 UH", "-200 UH", "-300 UH", "-400 UH", "-500 UH", "-600 UH", "-700 UH",
               "-100 UH TTT", "-200 UH TTT", "-300 UH TTT", "-400 UH TTT", "-500 UH TTT", "-600 UH TTT", "-700 UH TTT"}
dossier = "B:\\RADIOTHERAPIE\\Physique\\2-Projets en cours\\Bolus Virtuel VMAT\\Etude\\"
paths = "Résultats OARS - UH choisies\\figures\\"
dossier_name = "ALL_UH_seinG"
dossier = dossier +paths+dossier_name
os.makedirs(dossier,exist_ok= True)


#Recherche du fichier
Tk().withdraw()
fichier = filedialog.askopenfilename(
     title="Sélectionnez un fichier à ouvrir",
     filetypes=[("Fichiers CSV", "*.csv"), ("Tous les fichiers", "*.*")]
 )

chemin_fichier_excel = fichier

extension = os.path.splitext(fichier)[-1].lower()

try:
    if extension == ".csv":
        df = pd.read_csv(fichier, encoding="utf-8", sep=";",index_col=False)
    elif extension in [".xlsx", ".xls"]:
        df = pd.read_excel(fichier, sheet_name=0, engine="openpyxl")
    else:
        print("Format non supporté.")
        exit()
except Exception as e:
    print("Erreur de lecture :", e)
    exit()
 #df_set_test = df[~df.apply(lambda row: row.astype(str).str.contains(data_filter, case=False, na=False).any(), axis=1)]

 #df_ref = df[df.apply(lambda row: row.astype(str).apply(lambda x: any(f in x for f in data_filter)).any(), axis=1)]
   
df_ref = df[df.apply(lambda row: row.astype(str).str.contains(data_ref, case=False, na=False).any(), axis=1)]

df_set_test_multiple = {}
for f in data_filter:
    df_test = df[df.iloc[:, 0] == f].copy()
    df_test = df_test.loc[~df_test.iloc[:, 0].str.contains("ref", case=False)]
    df_set_test_multiple[f] = df_test

stats_ref = {}
stats_eval = {}
all_values = []
all_labels = []  
df_set_test_multiple = dict(sorted(df_set_test_multiple.items()))

for col in df_ref.columns[12:]:
    ref_values = pd.to_numeric(df_ref[col].astype(str).str.replace(",", "."), errors='coerce').dropna()  
    all_values = []
    all_labels = []
    palette_colors = ['skyblue']
    all_values.extend(ref_values)
    n_ref = len(ref_values)
    all_labels.extend([f'Référence n={len(ref_values)}']*len(ref_values))
    
    # Ajouter toutes les sous-DataFrames dynamiquement
    for f, df_test in df_set_test_multiple.items():
        Data_values = pd.to_numeric(df_test[col].astype(str).str.replace(",", "."), errors='coerce').dropna()
        n_test = len(Data_values)
        if len(Data_values) > 0:
            all_values.extend(Data_values)
            all_labels.extend([f'{f} n={len(Data_values)}']*len(Data_values))
            palette_colors.append('lightgreen')  # couleur pour les UH
          
    df_plot = pd.DataFrame({
        'Dose [Gy]': all_values,
        'Data': all_labels
    })
    
    # Figure unique
    plt.figure(figsize=(8, 5))        
    sns.boxplot(x='Data', y='Dose [Gy]', data=df_plot, palette=palette_colors)

    # Ajouter les moyennes en croix
    unique_labels = df_plot['Data'].unique()
    for i, label in enumerate(unique_labels):
        mean_val = df_plot[df_plot['Data'] == label]['Dose [Gy]'].mean()
        plt.scatter(i, mean_val, color='red', marker='X', s=80)

    plt.title(f'{col}')
    
    # axes
    y_min = df_plot['Dose [Gy]'].min() - 5
    y_max = df_plot['Dose [Gy]'].max() + 5
    plt.ylim(y_min, y_max)
    plt.xlabel("")
    plt.xticks(rotation=45, ha='right', fontsize=10)
    plt.gca().yaxis.set_major_locator(MultipleLocator(1))
    plt.gca().yaxis.set_minor_locator(MultipleLocator(0.1))
    plt.tick_params(axis='y', which='major', length=6, width=1.5)
    plt.tick_params(axis='y', which='minor', length=3, width=1)

    # Sauvegarde
    plt.tight_layout()
    plt.savefig(os.path.join(dossier, f"{col}_allUH.png"))
    plt.close()
    

    

