# -*- coding: utf-8 -*-
"""
Created on Wed Aug 28 08:17:10 2024

@author: 129391
"""
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
import pandas as pd

sheet = 'Analyse HDV OARs'
file = 'Data.xlsx'
column_df_1 = 'Dose'
column_df_2 = 'OAR_regroup'
column_df_3 = 'Plan'
Value = 'Valeur'
path = r'B:\RADIOTHERAPIE\Physique\2-Projets en cours\IMRT sein\Etude\Comparaison des doses 3D vs IMRT'
df = pd.read_excel(path+ r'\Comparaison_V2'+ '\\'+ file ,sheet_name= sheet)

organes = df['OAR_regroup'].dropna().unique()
techniques = df['Plan'].dropna().unique()

stats_doses_moyennes = {}
stats_doses_maximales = {}

for organe in organes:
    df_mean_dose = df[(df[column_df_1].str.contains('Mean Dose', na=False, case=False)) & 
                      (df[column_df_2].str.contains(organe, na=False, case=False))]
    
    df_max_dose = df[(df[column_df_1].str.contains('Max Dose', na=False, case=False)) & 
                     (df[column_df_2].str.contains(organe, na=False, case=False))]
    
    stats_mean = {
        'Moyenne': df_mean_dose[Value].mean(),
        'Médiane': df_mean_dose[Value].median(),
        'Minimum': df_mean_dose[Value].min(),
        'Maximum': df_mean_dose[Value].max(),
        'Dispersion': df_mean_dose[Value].std()
    }
    
    stats_max = {
        'Moyenne': df_max_dose[Value].mean(),
        'Médiane': df_max_dose[Value].median(),
        'Minimum': df_max_dose[Value].min(),
        'Maximum': df_max_dose[Value].max(),
        'Dispersion': df_max_dose[Value].std()
    }
    
    stats_doses_moyennes[organe] = stats_mean
    stats_doses_maximales[organe] = stats_max

df_stats_doses_moyennes = pd.DataFrame(stats_doses_moyennes).T
df_stats_doses_maximales = pd.DataFrame(stats_doses_maximales).T

data = pd.DataFrame({
    'Dose moyenne': np.random.normal(loc=0, scale=1, size=100),
    'Dose Max': np.random.normal(loc=1, scale=2, size=100),
})

# Créer un box plot
for i, organe in enumerate(organes):
    for technique in techniques:
        plt.figure(figsize=(20, 14))

        # Filtrer les données pour les doses moyennes et maximales
        df_mean_dose = df[(df[column_df_1].str.contains('Mean Dose', na=False, case=False)) & 
                          (df[column_df_2].str.contains(organe, na=False, case=False)) &
                          (df[column_df_3].str.contains(technique, na=False, case=False))]
        
        df_max_dose = df[(df[column_df_1].str.contains('Max Dose', na=False, case=False)) & 
                         (df[column_df_2].str.contains(organe, na=False, case=False))&
                         (df[column_df_3].str.contains(technique, na=False, case=False))]
        
        sns.boxplot(x='Plan', y='Valeur', data=df_mean_dose, palette='coolwarm', orient='v')
        plt.title(f'{organe} - Mean Dose en {technique}')
        plt.xlabel('',fontsize = 24)
        plt.ylabel('Dose',fontsize = 24)
        plt.xticks(fontsize=24)
        plt.yticks(fontsize=24)
        
        data_to_plot = [df_mean_dose['Valeur'], df_max_dose['Valeur']]
        sns.boxplot(data=data_to_plot,orient='v')
        
        
        moyenne_mean = df_mean_dose['Valeur'].mean()
        max_mean = df_mean_dose['Valeur'].max()
        
        moyenne_max = df_max_dose['Valeur'].mean()
        max_max = df_max_dose['Valeur'].max()
        
        plt.scatter([0], [moyenne_mean], color='r', zorder=5, label=f'Moyenne Mean Dose: {moyenne_mean:.2f}', s=100, edgecolor='k')
        plt.scatter([1], [moyenne_max], color='b', zorder=5, label=f'Moyenne Max Dose: {moyenne_max:.2f}', s=100, edgecolor='k')

        plt.title(f'Doses pour {organe} en {technique}')
        plt.ylabel('Dose')
        plt.ylabel('Moyenne [gauche] et Max [droite]')
        plt.legend(fontsize=18)
        plt.grid(True, which='both', linestyle='--', linewidth=0.9)
        plt.minorticks_on()
        plt.savefig(r'B:\RADIOTHERAPIE\Physique\2-Projets en cours\IMRT sein\Etude\Comparaison des doses 3D vs IMRT\Comparaison_V2\SaveFig\Figure_'+organe +'_'+technique+'.png')

        plt.figure(figsize=(20, 14))
        plt.title(f'Distribution pour {organe} en {technique}')
        plt.ylabel('Dose',fontsize = 24)
        plt.xlabel('Valeur',fontsize = 24)
        plt.xticks(fontsize=20)
        plt.yticks(fontsize=20)
        plt.legend(fontsize=18)
        plt.grid(True, which='both', linestyle='--', linewidth=0.9)
        plt.minorticks_on()
        x_mean = np.arange(1, len(df_mean_dose) + 1)
        x_max = np.arange(1, len(df_max_dose) + 1)
        plt.scatter(x_mean,df_mean_dose['Valeur'].sort_values( ascending=True),)
        plt.scatter(x_max,df_max_dose['Valeur'].sort_values( ascending=True))
        plt.savefig(r'B:\RADIOTHERAPIE\Physique\2-Projets en cours\IMRT sein\Etude\Comparaison des doses 3D vs IMRT\Comparaison_V2\SaveFig\Distribution_'+organe +'_'+technique+'.png')
 
# Ajuster l'espacement entre les sous-graphes
plt.tight_layout()


