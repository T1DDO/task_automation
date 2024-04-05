import pandas as pd

def lire_fichier_excel():
    ''' 
    Cette fonction lit la 2e colonne du fichier Excel donné et sélectionne
    les lignes à partir de la 15ème ligne.'''
    ds =pd.Series()
    # Charger seulement la 2e colonne du fichier
    ds =pd.read_excel('2015039834 COFELY HOLDING 7 LORA.xlsm', usecols=[1], header=None)
    # la sélection de la colonne 2 et tout les lignes à partir de la 15eme ligne
    ds_b = ds.iloc[14:]
    return ds_b

def traiter_donnees(ds_B15):
    ''' 
    Cette fonction divise les chaînes de la Series donnée en utilisant str.split('|', expand=True),
    elle ne conserve que les trois dernières colonnes de ce DataFrame.
    '''
    # Création d'un nouveau df(finale)
    df_fn = pd.DataFrame()
    split_data = ds_B15[1].str.split('|', expand=True)
    # Sélectionner les trois dernières colonnes de cette séparation
    df_fn = split_data.iloc[:, -3:]
    # Ajouter des noms pour les colonnes du df final
    df_fn.columns = ['DEVEUI', 'APPEUI', 'APPKEY']
    return df_fn

def sauvegarder_dataframe(df_fn):
    '''
    Sauvegarde le DataFrame dans un fichier Excel, en spécifiant le chemin du fichier de sortie.
    '''
    # Sauvegarder le DataFrame dans un fichier Excel
    file_path = 'nv_doc.xlsx'
    df_fn.to_excel(file_path, index=False)
    print("Le fichier a été sauvegardé à :", file_path)

ds_B15 = lire_fichier_excel()
df_final = traiter_donnees(ds_B15)
sauvegarder_dataframe(df_final)
