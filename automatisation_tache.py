import pandas as pd

ds =pd.Series()
# Charger seulement la 2e colonne du fichier
ds =pd.read_excel('2015039834 COFELY HOLDING 7 LORA.xlsm', usecols=[1], header=None)

# la sélection de la colonne 2 et tout les lignes à partir de la 15eme ligne
ds_B15 = ds.iloc[14:]

# Création d'un nouveau df(finale)
df_fn = pd.DataFrame()

split_data = ds_B15[1].str.split('|', expand=True)

# Sélectionner les trois dernières colonnes de cette séparation
df_fn = split_data.iloc[:, -3:]

# Ajouter des noms pour les colonnes du df final
df_fn.columns = ['DEVEUI', 'APPEUI', 'APPKEY']

# Sauvegarder le DataFrame dans un fichier Excel
file_path = 'nv_doc2.xlsx'
df_fn.to_excel(file_path, index=False)

print("Le fichier a été sauvegardé à :", file_path)
