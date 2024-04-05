import pandas as pd

# Charger seulement la 2e colonne du fichier
df = pd.read_excel('2015039834 COFELY HOLDING 7 LORA.xlsm', usecols=[1], header=None)

# la sélection de la colonne 2 et tout les lignes à partir de la 15eme ligne
df_B15 = df.iloc[14:]

# Création d'un nouveau df(finale)
df_fn = pd.DataFrame()

# Parcourir toutes lignes 
for index, ligne in df_B15.iterrows():
    # Découper chaque ligne à chaque fois qu'on trouve le symbole "|"
    colonnes = ligne[1].split('|')
    # Ajouter les données au temp_df
    temp_df = pd.DataFrame([colonnes[-3:]]) # convertie la liste en DataFrame
    df_fn = pd.concat([df_fn, temp_df], ignore_index=True)

# Ajouter des noms pour les colonnes du df final
df_fn.columns = ['DEVEUI', 'APPEUI', 'APPKEY']

# Sauvegarder le DataFrame dans un fichier Excel
file_path = 'nv_doc.xlsx'
df_fn.to_excel(file_path, index=False)

print("Le fichier a été sauvegardé à :", file_path)
