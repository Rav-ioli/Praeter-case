import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

#Verkrijg gehele dataframe
unfiltered_df = pd.read_excel("PraeterBV_Case.xlsx", sheet_name=1, skiprows=14, usecols="A:N")
unfiltered_df.rename(columns={'CATEGORIE': 'Utiliteitsbouw dienstensector'}, inplace=True)
#Verkrijg de sectoren
sectorNamen_df = pd.read_excel("PraeterBV_Case.xlsx", sheet_name=1, usecols="R")
sectorNamen_df = sectorNamen_df.iloc[0:18]
sectorNamen_df.columns = ["Sectoren"]
sectorNamen_list = sectorNamen_df["Sectoren"].dropna().tolist()
#Verkrijg de inputvelden
inputWaardes_df = pd.read_excel("PraeterBV_Case.xlsx", sheet_name=0, skiprows=4, usecols="C:D")
inputWaardes_df.set_index('Gegevens', inplace=True)
inputWaardes_df.columns = ['Waarde']

gasVerbruik = int(inputWaardes_df.loc['Gas', 'Waarde'])
elektriciteitsVerbruik = int(inputWaardes_df.loc['elektriciteit', 'Waarde'])
energetischeWaardeGE = int(inputWaardes_df.loc['energetische waarde gas-elektra', 'Waarde'])
bouwjaarPand = int(inputWaardes_df.loc['Bouwjaar', 'Waarde'])
inputCategorie = inputWaardes_df.loc['Categorie', 'Waarde']
oppervlaktePand = int(inputWaardes_df.loc['Oppervlakte', 'Waarde'])
hoogteEtage = float(inputWaardes_df.loc['Hoogte (etage)', 'Waarde'])
#Sector index op basis van keuze
gekozenSectorIndex = 0
if inputCategorie in sectorNamen_list:
    gekozenSectorIndex = sectorNamen_list.index(inputCategorie)
else:
    gekozenSectorIndex = len(sectorNamen_list) - 1

#Hier wordt de categorie berekend op basis van de oppervlakte
def BerekenCategorieOppervlakte(oppervlakte):
    dataCategorieBereik_df = pd.read_excel("PraeterBV_Case.xlsx", sheet_name=1, usecols="W:Y"  )
    dataCategorieBereik_df = dataCategorieBereik_df.iloc[0:5]
    dataCategorieBereik_df = dataCategorieBereik_df.astype(int)
    dataCategorieBereik_df.columns = ["Min", "Max", "Categorie"]
    #Controleer de bereiken
    row = dataCategorieBereik_df[(dataCategorieBereik_df['Min'] <= oppervlakte) & (dataCategorieBereik_df['Max'] >= oppervlakte)]
    if not row.empty:
        return row['Categorie'].values[0]
    #Als waarde hoger is dan getal, geef dan hoogste categorie
    if oppervlakte > dataCategorieBereik_df['Max'].max():
        return dataCategorieBereik_df['Categorie'].max()  

categorieOppervlakte = BerekenCategorieOppervlakte(oppervlaktePand)

#De DataFrame wordt eerst gefiltered op basis van sector, dan op bouwjaar en als laatst op oppervlakte categorie
def FilterDataframe(sectorNaam, bouwjaar, categorieOppervlakte):    
    filteredOpDS = unfiltered_df[unfiltered_df['Utiliteitsbouw dienstensector'] == sectorNaam].copy()
    filteredOpDS.rename(columns={'Onderwerp': 'MIN', 'Unnamed: 2': 'MAX', 'Unnamed: 3': 'CAT'}, inplace=True)

    filteredOpBouwjaar = filteredOpDS[(filteredOpDS['MIN'] <= bouwjaar) & (filteredOpDS['MAX'] >= bouwjaar)]

    resultaatGasElektriciteit = filteredOpBouwjaar[[categorieOppervlakte, f"{categorieOppervlakte}.1"]].astype(float)
    return resultaatGasElektriciteit.values.tolist()

resultaatGasElektriciteit_list = FilterDataframe( sectorNamen_list[gekozenSectorIndex], bouwjaarPand, categorieOppervlakte)

#Hier wordt de tabel gegenereerd. Er wordt eerst gekeken of het gaat om de eigen waardes of om de gemiddelde waardes
def GenereerTabel(gasVerbruik, elektriciteitsVerbruik, energetischeWaardeGE, oppervlaktePand, hoogteEtage, gemiddeldeWaardes=None):
     # Genereer 3x3 dataframe met benaming van kolommen en rijen
    genereerdeTabel = pd.DataFrame(np.empty((3,3), dtype=float), columns=["Per m2", "Per m3", "Totaal"], index=["Gas (m3)", "Elektriciteit (kWh)", "Totaal"])
    
    if gemiddeldeWaardes is None:
        genereerdeTabel.iloc[0, 0] = gasVerbruik / oppervlaktePand  # Gas per m2
        genereerdeTabel.iloc[1, 0] = elektriciteitsVerbruik / oppervlaktePand  # Elektriciteit per m2
        genereerdeTabel.iloc[0, 1] = gasVerbruik / (oppervlaktePand * hoogteEtage)  # Gas per m3
        genereerdeTabel.iloc[1, 1] = elektriciteitsVerbruik / (oppervlaktePand * hoogteEtage)  # Elektriciteit per m3
    else:
        genereerdeTabel.iloc[0, 0] = gemiddeldeWaardes[0][0]
        genereerdeTabel.iloc[1, 0] = gemiddeldeWaardes[0][1]
        genereerdeTabel.iloc[0, 1] = gemiddeldeWaardes[0][0] / 2.7
        genereerdeTabel.iloc[1, 1] = gemiddeldeWaardes[0][1] / 2.7

    genereerdeTabel.iloc[0, 2] = genereerdeTabel.iloc[0, 0] * oppervlaktePand
    genereerdeTabel.iloc[1, 2] = genereerdeTabel.iloc[1, 0] * oppervlaktePand
    genereerdeTabel.iloc[2, 0] = genereerdeTabel.iloc[0, 0] * energetischeWaardeGE + genereerdeTabel.iloc[1, 0]
    genereerdeTabel.iloc[2, 1] = genereerdeTabel.iloc[0, 1] * energetischeWaardeGE + genereerdeTabel.iloc[1, 1]
    genereerdeTabel.iloc[2, 2] = genereerdeTabel.iloc[0, 2] * energetischeWaardeGE + genereerdeTabel.iloc[1, 2]
    return genereerdeTabel.round(2)

huidigVerbruik = GenereerTabel(gasVerbruik, elektriciteitsVerbruik, energetischeWaardeGE, oppervlaktePand, hoogteEtage, None)
gemiddeldVerbruik = GenereerTabel(gasVerbruik, elektriciteitsVerbruik, energetischeWaardeGE, oppervlaktePand, hoogteEtage, resultaatGasElektriciteit_list)
print(f'Totaal kWh (factor): {elektriciteitsVerbruik*energetischeWaardeGE}')

#Plotten
huidigeWaardes = [float(huidigVerbruik.iloc[2, 0]), float(huidigVerbruik.iloc[2, 1])]
gemiddeldeWaardes = [float(gemiddeldVerbruik.iloc[2, 0]), float(gemiddeldVerbruik.iloc[2, 1])]
labels = ["Per m2", "Per m3"]
x = np.arange(len(labels))
barBreedte = 0.3
huidigBar = plt.bar(x - barBreedte/2, huidigeWaardes, width=barBreedte, label='Huidig')
gemiddeldBar = plt.bar(x + barBreedte/2, gemiddeldeWaardes, width=barBreedte, label='Gemiddeld')
plt.title("Huidig verbruik t.o.v. gemiddeld verbruik (kWh)")
plt.xticks(x, labels)
plt.grid(axis='y', linestyle='--', alpha=0.5)
for bars in [huidigBar,gemiddeldBar]:
    for bar in bars:
        max = bar.get_height()
        x_pos = bar.get_x() + bar.get_width()/2
        plt.text(x_pos, bar.get_height(), max, ha='center')
plt.legend()
plt.show()