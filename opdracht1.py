import pandas as pd

aaaa = pd.read_excel("PraeterBV_Case.xlsx", sheet_name=1, skiprows=14, usecols="A:N"  )
aaaa.rename(columns={'CATEGORIE': 'Utiliteitsbouw dienstensector'}, inplace=True)

sectoren = [
    "Detailhandel met koeling",
    "Detailhandel zonder koeling",
    "Groothandel zonder koeling",
    "Autobedrijf: showroom en garage",
    "Autobedrijf: autoschadeherstelbedrijven",
    "Horeca: café",
    "Horeca: restaurant",
    "Horeca: cafetaria",
    "Horeca: hotels, motels",
    "Kantoor: overheid",
    "Kantoor: overig",
    "Onderwijs: primair",
    "Gezondheidszorg: bijeenkomst",
    "Gezondheidszorg: praktijk",
    "Gezondheidszorg: tehuis",
    "Recreatie: binnensport",
    "Recreatie: buitensport",
    "Overig: religie"
]

# gasVerbruik = int(input("Wat is uw huidig gasverbruik in m3? "))
# elektriciteitsVerbruik = int(input("Wat is uw huidig elektriciteitsverbruik in kWh? "))
# energetischeWaardeGE = int(input("Wat is uw energetische waarde factor voor gas en elektra? "))
bouwjaarPand = int(input("Wat is het bouwjaar van uw pand? "))
sector_index = 0
for sector in sectoren:
    print(f"{sector_index}. {sector}")
    sector_index += 1
gekozenSectorIndex = int(input("Welke sector is voor u toepasselijk? "))
oppervlaktePand = int(input("Wat is het oppervlakte van uw pand? "))
# hoogteEtage = int(input("Wat is de hoogte van een etage?"))

#Hier 
def berekenCategorieOppervlakte(waarde):
    dataCategorie = {
    'min': [0, 251, 501, 1001, 2501],
    'max': [250, 500, 1000, 2500, 5000],
    'categorie': [1, 2, 3, 4, 5]
    }
    dfDataCategorie = pd.DataFrame(dataCategorie)

    row = dfDataCategorie[(dfDataCategorie['min'] <= waarde) & (dfDataCategorie['max'] >= waarde)]
    if not row.empty:
        return row['categorie'].values[0]
    #Als waarde hoger is dan getal, geef dan hoogste categorie
    if waarde > dfDataCategorie['max'].max():
        return dfDataCategorie['categorie'].max()  

categorieOppervlakte = berekenCategorieOppervlakte(oppervlaktePand)

def filterDataframe(Sector, Bouwjaar, categorieOppervlakte):    
    filteredDS = aaaa[aaaa['Utiliteitsbouw dienstensector'] == Sector].copy()
    print(filteredDS)
    filteredDS.rename(columns={'Onderwerp': 'MIN'}, inplace=True)
    filteredDS.rename(columns={'Unnamed: 2': 'MAX'}, inplace=True)
    filteredDS.rename(columns={'Unnamed: 3': 'CAT'}, inplace=True)

    filtered = filteredDS[(filteredDS['MIN'] <= Bouwjaar) & (filteredDS['MAX'] >= Bouwjaar)]
    print(filtered)
    fine = filtered[[categorieOppervlakte, f"{categorieOppervlakte}.1"]]
    print(fine)
   
    # return x

# print(aaaa)
filteredOpDienstenSector = filterDataframe( sectoren[gekozenSectorIndex], bouwjaarPand, categorieOppervlakte)
# print(filterDataframeOpDienstensector)



