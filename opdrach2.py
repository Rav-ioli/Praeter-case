import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

algemeneWaardes_df = pd.read_excel("PraeterBV_Case.xlsx", sheet_name=2, skiprows=3, usecols="B:C")[:5]
algemeneWaardes_df.set_index('Algemeen', inplace=True)
algemeneWaardes_df.columns = ['Waarde']

inputWaardes_df = pd.read_excel("PraeterBV_Case.xlsx", sheet_name=2, skiprows=11, usecols="B:C")[:15]
inputWaardes_df.set_index('Infrarood', inplace=True)
inputWaardes_df.columns = ['Waarde']


# column_name = inputWaardes_df.loc["Herinvestering"][lambda x: x == 4000].index
inflatie = algemeneWaardes_df.loc['Inflatie', 'Waarde']
afschrijving = algemeneWaardes_df.loc['Afschrijving', 'Waarde']
eigenVermogen = algemeneWaardes_df.loc['Eigen Vermogen', 'Waarde']
renteVV = algemeneWaardes_df.loc['Rente VV', 'Waarde']
belasting = algemeneWaardes_df.loc['Belasting', 'Waarde']

termijnInJaar = int(inputWaardes_df.loc['Termijn','Waarde'])
herinvesteringJaar = int(inputWaardes_df.loc['HerinvesteringJaar','Waarde'])

investering = inputWaardes_df.loc['Investering', 'Waarde']
subsidieEenmalig = inputWaardes_df.loc['Subsidie (eenmalig)', 'Waarde']
restwaarde = inputWaardes_df.loc['Restwaarde', 'Waarde']

inkomsten = inputWaardes_df.loc['Inkomsten', 'Waarde']
besparing = inputWaardes_df.loc['Besparing', 'Waarde']
subsidieJaarlijks = inputWaardes_df.loc['Subsidie (jaarlijks)', 'Waarde']

kosten = inputWaardes_df.loc['Kosten', 'Waarde']
eenmaligeKosten = inputWaardes_df.loc['Eenmalige kosten', 'Waarde']
vasteExploitatiekosten = inputWaardes_df.loc['Vaste exploitatiekosten', 'Waarde']
herinvestering = inputWaardes_df.loc['Herinvestering', 'Waarde']

totaleInvestering = investering - subsidieEenmalig
afschrijvingBerekening = (investering - subsidieEenmalig - restwaarde) / termijnInJaar
aflossing = totaleInvestering * (1-eigenVermogen) / termijnInJaar

custom_index = [
    "Infrarood",
    "---",
    "Inkomsten",
    "Besparing",
    "Subsidie (jaarlijks)",
    "",
    "Kosten",
    "Eenmalige kosten",
    "Vaste exploitatiekosten",
    "Herinvestering",
    "---",
    "EBITDA",
    "Afschrijvingskosten",
    "Financieringskosten",
    "Belasting",
    "---",
    "Winst na belasting",
    "Cashflow IRR",
    "Cashflow REV",
    "TVT"
]
# column_names = [int(i) for i in range(termijnInJaar + 1)]
genereerdeTabel = pd.DataFrame(np.nan, index=custom_index, columns=[int(i) for i in range(termijnInJaar + 1)])
# genereerdeTabel = genereerdeTabel.astype(object)
# genereerdeTabel.loc[genereerdeTabel.index == '---', :] = '--------'
for row in genereerdeTabel.index:
    if row == "---":
        genereerdeTabel.loc[row, :] = "--------"

# Set a name for the index
genereerdeTabel.index.name = "Toepassing"
genereerdeTabel.columns.name = "Jaar"

eerste_kolom = genereerdeTabel.columns[0]
genereerdeTabel.loc['Eenmalige kosten', eerste_kolom] = -eenmaligeKosten

# Set all other columns to 0
andere_kolommen = genereerdeTabel.columns[1:]
genereerdeTabel.loc['Eenmalige kosten', andere_kolommen] = 0

genereerdeTabel.loc['Herinvestering', herinvesteringJaar] = -herinvestering

# for i in range(termijnInJaar):
#     jaar = int(i)  # e.g. column name '8' → 8
#     genereerdeTabel.loc["Besparing", i] = besparing * (1 + inflatie) ** (jaar - 1)

# for jaar in column_names:
#     if int(jaar) == 0:
#         continue  # skip jaar 0
#     exponent = int(jaar) - 1
#     waarde = besparing * (1 + inflatie) ** exponent
#     genereerdeTabel.loc["Besparing", jaar] = waarde


def vul_rij_met_inflatie(naam: str, waarde: float):
    for jaar in genereerdeTabel.columns:
        if int(jaar) == 0:
            continue
        exponent = int(jaar) - 1
        groei = waarde * (1 + inflatie) ** exponent
        genereerdeTabel.loc[naam, jaar] = groei

vul_rij_met_inflatie("Besparing", besparing)
vul_rij_met_inflatie("Subsidie (jaarlijks)", subsidieJaarlijks)
vul_rij_met_inflatie("Vaste exploitatiekosten", -vasteExploitatiekosten)

rijen_voor_ebitda = [
    "Besparing",
    "Subsidie (jaarlijks)",
    "Kosten",
    "Eenmalige kosten",
    "Vaste exploitatiekosten",
    "Herinvestering"
]

genereerdeTabel.loc["EBITDA"] = genereerdeTabel.loc[rijen_voor_ebitda].sum()
genereerdeTabel.loc["Afschrijvingskosten", genereerdeTabel.columns[1:]] = -afschrijvingBerekening

for jaar in genereerdeTabel.columns:
        if int(jaar) == 0:
            continue
        d = jaar -1
        a = (totaleInvestering*(1-eigenVermogen) - (aflossing*d)) *renteVV
        genereerdeTabel.loc["Financieringskosten", jaar] = -a
        
rijen_voor_belasting = [
    "EBITDA",
    "Afschrijvingskosten",
    "Financieringskosten"
]

genereerdeTabel.loc["Belasting", 1:] = -genereerdeTabel.loc[rijen_voor_belasting].sum() * belasting

rijen_voor_belasting.append("Belasting")

genereerdeTabel.loc["Winst na belasting"] = genereerdeTabel.loc[rijen_voor_belasting].sum()

genereerdeTabel.loc["Cashflow IRR", 1:] = genereerdeTabel.loc[["EBITDA","Belasting"]].sum()


genereerdeTabel.loc["Cashflow IRR", 0] = genereerdeTabel.loc["Winst na belasting", 0] + totaleInvestering * -1

genereerdeTabel.loc["Cashflow REV", 0] = (totaleInvestering * eigenVermogen - genereerdeTabel.loc["Winst na belasting",0] ) * -1

genereerdeTabel.loc["Cashflow REV", 1:] = genereerdeTabel.loc["EBITDA"] + genereerdeTabel.loc["Financieringskosten"] + genereerdeTabel.loc["Belasting"] - aflossing

print(genereerdeTabel)




# print(a)
# print(genereerdeTabel.columns.values)

# column_name = inputWaardes_df.loc["Herinvestering"][lambda x: x == 4000].index

# print(column_name)