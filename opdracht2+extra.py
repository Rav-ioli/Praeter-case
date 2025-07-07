import pandas as pd
import numpy as np
import numpy_financial as npf

algemeneWaardes_df = pd.read_excel("PraeterBV_Case.xlsx", sheet_name=2, skiprows=3, usecols="B:C")[:5]
algemeneWaardes_df.set_index('Algemeen', inplace=True)
algemeneWaardes_df.columns = ['Waarde']

inputWaardes_df = pd.read_excel("PraeterBV_Case.xlsx", sheet_name=2, skiprows=11, usecols="B:C")[:15]
inputWaardes_df.set_index('Infrarood', inplace=True)
inputWaardes_df.columns = ['Waarde']

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

rijNamen = [
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

outputTabel = pd.DataFrame(np.nan, index=rijNamen, columns=[int(i) for i in range(termijnInJaar + 1)])

for row in outputTabel.index:
    if row == "---":
        outputTabel.loc[row, :] = "--------"

outputTabel.index.name = "Toepassing"
outputTabel.columns.name = "Jaar"

outputTabel.loc['Eenmalige kosten', outputTabel.columns[0]] = -eenmaligeKosten
outputTabel.loc['Eenmalige kosten', outputTabel.columns[1:]] = 0
outputTabel.loc['Herinvestering', herinvesteringJaar] = -herinvestering

def exponentieleGroei(naam, waarde):
    for jaar in outputTabel.columns:
        if int(jaar) == 0:
            continue
        exponent = int(jaar) - 1
        groei = waarde * (1 + inflatie) ** exponent
        outputTabel.loc[naam, jaar] = groei

exponentieleGroei("Besparing", besparing)
exponentieleGroei("Subsidie (jaarlijks)", subsidieJaarlijks)
exponentieleGroei("Vaste exploitatiekosten", -vasteExploitatiekosten)

rijenVoorEBITDA = [
    "Besparing",
    "Subsidie (jaarlijks)",
    "Kosten",
    "Eenmalige kosten",
    "Vaste exploitatiekosten",
    "Herinvestering"
]

outputTabel.loc["EBITDA"] = outputTabel.loc[rijenVoorEBITDA].sum()
outputTabel.loc["Afschrijvingskosten", outputTabel.columns[1:]] = -afschrijvingBerekening

for jaar in outputTabel.columns:
        if int(jaar) == 0:
            continue
        vorigJaar = jaar -1
        financieringsKosten = (totaleInvestering*(1-eigenVermogen) - (aflossing*vorigJaar)) *renteVV
        outputTabel.loc["Financieringskosten", jaar] = -financieringsKosten
        
rijenVoorBelasting = [
    "EBITDA",
    "Afschrijvingskosten",
    "Financieringskosten"
]

outputTabel.loc["Belasting", 1:] = -outputTabel.loc[rijenVoorBelasting].sum() * belasting

rijenVoorBelasting.append("Belasting")

outputTabel.loc["Winst na belasting"] = outputTabel.loc[rijenVoorBelasting].sum()

outputTabel.loc["Cashflow IRR", 1:] = outputTabel.loc[["EBITDA","Belasting"]].sum()


outputTabel.loc["Cashflow IRR", 0] = outputTabel.loc["Winst na belasting", 0] + totaleInvestering * -1

outputTabel.loc["Cashflow REV", 0] = (totaleInvestering * eigenVermogen - outputTabel.loc["Winst na belasting",0] ) * -1

outputTabel.loc["Cashflow REV", 1:] = outputTabel.loc["EBITDA"] + outputTabel.loc["Financieringskosten"] + outputTabel.loc["Belasting"] - aflossing

for col in outputTabel.columns:
    kolomIndex = outputTabel.columns.get_loc(col)
    vorigeJaren = outputTabel.columns[:kolomIndex + 1]
    totaalVorigeJaren = outputTabel.loc["Cashflow REV",  vorigeJaren].sum()

    if totaalVorigeJaren > 0:
        irr = outputTabel.loc["Cashflow IRR", col]
        rev = outputTabel.loc["Cashflow REV", col]
        tvt = kolomIndex - (rev / irr)
        outputTabel.loc["TVT", col] = tvt
#Afronden op 2 decimalen
outputTabel = outputTabel.applymap(lambda x: round(x, 2) if isinstance(x, (int, float)) else x)

irr = npf.irr(outputTabel.loc["Cashflow IRR"].values)
rev = npf.irr(outputTabel.loc["Cashflow REV"].values)
winstNaBelasting = outputTabel.loc["Winst na belasting"].sum()
tvt = outputTabel.loc["TVT"].dropna().min()

outputTabel.fillna("", inplace=True)

print(outputTabel)
print("")
print(f"IRR: {round(irr*100,2)}%")
print(f"REV: {round(rev*100,2)}%")
print(f"PAT: {round(winstNaBelasting,3)}")
print(f"TVT (Jaar): {tvt}")