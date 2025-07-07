# Praeter case
 
pandas
openpyxl
numpy
numpy financial
matplotlib

pip install pandas
pip install openpyxl
pip install numpy
pip install numpy-financial
python -m pip install -U matplotlib

# Graag de meegeleverde excel sheet gebruiken, omdat de originele sheet paar fouten had

# Opdracht 1

Eerst worden de waardes uit het Excel-sheet opgehaald.
Daarna wordt de categorie bepaald op basis van de oppervlakte.
Als dat is gebeurd, wordt de DataFrame gefilterd: eerst op sectornaam, daarna op bouwjaar, en als laatste op de oppervlaktecategorie. Er worden dan twee waardes teruggegeven: gas- en elektriciteitsverbruik per m².

De functie genereerTabel wordt gebruikt om een tabel te genereren. Er wordt gekeken of de functie al een bestaande tabel krijgt (de tabel van de vorige functie); zo niet, dan worden eerst de waardes voor gas en elektriciteit berekend.
Na die berekening worden ook de overige waardes, zoals het totaalverbruik, ingevuld.

Ten slotte wordt er een plot gegenereerd.

Om deze script te kunnen uitvoeren, moeten eerst pandas, openpyxl, numpy en matplotlib zijn geïnstalleerd.
De gebruiker hoeft alleen de inputwaardes in te vullen in het Excel-sheet bij opdracht_1_voorbeeld.
Let op: bij de categorie moet goed opgelet worden op spelfouten.

# Opdracht 2 + extra opdracht

Als eerste worden de stamwaardes en daarna de inputwaardes opgehaald.

De totale investeringen, afschrijvingen en aflossingen worden vervolgens berekend met behulp van de input.

Daarna wordt er een tabel gemaakt voor het financiële overzicht.

Voor besparing, subsidie en vaste exploitatiekosten wordt de functie exponentieleGroei aangeroepen.
Deze functie berekent de exponentiële groei met de formule:
1 * (1 + inflatie) ^ exponent
en vult de waardes in voor elk jaar behalve jaar 0.

Vervolgens wordt de EBITDA berekend en worden ook de afschrijvingskosten ingevuld.

Daarna volgen de berekeningen voor financieringskosten, winst na belasting, cashflow IRR, cashflow REV, en ten slotte de TVT (terugverdientijd), zodra de cumulatieve cashflow REV positief is.

Ook worden de uiteindelijke waardes voor IRR, REV, winst na belasting en TVT berekend.

Om deze script uit te voeren moet de gebruiker de inputwaardes invullen in het Excel-sheet bij opdracht_2_input_output.
Daarnaast moeten de pakketten pandas, openpyxl, numpy en numpy-financial geïnstalleerd zijn.
