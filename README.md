# SPIRA // CC BY-NC-SA 4.0 // RAZU
Script voor Pre-Ingest Regionaal Archief Zuid-Utrecht
SPIRA is een Powershell-script voor manipuleren van XMLs welke door de Exporttool van VHIC gegenereerd zijn aan de hand van de originele JSON-bestanden van de export van Zaaksysteem.nl 
SPIRA voorziet hiermee in de behoefte om tijdens de 'pre-ingest', de limbo tussen de export van het bronsysteem (Zaaksysteem.nl) en de import in het doelsysteem (De REE Archiefsystemen), modificaties te maken aan een afgeleide van de originele metadata. 
SPIRA is ontwikkeld door het Regionaal Archief Zuid-Utrecht (RAZU) en beschikbaar onder de Creative Commons Naamsvermelding-NietCommercieel-GelijkDelen 4.0 Internationaal-licentie (**CC BY-NC-SA 4.0**). Zie https://creativecommons.org/licenses/by-nc-sa/4.0/deed.nl
Ondersteuning voor gebruik door derden wordt niet aangeboden. Gebruik is op eigen kracht en eigen risico. 

SPIRA is ontwikkeld met dank aan PoshGUI.com voor het leveren van de tools om simpel een XAML-GUI te maken. Alle overige code is door het RAZU zelf ontwikkeld, uiteraard met gebruik van voorbeelden via Google gevonden. 

SPIRA versies:
============================================================================================
versies 1 t/m 3 heetten DnTSplitser, vernoemd naar de originele gewenste functionaliteit voor het splitsen van de datum-tijd velden in de XML in twee velden, datum en tijd.

**Versie 1.3**
Gedurende de ontwikkeling van DnTSplitser is de naam gewijzigd naar SPIRA. De eerste versie van SPIRA is derhalve 1.3

**Versie 1.4**
Ontwikkeling aan 1.4 startte snel na 1.3 en was met name gericht op het verder moduleren van het script. Het script werd erg lang, en hoewel dit niet veel schade deed zorgde het voor problemen met het lezen en zoeken van onderdelen en, belangrijker, gaf het geen ruimte aan de gebruiker om eigen extra velden toe te voegen.
De ontwikkeling van 1.4 is halverwege gestaakt met de opstart van versie 1.5

*Ontwikkelde functionaliteiten van versie 1.4 voor overgang naar 1.5*
	Powershell
		- Overgestapt op Powershell 7 (originele script geschreven in PS 5)
	Modulariteit:
		- alle afzonderlijke onderdelen als functies in het script gebouwd. 
		- User-input van losse vragen omgezet naar de SPIRA-opties xml.
		- Opties uit de XML worden nu gebruikt om te bepalen welke onderdelen worden uitgevoerd, en waar een bestand wordt opgeslagen en onder welke naam. (was hard-coded in de vorige versie)
	Nieuwe functies:
		- Function: Opties Inladen: laadt de opties-XML in en toont de opties aan de gebruiker. Geeft tevens foutmeldingen als de opties verkeerd zijn ingevuld.
		- Function: Tobeornottobe: Aangezien opties nu automatisch worden ingeladen, ipv individueel worden gevraagd, geeft Tobeornottobe de gebruiker de mogelijkheid om het script te stoppen als een optie niet klopt. Wordt tevens gebruikt voor andere bijzonderheden.
	Vervallen functies:
		- Function: Overschrijven -> vervallen door invoering van de opties-XML
		- Function: alt-pad -> Vervallen door invoering van de opties-XML
	Diversen:
		- Write-Host teksten opgeschoond en uniform gemaakt
		- Kleurcoderingen toegepast op de output, zodat de output sneller te begrijpen / lezen is.
		- Comments geschoond, grootste deel van de uitleg verplaatst naar de opties-XML
		- Script geplaatst onder de CC BY-NC-SA 4.0-Internationaal licentie
		- Gebruikers krijgen nu een foutmelding als er geen zaken zijn gevonden. 


**Versie 1.5**
Versie 1.5 werd opgestart tijdens de ontwikkeling van 1.4 en heeft er toe geleidt dat 1.4 achtergelaten werd. De reden om te starten aan 1.5 was de opname van een GUI (graphical user interface), of te wel een gebruikersvriendelijk venster met knopjes en vinkjes, i.p.v. een console met tekstuele instructies en input via XMLs. De ontwikkeling startte met het maken van een simpele interface op PoshGUI.com, en deze te testen in Visual Studio Code. Nadat dit naar tevredenheid behaald was is de code van SPIRA 1.4 over geheveld, waarbij de toen nog in ontwikkeling zijnde modulariteit direct is meegenomen. Voorbeeld: in SPIRA 1.3 stond 7 keer dezelfde 13 regels code achtereen waarin de verschillende standaard losse velden werden gewijzigd. In 1.5 is deze sectie van 13 regels 1x opgenomen als functie en wordt het 7 keer aangeroepen. Een stuk leaner, maar ook direct een stuk herbruikbaarder, mochten er extra velden moeten worden gesplitst.

*Nieuw in deze versie:*
- GUI! SPIRA start nu op als venster waarin de gebruiker simpel en snel zijn weg kan vinden. Dit zorgt er voor dat SPIRA als het goed is makkelijker te gebruiken is voor gebruikers die niet gewend zijn aan consoles. 
- Modulariteit is volledig doorgevoerd. Zodra de gebruiker op de start-knop klikt verloopt het proces als volgt: 
    0. Eerst wordt gecontroleerd of de standaardwaarden-xml aanwezig is (throw-able) en wordt de XAML-window ingelezen en verwerkt. Als dit succesvol is wordt het venster opgeroepen en krijgt de gebruiker de mogelijkheid om SPIRA in te stellen. Zodra de gebruiker op de start-knop klikt start Voorman.
    1. Voorman (functie) is verantwoordlijk voor het aansturen van de andere functies en het bijwerken van de GUI. Het roept eerst de functie Opstarten op, welke de standaard waarden uit de StWaarden-xml leest en alvast een aantal variabelen declareert. 
    2. Voorman start de functie BestandenZoeken, welke op de door de gebruiker aangewezen export-locatie alle zaak- en document-XMLs gaat indexeren
    3. Voorman start vervolgens op basis van de teller van BestandenZoeker (aantal_zaken en aantal_docs) de functie StappenDoorlopen net zo vaak op als het aantal zaken en documenten gevonden zijn.
    4. De functie StappenDoorlopen bepaalt eerst of het de bedoeling is dat datum-splits worden toegepast. Vervolgens doorloopt het 4 stappen:
      1). Events: op dt moment worden hier enkel de datum-velden van events gesplitst. StappenDoorlopen roept hier Datum_aanpassen_standaard op.
      2). Relaties: Indien het een zaak-document betreft (dit geeft de Voorman mee) dan start StappenDoorlopen hier de functie Relaties_verwerken
      3). Losse velden: dit is de eerder genoemde mulit-call. Datum_aanpassen_standaard wordt hier in totaal 11 keer opgeroepen, 3x als het een document betreft, 7x als het een zaak betreft en 1x om het overzicht van relaties weg te schrijven (zie hierna)
      4). XML opslaan: zodra alle verwerkingen op het bestand zijn uitgevoerd (stappen I-III) dan worden hier vervolgens de bestandsnaam en de opslaglocatie bepaald en het bestand daadwerkelijk opgeslagen.
    5. De functie Datum_aanpassen_standaard leest de waarde uit van het door StappenDoorlopen opgegeven veld en splitst dit vervolgens in 2 nieuwe nodes, die aan het einde van de meegegeven node worden geplakt. Voor events is dat in het event, voor relaties in de relatie, voor de rest is dat aan het einde van de values-node.
      Datum_aanpassen kan, zoals aangekondigd, ook de volledige relaties plakken als #### (relatietype),... Dit gebeurd als Relatie koppelen is aangezet. Het wordt hier opgeroepen omdat dit pas kan worden uitgevoerd als het script in de values-node is aanbeland en alle relaties heeft verwerkt. (kan technisch gezien wel eerder, maar dat had tot meer code geleidt)
     6. De functie Relaties_verwerken vertaalt (indien aangevinkt) en koppelt (indien aangevinkt) de relaties. Dit leidt tot een extra node in elke relatie-node met de opmaak #### (relatietype). De vertaling wordt door Opstarten uit het standaardwaarden-xmlbestand gehaald. 
     Relaties_verwerken plaatst een verwerkte relatie ook direct in related_cases_volledig, zodat deze door Datum_aanpassen_standaard uiteindelijk aan het einde van de values-node kan worden geplakt. 
     7. De Voorman eindigt vervolgens het draaiende script en werkt de GUI bij. De gebruiker wordt verzocht om het venster te sluiten.
     8. Alle overige functies zijn verbonden aan de werking van de GUI. 

Toekomstige ontwikkelingen
============================================================================================

2020:
- Vertaling van documentlabels

Onbekend:
- aangepaste waarden en datum-velden
