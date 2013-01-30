BIBDUCK
=============

Hva er BIBDUCK?
-------------
* BIBDUCK er et generelt JavaScript-basert bibliotek for kommunikasjon med BIBSYS ved hjelp av SecureNetTerms ActiveX-motorer, samt et HTA-basert brukergrensesnitt.

* Foreløpig kan BIBDUCK brukes til å automatisk sette et RFID-kontrollprogram i riktig modus ("aktiver alarm", "deaktiver alarm" eller "kun lesing") basert på hvilken BIBSYS-skjerm som vises i SecureNetTerm-vinduet som har fokus.

Hvordan bruke BIBDUCK-grensesnittet?
-------------
BIBDUCK startes fra ikonet på skrivebordet:
![BIBDUCK icon](http://localhostr.com/file/CjlJkWeoyZCa/desktop-icon.jpg)

Man får da opp BIBDUCK-grensesnittet:
![BIBDUCK grensesnitt](http://localhostr.com/file/RS1x1zDwd9q4/interface.png)

Siden BIBDUCK er i utviklingsfasen, er grensesnittet tilrettelagt for testing, med et stort loggområde som 
viser tilbakemeldinger som kan være nyttige for feilsøking.
Foreløpig er det egentlig bare ett element du som bruker trenger å legge merke til: knappen **Nytt vindu**.
Trykker du på den, startes en ny BIBSYS-instans, der du logger inn som vanlig.
Som et eksempel på makrofunksjonalitet, leder BIBDUCK deg imidlertid automatisk gjennom 
innledningsskjermene frem til BIBSYS Søking.

Legg merke til at vinduet får tittelen "BIBSYS 1 - RFID: Skrudd av".
RFID-modusen vises i tittelen, slik at man raskt kan sjekke den selv om RFID-kontrollerprogrammet er minimert. 
I begynnelsen anbefales det imidlertid at man har RFID-kontrollerprogrammet oppe for å sjekke at BIBDUCK endrer modus på korrekt måte.

![BIBSYS](http://localhostr.com/file/YCCXADruUHV2/snetterm.png)

Prøv å gå til REG-skjermen, og legg merke til at RFID-kontrollerprogrammet endrer modus:
![RFID-kontroller](http://localhostr.com/file/YKCVBoZu9TZn/rfid.png)

Trykker du på **Nytt vindu** igjen, startes en ny instans "BIBSYS 2", osv... 
Vinduet som har fokus indikeres med gult i BIBDUCK. Et vindu får fokus når du skriver i det, eller
trykker i det blå området. Normalt trenger man ikke å tenke på fokus.
![BIBDUCK med flere vinduer](http://localhostr.com/file/cm8PMuVSrjRK/bibduck-multi.png)
