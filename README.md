BIBDUCK
=============

BIBDUCK har vokst ut fra en idé om at det måtte være mulig å lage et script som automatisk satte RFID-platene 
fra Bibliotheca i riktig modus ift. det man holder på med i BIBSYS, f.eks. deaktivere alarm ved utlån, aktivere ved retur, 
og skru av lesing når det ikke er nødvendig for å unngå at man kommer borti med en bok – altså få BIBSYS til å snakke
med RFID-platene.

- [Hva er BIBDUCK?](#hva-er-bibduck)
- [Hvordan bruke BIBDUCK-grensesnittet?](#hvordan-bruke-bibduck-grensesnittet)
- [Tips](#tips)

Hva er BIBDUCK?
-------------
* BIBDUCK er et generelt JavaScript-basert bibliotek for kommunikasjon med BIBSYS ved hjelp av SecureNetTerms ActiveX-motorer, samt et HTA-basert brukergrensesnitt.

* Foreløpig kan BIBDUCK brukes til å kalle et RFID-kontrollprogram basert på hvilken BIBSYS-skjerm som vises i SecureNetTerm-vinduet som har fokus.

* JavaScript-biblioteket åpner for å lage hendelsesbaserte makroer.

Hvordan bruke BIBDUCK-grensesnittet?
-------------
Lukk eventuelle åpne BIBSYS-vinduer før du starter. BIBDUCK kan kun kommunisere med BIBSYS-instanser startet fra BIBDUCK.
Start deretter BIBDUCK fra ikonet på skrivebordet:

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

Merk
-------------
* RFID-kontrollerprogramvaren fra Bibliotheca håndterer ikke modus-endringer mens det 
ligger en (eller flere) bøker på platen. Det går derfor ikke an å f.eks. deaktivere og deretter aktivere
alarmen mens en bok ligger på platen. 
