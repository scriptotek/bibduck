BIBDUCK
=============

BIBDUCK er et generelt JavaScript-bibliotek for automatisert kommunikasjon med BIBSYS ved hjelp av SecureNetTerms 
ActiveX-motorer, samt et HTA-basert brukergrensesnitt.

BIBDUCK-grensesnittet har en modulbasert tilnærming, og laster automatisk inn moduler fra mappen `plugins`.

- [Moduler](#moduler)
- [Hvordan bruke BIBDUCK-grensesnittet?](#hvordan-bruke-bibduck-grensesnittet)
- [Merknader](#merknader)


Moduler
------------

* `rfid.js` kaller et eksternt RFID-kontrollprogram basert på hvilken BIBSYS-skjerm som vises i 
  SecureNetTerm-vinduet som til enhver tid har fokus, slik at RFID-platen deaktiverer alarmen ved utlån
  (REG-skjerm), aktiverer den ved retur (RET-skjerm), og leser brikker ved behov  (f.eks. DOKST-skjerm). 
  Lesing skrus av når det ikke er nødvendig for å unngå at man kommer borti med en bok. For en oversikt
  over hvilke skjermer som er støttet, se [rfid.js](/scriptotek/bibduck/blob/master/plugins/rfid.js#L133).
  Vi har bare testet dette med RFID-platene fra Bibliotheca.

* `general.js` aksepterer dokid i forfatter-feltet (det øverste) på DOKST-skjermen, slik at man 
  slipper å tabbe ned til dokid-feltet. Modulen definerer også kommandoen `!!` for å tømme et felt
  man holder på å skrive i.

* `stikksedler.js` definerer kommandoen `stikkseddel!` for å skrive ut stikkseddel. Scriptet benytter seg 
  av den generelle konfigurasjonsfilen `stikksedler/config.json` og bibliotekspesifikk konfigurasjonsfil
  og maler (f.eks. `stikksedler/ureal.js` og `stikksedler/ureal/`) basert på bibliotekskoden man 
  angir i BIBDUCK-innstillingene.

* `imott-iret-auto.js` er en automatiseringsprosedyre for å registrere utlån, skrive ut stikkseddel 
  og sende hentebeskjed for mottatte *lån*. Den gjør ikke noe med mottatte *kopier*.
  Modulen er under uttesting.

* `loggers.js` logger utlån, retur og dokst-besøk i loggvinduet, slik at man kan finne tilbake til 
  informasjon man "mister" fra BIBSYS-vinduet etter at en bruker har forlatt skranken. 
  Informasjonen lagres ikke til disk, og forsvinner når man logger BIBDUCK eller trykker på knappen "Tøm logg".
  Modulen definerer kommandoene `dokid!` og `ltid!` for å skrive ut hhv. siste dokid og ltid.

* `lstatus.js` er en eksperimentell modul for å skrive ut en liste over alle brukerens lån. Dette gjøres med
  kommandoen `print!` på LTSTatus-skjermen. Modulen benytter Excel-arket `ltstatus.xls`.

* `dualbib.js` er en svært eksperimentell modul for å arbeide med to vinduer samtidig. Modulen definerer
  kommandoen `n!` for å hoppe mellom vinduer og `skr,*x*!` for å vise resultat *x* fra et BIB-søk i 
  et annet vindu.

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

Merknader
-------------
* RFID-kontrollerprogramvaren fra Bibliotheca håndterer ikke modus-endringer mens det 
ligger en (eller flere) bøker på platen. Det går derfor ikke an å f.eks. deaktivere og deretter aktivere
alarmen mens en bok ligger på platen. 
