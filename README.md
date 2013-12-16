[![Stories in Ready](https://badge.waffle.io/scriptotek/bibduck.png?label=ready)](https://waffle.io/scriptotek/bibduck)  
BIBDUCK
=============

BIBDUCK er et generelt JavaScript-bibliotek for automatisert kommunikasjon med BIBSYS ved hjelp av SecureNetTerms 
ActiveX-motorer, samt et HTA-basert brukergrensesnitt.

BIBDUCK-grensesnittet har en modulbasert tilnC&rming, og laster automatisk inn moduler fra mappen `plugins`.

- [Standardmoduler](#standardmoduler)
- [Hvordan bruke BIBDUCK-grensesnittet?](#hvordan-bruke-bibduck-grensesnittet)
- [Merknader](#merknader-og-tips)


Standardmoduler
------------

* [rfid.js](plugins/rfid.js) kaller et eksternt RFID-kontrollprogram basert pC% hvilken BIBSYS-skjerm som vises i 
  SecureNetTerm-vinduet som til enhver tid har fokus, slik at RFID-platen deaktiverer alarmen ved utlC%n
  (REG-skjerm), aktiverer den ved retur (RET-skjerm), og leser brikker ved behov  (f.eks. DOKST-skjerm). 
  Lesing skrus av nC%r det ikke er nC8dvendig for C% unngC% at man kommer borti med en bok. For en oversikt
  over hvilke skjermer som er stC8ttet, se [rfid.js](/scriptotek/bibduck/blob/master/plugins/rfid.js#L133).
  Vi har bare testet dette med RFID-platene fra Bibliotheca.

* [general.js](plugins/general.js) aksepterer dokid i forfatter-feltet (det C8verste) pC% DOKST-skjermen, slik at man 
  slipper C% tabbe ned til dokid-feltet. Modulen definerer ogsC% kommandoen `!!` for C% tC8mme et felt
  man holder pC% C% skrive i.

* [stikksedler.js](plugins/stikksedler.js) definerer kommandoen `stikkseddel!` for C% skrive ut stikkseddel. Scriptet benytter seg 
  av den generelle konfigurasjonsfilen `stikksedler/config.json` og bibliotekspesifikk konfigurasjonsfil
  og maler (f.eks. `stikksedler/ureal.js` og `stikksedler/ureal/`) basert pC% bibliotekskoden man 
  angir i BIBDUCK-innstillingene.

* [imott-iret-auto.js](plugins/imott-iret-auto.js) er en automatiseringsprosedyre for C% registrere utlC%n, skrive ut stikkseddel 
  og sende hentebeskjed for mottatte *lC%n*. Den gjC8r ikke noe med mottatte *kopier*.
  Modulen er under uttesting.

* [loggers.js](plugins/loggers.js) logger utlC%n, retur og dokst-besC8k i loggvinduet, slik at man kan finne tilbake til 
  informasjon man "mister" fra BIBSYS-vinduet etter at en bruker har forlatt skranken. 
  Informasjonen lagres ikke til disk, og forsvinner nC%r man logger BIBDUCK eller trykker pC% knappen "TC8m logg".
  Modulen definerer kommandoene `dokid!` og `ltid!` for C% skrive ut hhv. siste dokid og ltid.

* [lstatus.js](plugins/lstatus.js) er en eksperimentell modul for C% skrive ut en liste over alle brukerens lC%n. Dette gjC8res med
  kommandoen `print!` pC% LTSTatus-skjermen. Modulen benytter Excel-arket `ltstatus.xls`.

* [dualbib.js](plugins/dualbib.js) er en svC&rt eksperimentell modul for C% arbeide med to vinduer samtidig. Modulen definerer
  kommandoen `n!` for C% hoppe mellom vinduer og `skr,*x*!` for C% vise resultat *x* fra et BIB-sC8k i 
  et annet vindu.

* [roald.js](plugins/roald.js) definerer kommandoen `roald!` for C% C%pne emneordsprogrammet [Roald](http://folk.uio.no/knuthe/progdist/). 
  Modulen er et godt utgangspunkt for C% lage egne moduler for C% kalle eksterne kommandoer.

* [efo.js](plugins/efo.js) definerer kommandoen `efo!` for C% fornye lC%n med purrestatus E i C)n operasjon.

Hvordan bruke BIBDUCK-grensesnittet?
-------------

Lukk eventuelle C%pne BIBSYS-vinduer fC8r du starter. BIBDUCK kan kun kommunisere med BIBSYS-instanser startet fra BIBDUCK.
Start deretter BIBDUCK fra ikonet pC% skrivebordet:

![BIBDUCK icon](http://localhostr.com/file/CjlJkWeoyZCa/desktop-icon.jpg)

Man fC%r da opp BIBDUCK-grensesnittet:

![BIBDUCK grensesnitt](screenshot.png)

Trykk pC% knappen **Nytt vindu** for C% starte en ny BIBSYS-instans, der du logger inn som vanlig.
Som et eksempel pC% makrofunksjonalitet, leder BIBDUCK deg imidlertid automatisk gjennom 
innledningsskjermene frem til skjermen BIBSYS SC8king. BIBDUCK sC8rger dessuten for at NumLock-tilstanden bevares 
gjennom innloggingsprosessen (SNetTerm skrur vanligvis av NumLock).

Trykker du pC% **Nytt vindu** igjen, startes en ny instans "BIBSYS 2", osv... 
Vinduet som har fokus indikeres med blC% bakgrunnsfarge i BIBDUCK. Et vindu fC%r fokus nC%r du skriver i det, eller
trykker i det blC% omrC%det. Normalt trenger man ikke C% tenke pC% fokus.

PrC8v C% gC% til REG-skjermen, og legg merke til at RFID-kontrollerprogrammet endrer modus:

![RFID-kontroller](http://localhostr.com/file/YKCVBoZu9TZn/rfid.png)

Legg merke til at brukernavn vises i vindustittelen. RFID-status vises ogsC% etter fC8rste modus-endring.
Man kan da sjekke RFID-status selv om RFID-kontrollerprogrammet er minimert. 
SC% lenge man ikke er ekstremt presset pC% skjermplass anbefales det imidlertid C% ha RFID-kontrollerprogrammet 
oppe, sC% man kan holde et lite C8ye med at modus endres som den skal.
En sjelden gang hender det f.eks. at RFID-kontrollerprogrammet mister kontakten med RFID-plata, 
og det vises da en feilmelding med hvit skrift pC% rC8d bakgrunn i vinduet. 
BIBDUCK klarer imidlertid *ikke* C% fange opp denne informasjonen og vil derfor fortsette som om alt var normalt.


Merknader og tips
-------------
* RFID-kontrollerprogramvaren fra Bibliotheca hC%ndterer ikke modus-endringer mens det 
ligger en (eller flere) bC8ker pC% platen. Det gC%r derfor ikke an C% f.eks. deaktivere og deretter aktivere
alarmen mens en bok ligger pC% platen. 
* Man kan laste inn tilleggsfunksjoner pC% nytt ved C% trykke ctrl-r
* Man kan endre loggnivC% med ctrl-0 (debug), ctrl-1 (info), ctrl-2 (warn) og ctrl-3 (error)
* Erfaringer og kjente feil: [For IT ansvarlige ved UBO og andre nysgjerrige](//github.com/scriptotek/bibduck/wiki/For-IT-ansvarlige-og-andre-nysgjerrige)
