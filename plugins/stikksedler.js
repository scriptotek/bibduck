/*****************************************************************************
 * <stikksedler.js>
 * Modul for å skrive ut stikksedler ved hjelp av Excel-maler
 * Av: Bård S. Tuseth (c) 2009
 *     Fredrik Hovind Juell (c) 2010
 *     Dan Michael O. Heggø (c) 2013
 *
 * Nye kommandoer:
 *   stikk!     : Skriver stikkseddel
 *
 * Avhenger av:
 *   hentebeskjed.js
 *****************************************************************************/




/*****************************************************************************
 * Klasse for stikkseddel
 *****************************************************************************/

var Stikkseddel = function(libnr, beststed, template_dir) {

	// $.bibduck.log('Template dir: ' + template_dir);

	// Settes under Innstillinger i brukergrensesnittet
	this.beststed = beststed;
	this.libnr = libnr;
	this.template_dir = template_dir;
	
	this.dokument = {};
	this.laaner = {};
	this.bibliotek = {};
	
	var that = this;
	
	this.print = function(type) {

		var fil = '';

		switch (type) {
			
			case 'reg':
				fil = this.laaner.spraak === 'ENG' ? 'reg_en.xls' : 'reg_no.xls';
				break;
			
			case 'ret':
				fil = this.bibliotek.navn === 'xxx' ? 'ret_en.xls' : 'ret_no.xls';
				break;
			
			case 'avh':
				fil = 'avh.xls';
				break;
				
			case 'avh_copy':
				fil = 'avh_copy.xls';
				break;

			case 'res':
				fil = 'res.xls';
				break;
		}

		if (fil === '') {
			$.bibduck.log('Ukjent stikkseddeltype', 'error');
			return;
		}

		ferdiggjor(fil);
	
	};

	// Ferdigstill, skriv ut og rydd opp
	var ferdiggjor = function(fil) {

		//$.bibduck.log('tpl: ' + that.template_dir);
		//$.bibduck.log('tpl: ' + that.template_dir);

		var libnr = $.bibduck.config.libnr,
			path = that.template_dir + fil,
			fso = new ActiveXObject("Scripting.FileSystemObject");
		
		// Sjekk at Excel-malfilen finnes:
		if (!fso.FileExists(path)) {
			$.bibduck.log('Stikkseddelfilen "' + path + '" finnes ikke!', 'error');
			return;
		}
		
		// Last inn malfilen:
		var excel = load_xls_template(path);
		
		// Utvid malsyntaks
		process_template_replacements(excel);

		// Skriv ut...
		excel.ActiveWorkbook.PrintOut();

		// ... og rydd opp
		excel.ActiveWorkbook.Close(0);
		excel.Quit();
		delete excel;
		excel = undefined;
	};

	// Formaterer datoer på norsk og engelsk
	var format_date = function(dt, lang) {
		if (dt === undefined) return '';
		var fdato = dt.split('-');
		if (lang === 'ENG') {
			return fdato[2] + '. ' + month_names_en[fdato[1]-1] + ' ' + fdato[0];
		} else {
			return fdato[2] + '. ' + month_names[fdato[1]-1] + ' ' + fdato[0];
		}
	};
	
	// Laster inn en Excel-malfil
	var load_xls_template = function (filename) {
		var printerStr = $.bibduck.config.printerName + ' on ' + $.bibduck.config.printerPort;
		var excel = new ActiveXObject('Excel.Application');
		excel.Visible = false;
		$.bibduck.log(getCurrentDir() + filename);
		excel.Workbooks.Open(getCurrentDir() + filename);
		if ($.bibduck.config.printerPort === '') {
			$.bibduck.log('Ingen stikkseddelskriver satt. Bruker standardskriver');
		} else if ($.bibduck.config.printerPort === 'none') {
			// bruk standardskriver
		} else {
			try {
				excel.Application.ActivePrinter = printerStr;
			} catch (e) {
				$.bibduck.log('Klarte ikke sette skriver. Bruker standardskriver', 'warn');
			}
		}
		return excel;
	};

	// Utvider malsyntaks i Excel-malfilen
	var process_template_replacements = function (excel) {
		var cells = new Enumerator(excel.ActiveSheet.UsedRange.Cells),
			cell,
			libv = '',
			libh = '',
			navn = that.laaner.etternavn + ', ' + that.laaner.fornavn,
			libnavn = '';

		if (that.dokument.utlstatus !== 'AVH') {
			if (that.laaner.kind === 'bibliotek') {

				// Låner er et bibliotek: Fjernlån
				libv = that.laaner.ltid.substr(3,3);
				libh = that.laaner.ltid.substr(6);
				navn = 'Fjernlån';  // til ' + that.laaner.navn;
				libnavn = that.laaner.navn;

			} else if (that.laaner.beststed !== that.beststed && !that.bibliotek.gangavstand) {
			
				// Sendes
				libv = that.bibliotek.ltid.substr(3,3);    // Venstre del av lib-nr.
				libh = that.bibliotek.ltid.substr(6);      // Høyre del av lib-nr.
				libnavn = that.bibliotek.navn;

			} else {
			
				// Sendes ikke

			}
		}
		if (that.dokument.utlaansdato === undefined) that.dokument.utlaansdato = iso_date();
		if (that.dokument.forfallsdato === undefined) that.dokument.forfallsdato = iso_date();
		if (that.dokument.forfvres === undefined) that.dokument.forfvres = iso_date();
		if (that.laaner.spraak === undefined) that.laaner.spraak = '';
		
		var infoEgenfornyningLinje1 = '',
			infoEgenfornyningLinje2 = '';

		if (that.laaner.spraak === 'ENG') {
			// Hvis ikke fjernlån, skriv ut litt ekstra info om fornying:
			if (that.laaner.kind === 'person') {
				if (that.dokument.purretype === 'E') {
					if (that.dokument.utlstatus === 'UTL/RES') {
						infoEgenfornyningLinje1 = "Please note:";
						infoEgenfornyningLinje2 = "This document can not be renewed as it has been reserved by someone else.";
					} else {
						infoEgenfornyningLinje1 = "This document can not be renewed online at BIBSYS Ask.";
						infoEgenfornyningLinje2 = "Please visit the library if you want to renew it.";
					}
				} else {
					infoEgenfornyningLinje1 = "Unless requested by someone else, this dokument can be";
					infoEgenfornyningLinje2 = "renewed online at BIBSYS Ask.";
				}
			}

		} else {
			// Skal boka til et annet bibliotek innad i organisasjonen?
			// Hvis ikke fjernlån, skriv ut litt ekstra info om fornying:
			if (that.laaner.kind === 'person') {
				if (that.dokument.purretype === 'E') {
					if (that.dokument.utlstatus === 'UTL/RES') {
						infoEgenfornyningLinje1 = "NB:";
						infoEgenfornyningLinje2 = "Dette dokumentet kan ikke fornyes, da det er reservert for en annen låntaker.";
					} else {
						infoEgenfornyningLinje1 = "Dette lånet kan du ikke fornye selv på BIBSYS Ask.";
						infoEgenfornyningLinje2 = "Kom innom biblioteket hvis du ønsker å fornye dette lånet.";
					}
				} else {
					infoEgenfornyningLinje1 = "Dette lånet kan du fornye selv på BIBSYS Ask";
					infoEgenfornyningLinje2 = "hvis det ikke kommer reserveringer.";
				}
			}
		}
		
		var forfallVedRes = format_date(that.dokument.forfvres, that.laaner.spraak);
		if (forfallVedRes) {
			forfallVedResDato = forfallVedRes;
			if (that.dokument.forfvres !== that.dokument.forfallsdato) {
				forfallVedRes = 'Ved reservasjoner kan dokumentet bli innkalt fra ' + forfallVedRes;
			}
		}
		
		var adresse = that.bibliotek.adresse;

		for (; !cells.atEnd(); cells.moveNext()) {
			cell = cells.item();
			if (cell.Value !== undefined && cell.Value !== null) {
				cell.Value = cell.Value
						.replace('{{Navn}}', navn)
						.replace('{{Ltid}}', that.laaner.ltid ? that.laaner.ltid : '-')
						.replace('{{Adresse}}', adresse)
						.replace('{{Tittel}}', that.dokument.tittel ? that.dokument.tittel : '-')
						.replace('{{Dokid}}', that.dokument.dokid ? that.dokument.dokid : '-')
						.replace('{{Bestnr}}', that.dokument.bestnr ? that.dokument.bestnr : '-')
						.replace('{{Utlånsdato}}', format_date(that.dokument.utlaansdato, that.laaner.spraak))
						.replace('{{Forfallsdato}}', format_date(that.dokument.forfallsdato, that.laaner.spraak))
						.replace('{{ForfallVedRes}}', forfallVedRes)
						.replace('{{ForfallVedResDato}}', forfallVedResDato)
						.replace('{{DagensDato}}', format_date(iso_date(), that.laaner.spraak))
						//.replace('{{Dato}}', that.format_date(iso_date()))
						.replace('{{Libnavn}}', libnavn)
						.replace('{{LIBV}}', libv)
						.replace('{{LIBH}}', libh)
						.replace('{{Hentenr}}', that.dokument.hentenr)
						.replace('{{Avsender}}', that.beststed)
						.replace('{{InfoEgenfornyingLinje1}}', infoEgenfornyningLinje1)
						.replace('{{InfoEgenfornyingLinje2}}', infoEgenfornyningLinje2);
			}
		}
	};

};

/*****************************************************************************
 * Resten
 *****************************************************************************/

(function() {
	var worker,
		client,
		dok = {},
		laaner = {},
		lib = {},
		hjemmebibliotek = '',
		config,
		seddel,
		callback,
		//siste_bestilling = { active: false },
		fso = new ActiveXObject("Scripting.FileSystemObject"),
			shell = new ActiveXObject("WScript.Shell"),
			appdata = shell.ExpandEnvironmentStrings("%ALLUSERSPROFILE%"),
			stikk_path = appdata + '\\Scriptotek\\Bibduck\\stikk.txt';


	function les_dokstat_skjerm() {

		if (client.get(2, 1, 28) !== 'Utlånsstatus for et dokument') {
			client.alert('Vi er ikke på DOKST-skjermen :(');
			client.setBusy(false);
			return;
		}

		// Sjekker hvilken linje tittelen står på:
		if (client.get(7, 2, 7) == 'Tittel') {
			// Lån fra egen samling
			seddel.dokument.tittel = client.get(7, 14, 80).trim();
		} else if (client.get(8, 2, 7) == 'Tittel') {
			// ik...
			seddel.dokument.tittel = client.get(8, 13, 80).trim();
		} else {
			// Relativt sjelden case? Linje 7-10 er fritekst, og 
			// tittel og forfatter bytter typisk mellom linje 7 og 8.
			// En enkel test, som sikkert vil feile i flere tilfeller:
			var tittel1 = client.get(7, 2, 80).trim(),
				tittel2 = client.get(8, 2, 80).trim();
			if (tittel1.length > tittel2.length) {
				seddel.dokument.tittel = tittel1;
			} else {
				seddel.dokument.tittel = tittel2;
			}
		}

		seddel.dokument.dokid = client.get( 6, 31, 39);
		
		/*
		
		Må teste mer hva som skiller et heftid fra annen informasjon som kan dukke opp her.
		
		var heftid = client.get(5,2,10);  // trimmes automatisk
		if (heftid.length === 9) {
			seddel.dokument.heftid = heftid;
			$.bibduck.log('Det kan se ut som vi har et HEFTID');
			seddel.dokument.dokid = seddel.dokument.heftid;
		}
		*/

		if (seddel.dokument.dokid === '') {
			client.alert('Har du husket å trykke enter?');
			client.setBusy(false);
			return;
		}

		seddel.dokument.utlstatus = client.get( 3, 46, 65);   // AVH, RES, UTL, UTL/RES, ...


		if (client.get(10, 2, 7) == 'Tittel') {
			seddel.dokument.tittel = client.get(10, 14, 79);
		} else if (client.get(11, 2, 7) == 'Tittel') {
			seddel.dokument.tittel = client.get(11, 14, 79);
		} else if (client.get(12, 2, 7) == 'Tittel') {
			seddel.dokument.tittel = client.get(12, 14, 79);
		} else if (client.get(13, 2, 7) == 'Tittel') {
			seddel.dokument.tittel = client.get(13, 14, 79);
		}

		// Dokument til avhenting?
		if (seddel.dokument.utlstatus === 'AVH') {

			seddel.dokument.hentenr = client.get(1, 44, 50);
			seddel.dokument.hentefrist = client.get(1, 26, 35);
			
			if (seddel.dokument.hentenr == '') {
				client.alert('Hentenummer mangler på DOKST-skjermen.');
				$.bibduck.writeErrorLog(client, 'dokst_hentenr_mangler2');
				$.bibduck.log('Hentenummer mangler på DOKST-skjermen.','error');
				return;
			}

		} else {

			seddel.laaner.ltid      = client.get(14, 11, 20);
			seddel.dokument.utlaansdato  = client.get(18, 18, 27);   // Utlånsdato
			seddel.dokument.forfvres     = client.get(20, 18, 27);   // Forfall v./res
			seddel.dokument.forfallsdato = client.get(21, 18, 27);   // Forfallsdato
			seddel.dokument.purretype    = client.get(17, 68, 68);
			seddel.dokument.kommentar    = client.get(23, 17, 80).trim();

			//Tester om låntaker er et bibliotek:
			if (seddel.laaner.ltid.substr(0,3) == 'lib') {
				seddel.laaner.kind = 'bibliotek';
				seddel.laaner.navn = client.get(14, 22, 79).trim();
			} else {
				seddel.laaner.kind = 'person';
			}

		}

		// DEBUG:
		/*
		$.bibduck.log('Info om lånet:');
		$.each(dok, function(k,v) {
			$.bibduck.log('  ' + k + ': ' + v);
		});
*/

		worker.resetPointer();

		// Hva gjør vi ift. UTL/RES?
		// Skriver ut stikkseddel for det utlånet eks. eller det reserverte?

		if (seddel.dokument.utlstatus === 'AVH') {

			// Vi trenger ikke mer informasjon. 
			// La oss kjøre i gang Excel-helvetet, joho!!
			seddel.print('avh');
			emitComplete();

		} else if (seddel.dokument.utlstatus === 'RES') {
		
			client.alert('For å sende hentebeskjed, gå til RLIST og bruk RES-O-MAT-knappen');
			return;

		} else if (seddel.laaner.kind === 'person') {

			// Vi trenger mer info om låneren:
			worker.send('ltsø,' + seddel.laaner.ltid + '\n');
			worker.wait_for('Fyll ut:', [5,1], function() {
				// Vi sender enter på nytt
				worker.send('\n');
				worker.wait_for('Sist aktiv dato', [22,1], les_ltsok_skjerm);
			});

		} else {

			// Vi trenger ikke mer informasjon. 
			// La oss kjøre i gang Excel-helvetet, joho!!
			seddel.print('reg');
			emitComplete();

		}

	}

	function emitComplete() {
		$.bibduck.log("Stikkseddel ferdig");
		if (callback !== undefined) {
			setTimeout(function() { // a slight delay never hurts
				var data = {
					patron: seddel.laaner,
					library: seddel.bibliotek,
					document: seddel.dokument,
					beststed: seddel.beststed
				};
				callback(data);
			}, 200);
		}
	}

	function res_sendes() {

		if (worker.get(2, 1, 24) !== 'Opplysninger om låntaker') {
			client.alert("Vi er ikke på LTSØ-skjermen :(");
			client.setBusy(false);
			return;
		}

		// Vi avslutter med å gå til RLIST igjen for å skrive kommentar
		client.send('\tRLIST,' + seddel.dokument.dokid + '\n');
		client.wait_for([
			['Reserveringsliste', [4,8], function() {
				client.send(seddel.dokument.dokid + '\n');
				client.wait_for('Hentefrist', [20,5], res_sendes_2);
			}],
			['Hentefrist', [20,5], res_sendes_2]
		]);

	}

	function skrivRlistMelding(ltid, melding) {

		var checkPage = function(pageNo) {
			var fndEntry = false,
				tabs = '\t';
			for (var line = 3; line <= 17; line+=7) {
				//$.bibduck.log(client.get(line, 15, 24));
				if (client.get(line, 15, 24) == ltid) {
					client.send(tabs + melding);
					fndEntry = true;
					break;
				}
				tabs += '\t\t';
			}
			if (fndEntry) {

				// Vi fant dokid på denne siden

			} else {
				// Vi fant ikke dokid på denne siden. Vi sjekker neste side hvis den finnes
				if (client.get(25,49,51) !== '') {
					$.bibduck.sendSpecialKey('F8');
					var firstEntryOnNextPage = (pageNo*3 + 1);
					firstEntryOnNextPage = (firstEntryOnNextPage < 10) ? ' ' + firstEntryOnNextPage : '' + firstEntryOnNextPage;
					client.wait_for(firstEntryOnNextPage, [3, 6], function() { checkPage(pageNo+1); });
				} else {
					$.bibduck.log('Fant ikke ltid ' + seddel.laaner.ltid + ' i reservasjonslisten!', 'error');
				}
			}

		};

		checkPage(1);

	}
	
	function res_sendes_2() {

		// Hvilket bibliotek skal dokumentet sendes til? Finn signatur for LIBNR
		for (var s in config.sigs) {
			if (config.sigs[s] === seddel.bibliotek.ltid) {
				sig = s;
			}
		}

		client.alert('Låner har bestillingssted ' + seddel.laaner.beststed +
					', så dokumentet må sendes. Du skal få en stikkseddel.');

		skrivRlistMelding(seddel.laaner.ltid, iso_date() + ': ' + seddel.dokument.dokid + ' sendt ' + sig + '          ');

		emitComplete();
		seddel.print('res');

		if (config.har_rfid['lib' + hjemmebibliotek] && !config.har_rfid[seddel.bibliotek.ltid]) {
			client.alert('NB! Mottakerbiblioteket bruker ikke RFID. \n' +
						'Du må derfor avalarmisere dokumentet før sending.');
		}

	}

	function send_hentebeskjed() {

		(new Hentebeskjed(client, seddel.laaner.ltid, seddel.dokument.dokid, function(response) {

			seddel.dokument.hentenr = response.hentenr;
			seddel.dokument.hentefrist = response.hentefrist;
			
			// Sjekker hvilken linje tittelen står på:
			if (client.get(7, 2, 7) == 'Tittel') {
				// Lån fra egen samling
				seddel.dokument.tittel = client.get(7, 14, 80).trim();
			} else if (client.get(8, 2, 7) == 'Tittel') {
				// ik...
				seddel.dokument.tittel = client.get(8, 13, 80).trim();
			} else {
				// Relativt sjelden case? Linje 7-10 er fritekst, og 
				// tittel og forfatter bytter typisk mellom linje 7 og 8.
				// En enkel test, som sikkert vil feile i flere tilfeller:
				var tittel1 = client.get(7, 2, 80).trim(),
					tittel2 = client.get(8, 2, 80).trim();
				if (tittel1.length > tittel2.length) {
					seddel.dokument.tittel = tittel1;
				} else {
					seddel.dokument.tittel = tittel2;
				}
			}

			if (client.getCursorPos().row == 3) {
				client.send('RLIST,' + seddel.dokument.dokid + '\n');  // for å oppfriske skjermen slik at status endres fra RES til AVH
			} else {
				client.send('\tRLIST,' + seddel.dokument.dokid + '\n'); // for å oppfriske skjermen slik at status endres fra RES til AVH
			}
			client.wait_for('Hentefrist', [20,5], avslutt_avh_prosedyre_paa_rlist); // Vi har alltid DOKID, ikke KNYTTID

		})).send();
	}

	function avslutt_avh_prosedyre_paa_rlist() {

		// Hvilket bibliotek skal dokumentet sendes til? Finn signatur for LIBNR
		for (var s in config.sigs) {
			if (config.sigs[s] === seddel.libnr) {
				sig = s;
			}
		}

		skrivRlistMelding(seddel.laaner.ltid, iso_date() + ': Til avhenting ved ' + sig + '                  ' );

		// La oss kjøre i gang Excel-helvetet, joho!!
		if (seddel.dokument.artikkelkopi) {
			$.bibduck.log('[STIKK] Hentebeskjed sendt til ' + seddel.laaner.ltid +'. Hentenr.: ' + seddel.dokument.hentenr + ' (artikkelkopi)', 'info');
			seddel.print('avh_copy');
		} else {
			$.bibduck.log('[STIKK] Hentebeskjed sendt til ' + seddel.laaner.ltid +'. Hentenr.: ' + seddel.dokument.hentenr, 'info');
			seddel.print('avh');
		}
		emitComplete();

	}

	function les_ltsok_skjerm() {
		var that = this;
		if (worker.get(2, 1, 24) !== 'Opplysninger om låntaker') {
			client.alert("[STIKK] Vi er ikke på LTSØ-skjermen :(");
			client.setBusy(false);
			return;
		}
		seddel.laaner.beststed  = worker.get( 7, 71, 80).trim();
		seddel.laaner.etternavn = worker.get( 5, 18, 58).trim();
		seddel.laaner.fornavn   = worker.get( 6, 18, 58).trim();
		seddel.laaner.spraak    = worker.get(19, 41, 44).trim();
		seddel.laaner.kategori  = worker.get(18, 41, 41).trim();
		seddel.laaner.innkallingsadresse  = worker.get(11, 18, 80).trim();

		// DEBUG:
		/*
		$.bibduck.log('Info om låner:');
		$.each(seddel.laaner, function(k,v) {
			$.bibduck.log('  ' + k + ': ' + v);
		});*/

		seddel.bibliotek.ltid = '';
		seddel.bibliotek.navn = '';
		if (seddel.laaner.beststed in config.bestillingssteder) {
			seddel.bibliotek.ltid = config.bestillingssteder[seddel.laaner.beststed];
			//$.bibduck.log(seddel.bibliotek.ltid);
		} else {
			// En bruker med lånekort fra f.eks. tek (NTNU) 
			// som vi kobler, vil beholde beststed tek.
			$.bibduck.log("[STIKK] Kjenner ikke libnr for bestillingssted: " + seddel.laaner.beststed, 'warn');
			// return;
		}
		if (seddel.bibliotek.ltid in config.biblnavn) {
			seddel.bibliotek.navn = config.biblnavn[seddel.bibliotek.ltid];
		} else if (seddel.bibliotek.ltid !== '') {
			$.bibduck.log("[STIKK] Kjenner ikke navn for libnr: " + seddel.bibliotek.ltid, 'warn');
		}
		
		seddel.bibliotek.adresse = '';
		if (seddel.bibliotek.ltid in config.adresser) {
			seddel.bibliotek.adresse = config.adresser[seddel.bibliotek.ltid];
		}
		
		seddel.bibliotek.gangavstand = false;
		
		$.bibduck.log('[STIKK] Låner har beststed: ' + seddel.laaner.beststed + '. Vi er: ' + seddel.beststed);

		// Sjekk om låners bibliotek er innen gangavstand
		if (seddel.laaner.beststed !== seddel.beststed && config.gangavstand[seddel.libnr]) {
			for (var key in config.gangavstand[seddel.libnr]) {
				if (config.gangavstand[seddel.libnr][key] == seddel.bibliotek.ltid) {
					seddel.bibliotek.gangavstand = true;
					$.bibduck.log('[STIKK] Låner har bestillingssted ' + seddel.bibliotek.ltid + ', som er innen gangavstand fra ' + seddel.libnr + ', så vi sender ikke boka.', 'info');
				}
			}
		}

		// Sjekk om låner er importert (LTKOP)
		// En bruker med lånekort fra f.eks. tek (NTNU)
		// som vi importerer, vil beholde beststed tek.
		seddel.laaner.importert = undefined;
		for (var key in config.bestillingssteder) {
			if (key === seddel.laaner.beststed) {
				seddel.laaner.importert = false;
			}
		}
		if (seddel.laaner.importert === undefined) {
			seddel.laaner.importert = true;
			$.bibduck.log('[STIKK] Låner har bestillingssted ' + seddel.laaner.beststed + '. ' +
				'Vi antar at låner er koplet og behandler hen som en lokal bruker.', 'info');
		}

		// DEBUG:
		/*
		$.bibduck.log('Info om bibliotek:');
		$.each(lib, function(k,v) {
			$.bibduck.log('  ' + k + ': ' + v);
		});*/

		// Nå har vi informasjonen vi trenger. La oss kjøre i gang Excel-helvetet, joho!!

		// @TODO: Hva med UTL/RES ?
		if (seddel.dokument.utlstatus === 'RES') {

			res_sendes();

		} else if (seddel.dokument.utlstatus === 'AVH') {

			if (seddel.laaner.beststed === seddel.beststed || seddel.laaner.importert || seddel.bibliotek.gangavstand) {

				seddel.dokument.utlstatus = 'AVH';

				client.send('dokst,' + seddel.dokument.dokid + '\n');
				client.wait_for([

					// Vanlig dokstat-skjerm:
					['Utlkommentar', [23,1], function() {
						// dok.dokid fra rlist kan være et knyttid. Vi overskriver derfor med
						// det virkelige dokid-et.
						seddel.dokument.dokid = client.get(6,31,39);
						var status = client.get(3,46,52);
						if (status === 'UTL/RES' || status === 'UTL') {
							$.bibduck.log('Dokumentet med dokid ' + seddel.dokument.dokid + ' er allerede utlånt!', 'warn');
							client.alert('Dokumentet med dokid ' + seddel.dokument.dokid + ' er utlånt! Hvis du vil gi det til en ny person må du ta retur først.');
							client.setBusy(false);
						} else {
							send_hentebeskjed();
						}
					}],

					// Nesten blank dokstat-skjerm vises hvis det er et annet dokid enn det man har skannet som er reservert:
					['Et annet dokument har reserveringer', [19,2], function() {
						// dok.dokid fra rlist kan være et knyttid. Vi overskriver derfor med
						// det virkelige dokid-et.
						seddel.dokument.dokid = client.get(5,31,39);
						send_hentebeskjed();
					}]
				]);
					
			} else {

				seddel.dokument.utlstatus = 'RES';

				res_sendes();
			
			}

		} else {

			// Gi beskjed hvis boka skal ut av huset
			if (seddel.laaner.kind === 'person' && seddel.laaner.beststed !== seddel.beststed && !seddel.bibliotek.gangavstand && seddel.bibliotek.ltid !== '') {

				$.bibduck.log('Låner har et eksternt bestillingssted: ' + seddel.laaner.beststed + ' (' + seddel.bibliotek.ltid + ')', 'warn');
				client.alert('Obs! Låner har beststed ' + seddel.laaner.beststed);

				// Hvis boken skal sendes, så gå til utlånskommentarfeltet.
				client.send('en,' + seddel.dokument.dokid + '\n');
				client.wait_for('Utlmkomm:', [8,1], function() {
					client.send('\t\t\t');
					seddel.print('reg');
					emitComplete();
				});

			// Hvis ikke går vi tilbake til dokst-skjermen:
			} else {

				if (seddel.laaner.kind === 'person' && seddel.laaner.beststed !== seddel.beststed) {
					$.bibduck.log('NB! Låner har et eksternt bestillingssted: ' + seddel.laaner.beststed, 'warn');
				}

				//result = snt.MessageBox("Vil du gå til REG for å låne ut flere bøker?", "Error", ICON_QUESTION Or BUTTON_YESNO Or DEFBUTTON2)

				//if (result == IDYES) {
				//  // ... tilbake til utlånsskjerm for å registrere flere utlån.
				//  snt.Send("reg,"+ltid)
				//  snt.QuickButton("^M")
				//Else
					// ... tilbake til dokst, for å sende hentebeskjed
					client.send('dokst,' + seddel.dokument.dokid + '\n');
					client.wait_for('Utlkommentar', [23,1], function() {
						// FINITO, emit
						seddel.print('reg');
						emitComplete();
					});
				//}
			}
		}
	}

	function start_from_imo(options) {
		seddel.laaner.kind = 'person';
		seddel.dokument.utlstatus = 'AVH';

		if (options && options.bestnr) {
			seddel.dokument.bestnr = options.bestnr;
		}
		if (options && options.artikkelkopi) {
			seddel.dokument.artikkelkopi = true;
		}

		var firstline = client.get(1);
		var tilhvem = firstline.match(/på (sms|Email) til (.+) merket (.+)/);
		if (tilhvem === null) {
			$.bibduck.log("Ikke noe hentenummer på skjermen", "warn");
			client.alert("Hentebeskjed er sendt, men BIBSYS har visst sluttet å lage hentenummer til kopibestillingene våre (noen som vet hvorfor?). For stikkseddel; gå til LTSØK og trykk på \"Navn og dato\"-knappen.");
			//$.bibduck.writeErrorLog(client, 'hentenr_mangler');
			return;
		}
		var name = tilhvem[2];
		var nr = tilhvem[3].trim();
		if (nr === '') {
			$.bibduck.log('Fant ikke noe hentenr.', 'error');
			client.setBusy(false);
			return;
		}

		seddel.dokument.hentenr = nr;
		seddel.dokument.hentefrist = '-';

		// Vi trenger ikke mer informasjon.
		// La oss kjøre i gang Excel-helvetet, joho!!
		if (seddel.dokument.artikkelkopi) {
			$.bibduck.log('Hentenr.: ' + nr + ' (kopibestilling)', 'info');
			seddel.print('avh_copy');
		} else {
			$.bibduck.log('Hentenr.: ' + nr + ' (lånebestilling)', 'info');
			seddel.print('avh');
		}
		emitComplete();

	}

	function utlaan() {

		if (client.get(2, 1, 22) == 'Registrere utlån (REG)') {
			var dokid = client.get(10, 7, 15);
			if (dokid == '') {
				$.bibduck.log('Finner ikke DOKID. Er utlånet fullført?','warn');
				client.alert('Finner ikke DOKID. Er utlånet fullført?');
				client.setBusy(false);
				return;
			}
			// Gå til DOKST-skjerm:
			worker.resetPointer();
			worker.send('dokst\n');
			//Kan ikke ta dokst, (med komma) for da blir dokid automatisk valgt og aldri refid, sender separat
			worker.wait_for('Utlånsstatus for et dokument', [2,1], function() {
				worker.send(dokid + '\n');
				worker.wait_for('Utlkommentar', [23,1], function() {
					les_dokstat_skjerm(worker);
				});
			});
		
		} else if (client.get(2, 1, 28) == 'Utlånsstatus for et dokument') {
			les_dokstat_skjerm();
		}

	}
	
	function start_from_ltsok(info) {
		worker.resetPointer();
		
		//$.bibduck.log('start from ltsøk');

		if (info !== undefined && info.ltid !== undefined) {
			
			// Vi har mottatt en stikkseddelforespørsel fra en annen prosess
			seddel.dokument.dokid = info.dokid;
			seddel.laaner.ltid = info.ltid;
			seddel.dokument.utlstatus = 'AVH';
			
			les_ltsok_skjerm();
			
		} else {
		
			// Vi skriver ut en retur-seddel. Nyttig f.eks. hvis 
			// man ikke får stikkseddel fra IRET

			if (client.get(18, 18, 20) !== 'lib') {
				client.alert('Beklager, kan ikke skrive returseddel når låntakeren ikke er et bibliotek.');
				$.bibduck.log('Kan ikke skrive returseddel når låntakeren ikke er et bibliotek.', 'warn');
				client.setBusy(false);
				return;
			}

			seddel.laaner.ltid = client.get(18, 18, 27);
			seddel.laaner.navn = client.get(10, 18, 50);
			seddel.laaner.kind = 'bibliotek';
			seddel.bibliotek.ltid = seddel.laaner.ltid;
			seddel.bibliotek.navn = seddel.laaner.navn;

			seddel.print('ret');
			emitComplete();
		}
	}

	function retur() {
		worker.resetPointer();

		seddel.laaner = {};
		seddel.bibliotek = {};
		seddel.dokument = {};

		if (client.get(2).indexOf('IRETur') !== -1) {

			seddel.dokument.dokid = client.get(1, 1, 9);
			seddel.dokument.bestnr = client.get(4, 49, 57);

			seddel.laaner.ltid = client.get(6, 15, 24);
			seddel.laaner.navn = client.get(7, 20, 50);
			seddel.laaner.kind = 'bibliotek';
			seddel.bibliotek.ltid = client.get(6, 15, 24);
			seddel.bibliotek.navn = client.get(7, 20, 50);
			if (seddel.laaner.navn === 'xxx') {
				seddel.laaner.navn = '';
				seddel.laaner.navn = '';
				seddel.bibliotek.ltid = '';
				seddel.bibliotek.navn = '';
			}

		} else {

			// Retur til annet bibliotek innad i organisasjonen

			var sig = client.get(11, 14, 40).split(' ')[0];
			seddel.dokument.dokid = client.get(6, 31, 39);
			seddel.dokument.bestnr = '';

			if (sig in config.sigs) {
				seddel.bibliotek.ltid = config.sigs[sig];
				seddel.bibliotek.navn = config.biblnavn[seddel.bibliotek.ltid];
			} else {
				client.alert('Beklager, BIBDUCK kjenner ikke igjen signaturen "' + sig + '".');
				client.setBusy(false);
				return;
			}

		}

		seddel.dokument.tittel = '';
		if (client.get(7, 2, 7) == 'Tittel') {
			seddel.dokument.tittel = client.get(7, 14, 79);
		} else if (client.get(8, 2, 7) == 'Tittel') {
			seddel.dokument.tittel = client.get(8, 14, 79);
		} else if (client.get(9, 2, 7) == 'Tittel') {
			seddel.dokument.tittel = client.get(9, 14, 79);
		} else if (client.get(10, 2, 7) == 'Tittel') {
			seddel.dokument.tittel = client.get(10, 14, 79);
		}

		if (hjemmebibliotek === '') {
			client.alert('Libnr. er ikke satt. Dette setter du under Innstillinger.');
		}
		if (seddel.bibliotek.ltid === 'lib' + hjemmebibliotek) {
			client.alert('Boka hører til her. Returseddel trengs ikke.');
			client.setBusy(false);
			return;
		}

		seddel.print('ret');
		emitComplete();

	}

	function start(options) {

		try {
			$('audio#ping').get(0).play();
		} catch (e) {
			// IE8?
		}
		
		hjemmebibliotek = $.bibduck.config.libnr;
		var libnr = 'lib' + $.bibduck.config.libnr;
		var template_dir = 'plugins\\stikksedler\\' + config.malmapper[libnr] + '\\';
		var beststed = '';

		for (var key in config.bestillingssteder) {
			if (config.bestillingssteder[key] == libnr) {
				beststed = key;
			}
		}
		
		if (libnr === 'lib') {
			client.alert('Obs! Libnr. er ikke satt enda. Dette setter du under Innstillinger i Bibduck.');
			client.setBusy(false);
			return;
		} else if (beststed === '') {
			client.alert('Fant ikke et bestillingssted for biblioteksnummeret ' + libnr + ' i config.json!');
			client.setBusy(false);
			return;
		}
		
		//$.bibduck.log('template dir: ' + template_dir);

		seddel = new Stikkseddel(libnr, beststed, template_dir);

		if (client.get(2, 1, 22) === 'Registrere utlån (REG)') {
			$.bibduck.log('Lager stikkseddel fra REG-skjermen');
			utlaan();
		} else if (client.get(14, 1, 8) === 'Låntaker') { // DOkstat
			$.bibduck.log('Lager stikkseddel fra DOKST-skjermen');
			utlaan();
		} else if (client.get(15, 2, 13) === 'Returnert av') {
			$.bibduck.log('Lager stikkseddel fra RET-skjermen');
			retur();
		} else if (client.get(1).indexOf('er returnert') !== -1 && client.get(2).indexOf('IRETur') !== -1) { // Retur innlån (IRETur)
			$.bibduck.log('Lager stikkseddel fra IRET-skjermen');
			retur();
		} else if (client.get(2, 1, 32) === 'Opplysninger om låntaker (LTSØk)') {
			$.bibduck.log('Lager stikkseddel fra LTSØK-skjermen');
			start_from_ltsok(options);
		} else if (client.get(2, 1, 15) === 'Reservere (RES)') {
			//$.bibduck.log('Lager stikkseddel fra RES-skjermen');
			$.bibduck.log('Feil stikkseddelknapp');
			client.setBusy(false);
			client.alert('Prøv den andre knappen istedet (RES-O-MAT)');
		} else if (client.get(2, 1, 25) === 'Reserveringsliste (RLIST)') {
			//$.bibduck.log('Lager stikkseddel fra RLIST-skjermen');
			//start_from_rlist();
			$.bibduck.log('Feil stikkseddelknapp');
			client.alert('For å sende hentebeskjed, bruk RES-O-MAT-knappen. For å skrive ut stikkseddel for en bok som allerede har status AVH, gå til DOKST og prøv igjen.');
		/*} else if (client.get(2, 1, 12) === 'Motta innlån') {
			$.bibduck.alert('Lager stikkseddel fra IMO-skjermen');
			//start_from_imo(options);
		*/
		} else {
			client.setBusy(false);
			$.bibduck.log('Stikkseddel fra denne skjermen er ikke støttet', 'warn');
			client.alert('Stikkseddel fra denne skjermen er ikke støttet (enda). Ta DOKST og prøv igjen');
		}
	}

	$.bibduck.plugins.push({

		name: 'Stikkseddel-tillegg',

		initialize: function() {
			var that = this;

			$('#settings-form table').append('<tr>' +
			  '<th>' +
			   ' Automatisk stikkseddel' +
				'</th><td>' +
				'<input type="radio" name="autostikk_reg" id="autostikk_reg_alle" ' + ($.bibduck.config.autoStikkEtterReg == 'autostikk_reg_alle' ? ' checked="checked"' : '') + ' />' +
			   '   <label for="autostikk_reg_alle">etter alle utlån</label>' +
				'<input type="radio" name="autostikk_reg" id="autostikk_reg_lib" ' + ($.bibduck.config.autoStikkEtterReg == 'autostikk_reg_lib' ? ' checked="checked"' : '') + ' />' +
			   '   <label for="autostikk_reg_lib">etter utlån til bibliotek</label>' +
				'<input type="radio" name="autostikk_reg" id="autostikk_reg_ingen" ' + ($.bibduck.config.autoStikkEtterReg == 'autostikk_reg_ingen' ? ' checked="checked"' : '') + ' />' +
				'  <label for="autostikk_reg_ingen">aldri</label>' +
				' </td>' +
				'</tr>');

			that.listenForNotificationFile();
		},

		/**
		 * Kalles av Bibduck når innstillingene skal lagres
		 */
		saveSettings: function(file) {

			var autostikk_reg = $('input[name="autostikk_reg"]:checked').attr('id');
			$.bibduck.config.autoStikkEtterReg = autostikk_reg;
			file.WriteLine('autoStikkEtterReg=' + $.bibduck.config.autoStikkEtterReg);

		},

		/**
		 * Kalles av Bibduck når innstillingene skal lastes
		 */
		loadSettings: function(data) {

			// Default
			$.bibduck.config.autoStikkEtterReg = 'autostikk_reg_ingen';

			var line;
			for (var i = 0; i < data.length; i += 1) {
				line = data[i];
				if (line[0] === 'autoStikkEtterReg') {
					if (line[1] !== 'undefined') {
						$.bibduck.config.autoStikkEtterReg = line[1];
					}
				}
			}

		},

		/**
		 * Primitivt meldingssystem: Vi sjekker om en bestemt fil finnes. Hvis
		 * den finnes skriver vi ut stikkseddel og sletter filen. Slik kan andre
		 * prosesser be om stikksedler.
		 */
		listenForNotificationFile: function() {
			var that = this,
				check = function() {
				if (fso.FileExists(stikk_path)){
				
					var bibsys = $.bibduck.getFocused(),
						txt = readFile(stikk_path);
						
					if (bibsys && bibsys.connected) {

						$.bibduck.log('[STIKK] Fant stikk-fil');

						//$.bibduck.log(txt);
						var request;

						try {
							request = $.parseJSON(txt);
						} catch (e) {
							
						}

						fso.DeleteFile(stikk_path);

						bibsys = $.bibduck.getFocused();
						
						if (bibsys.busy) {
							$.bibduck.log('[STIKK] Bibsys-vinduet er opptatt.', 'error');
							bibsys.alert("Bibsys-vinduet er opptatt. Om problemet vedvarer kan du omstarte BIBDUCK.");

						} else if (request) {

							if (request.ltid) {
								$.bibduck.log('[STIKK] >>> Forespørsel fra RES-O-MAT. Ltid: ' + request.ltid + ', dokid: ' + request.dokid + ' <<<', 'info');
							} else {
								$.bibduck.log('[STIKK] >>> Forespørsel om DUCKSEDDEL <<<', 'info');
							}

							bibsys.setBusy(true);
							callback = function() {
								bibsys.setBusy(false);
								$.bibduck.log('[STIKK] Ferdig');
							};
							that.forbered_stikkseddel(bibsys, function() {
								//$.bibduck.log('forbered_stikkseddel callback');
								bibsys.unidle();
								bibsys.update();   // force update
								start(request);
							});
						}
					}
				}
				that.timer = setTimeout(check, 500);
			};
			if (that.timer === undefined) {
				that.timer = setTimeout(check, 500);
			}
		},

		lag_stikkseddel: function(bibsys, cb, options) {

			bibsys.off('waitFailed');
			bibsys.on('waitFailed', function() {
				$.bibduck.log('[STIKK] Stikkseddelutskriften ble avbrutt', 'error');
				bibsys.setBusy(false);
			});
			callback = cb;
			
			this.forbered_stikkseddel(bibsys, function() { start(options); });
		},

		forbered_stikkseddel: function(bibsys, startfn) {

			client = bibsys;
			
			//$.bibduck.log('forbered_stikkseddel');
		
			if ($.bibduck.config.printerPort === '') {
				//client.alert('Sett opp stikkseddelskriver ved å trykke på knappen «Innstillinger» først.');
				//setWorking(false);
				//return;
			}

			if ($.bibduck.getBackgroundInstance() !== null) {
				worker = $.bibduck.getBackgroundInstance();
			} else {
				worker = client;
			}

			// Load config if not yet loaded
			if (config === undefined) {
				$.bibduck.log(' -> Load: plugins/stikksedler/config.json');
				$.getJSON('plugins/stikksedler/config.json', function(json) {
					config = json;
					startfn();
				});
			} else {
				startfn();
			}

		},

		waiting: false,

		/*
		 * Denne metoden kalles av Bibduck med rundt 100 millisekunds mellomrom.
		 */
		update: function(bibsys) {

			// Vi må huske om siste mottatte bestilling var en kopi (K) eller et lån (L)
			// for å skrive ut rett stikkseddel (lån må lånes ut på automat, kopier ikke)
			/*if (bibsys.get(1, 11, 20) === 'er mottatt') {
				if (!siste_bestilling.active) {
					siste_bestilling = {
						bestnr: bibsys.get(1, 1, 9),
						laankopi: bibsys.get(8, 37, 37)
					};
				}
				siste_bestilling.active = true;
			} else if (siste_bestilling.active) {
				siste_bestilling.active = false;
			}*/

			//trigger2 = (bibsys.get(1).indexOf('er returnert') !== -1 && bibsys.get(2).indexOf('IRETur') !== -1),

			// Stikkseddel når man skriver "stikk!"?
			var trigger3 = (bibsys.getCurrentLine('lower').indexOf('stikk!') !== -1);
			
			// Automatisk stikkseddel etter utlån (til bibliotek)?
			var trigger4 = (bibsys.get(1,1,14) === 'Lån registrert' &&
					($.bibduck.config.autoStikkEtterReg === 'autostikk_reg_alle' ||
					$.bibduck.config.autoStikkEtterReg === 'autostikk_reg_lib' && bibsys.get(1, 20, 22) == 'lib')
					);
				
				/* Automatisk stikskeddel etter hentebeskjed fra 
				trigger5 = ($.bibduck.config.autoStikkEtterRes === true &&
							bibsys.get(1).indexOf('Hentebeskjed er sendt') !== -1 &&
							(bibsys.get(2, 1, 12) === 'Motta innlån'));*/


			if (this.waiting === false && (trigger3 || trigger4)) {
				this.waiting = true;
				var that = this;
				setTimeout(function(){
					if (trigger3) bibsys.clearInput();
					if (!trigger3) $.bibduck.log('[STIKK] >>> Automatisk stikkseddel <<<', 'info');
	
					if (bibsys.busy) {
						$.bibduck.log('[STIKK] Bibsys-vinduet er opptatt', 'error');
						//bibsys.alert("Bibsys-vinduet er opptatt. Om problemet vedvarer kan du omstarte BIBDUCK.");
						return;
					}
					if (bibsys.confirm('Vil vi ha stikkseddel?', 'Stikkseddel?')) {
						bibsys.setBusy(true);
						that.lag_stikkseddel(bibsys, function() {
							$.bibduck.log('[STIKK] Ferdig');
							bibsys.setBusy(false);
						});
					}
				}, 250); // add a small delay
			} else if (this.waiting === true && !trigger3 && !trigger4) {
				this.waiting = false;
			}
		}

	});

})();
