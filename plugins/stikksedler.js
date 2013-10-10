/*****************************************************************************
 * <stikksedler.js>
 * Modul for å skrive ut stikksedler ved hjelp av Excel-maler
 * Av: Bård S. Tuseth (c) 2009
 *     Fredrik Hovind Juell (c) 2010
 *     Dan Michael O. Heggø (c) 2013
 *
 * Nye kommandoer:
 *   stikk!     : Skriver stikkseddel
 *****************************************************************************/

// Stikkseddel-objekt 
$.bibduck.stikksedler = {

	// Settes under Innstillinger i brukergrensesnittet
	beststed: '',
	libnr: '',
	template_dir: '',

	// Formaterer dagens dato
	current_date: function() {
		var today = new Date(),
			dd = today.getDate(),
			mm = today.getMonth() + 1,  //January is 0!
			yyyy = today.getFullYear();
		if (dd < 10) {
			dd = '0' + dd;
		}
		if(mm < 10) {
			mm = '0' + mm;
		}
		return yyyy + '-' + mm + '-' + dd;
	},
	
	// Utlånseddel
    reg: function(doc, user, library) {
		var fil = user.spraak === 'ENG' 
			? 'reg_en.xls' 
			: 'reg_no.xls';
		this.ferdiggjor(doc, user, library, fil);
    },
	
    // Returseddel
    ret: function(doc, user, library) {
		// Hvis biblioteket vi returnerer til har navn "xxx" er 
		// det retur til utlandet. Da skriver vi ut en stikkseddel på engelsk.
		var fil = library.navn === 'xxx' 
			? 'ret_en.xls' 
			: 'ret_no.xls';
		this.ferdiggjor(doc, user, library, fil);		
    },
	
    // Avhentingsseddel for utlån
    avh: function (doc, user, library) {
		var fil = 'avh.xls'; 
        this.ferdiggjor(doc, user, library, fil);
    },

	// Avhentingsseddel for artikkelkopier
    avh_copy: function (doc, user, library) {
		var fil = 'avh_copy.xls'; 
        this.ferdiggjor(doc, user, library, fil);
    },

    // Seddel for reserverte dokumenter som skal til annet UBO-bibliotek
    res: function (doc, user, library) {
        var fil = 'res.xls'; 
        this.ferdiggjor(doc, user, library, fil);
    },

	// Ferdigstill, skriv ut og rydd opp
	ferdiggjor: function(doc, user, library, fil) {
		var libnr = $.bibduck.config.libnr,
			path = this.template_dir + fil,
			fso = new ActiveXObject("Scripting.FileSystemObject");
		
		// Sjekk at Excel-malfilen finnes:
		if (!fso.FileExists(path)) {
			$.bibduck.log('Stikkseddelfilen "' + path + '" finnes ikke!', 'error');
			return;
		}
		
		// Last inn malfilen:
        var excel = this.load_xls_template(path);
		
		// Utvid malsyntaks
        this.process_template_replacements(doc, user, library, excel);

		// Skriv ut...
        this.excel.ActiveWorkbook.PrintOut();

		// ... og rydd opp
		this.excel.ActiveWorkbook.Close(0);
		this.excel.Quit();
		delete this.excel;
		this.excel = undefined;
    },

	// Formaterer datoer på norsk og engelsk
    format_date: function(dt, lang) {
        if (dt === undefined) return '';
        var fdato = dt.split('-');
        if (lang === 'ENG') {
            return fdato[2] + '. ' + month_names_en[fdato[1]-1] + ' ' + fdato[0];
        } else {
            return fdato[2] + '. ' + month_names[fdato[1]-1] + ' ' + fdato[0];
        }
    },
	
	// Laster inn en Excel-malfil
	load_xls_template: function (filename) {
		var printerStr = $.bibduck.config.printerName + ' on ' + $.bibduck.config.printerPort;
		this.excel = new ActiveXObject('Excel.Application');
		this.excel.Visible = false;
		$.bibduck.log(getCurrentDir() + filename);
		this.excel.Workbooks.Open(getCurrentDir() + filename);
		if ($.bibduck.config.printerPort === '') {
			$.bibduck.log('Ingen stikkseddelskriver satt. Bruker standardskriver');
		} else if ($.bibduck.config.printerPort === 'none') {
			// bruk standardskriver
		} else {
			try {
				this.excel.Application.ActivePrinter = printerStr;
			} catch (e) {
				$.bibduck.log('Klarte ikke sette skriver. Bruker standardskriver', 'warn');
			}
		}
		return this.excel;
	},

    // Utvider malsyntaks i Excel-malfilen
    process_template_replacements: function (doc, user, library, excel) {
        var cells = new Enumerator(excel.ActiveSheet.UsedRange.Cells),
            cell,
            libv = '',
            libh = '',
            navn = user.etternavn + ', ' + user.fornavn,
			libnavn = '';

        if (doc.utlstatus !== 'AVH') {
            if (user.kind === 'bibliotek') {

				// Låner er et bibliotek: Fjernlån
				libv = user.ltid.substr(3,3);
                libh = user.ltid.substr(6);
                navn = 'Fjernlån';  // til ' + user.navn;
				libnavn = user.navn;

			} else if (user.beststed !== this.beststed && !library.gangavstand) {
			
				// Sendes
                libv = library.ltid.substr(3,3);    // Venstre del av lib-nr.
                libh = library.ltid.substr(6);      // Høyre del av lib-nr.
				libnavn = library.navn;

			} else {
			
				// Sendes ikke

			}
        }
        if (doc.utlaansdato === undefined) doc.utlaansdato = this.current_date();
        if (doc.forfallsdato === undefined) doc.forfallsdato = this.current_date();
        if (doc.forfvres === undefined) doc.forfvres = this.current_date();
        if (user.spraak === undefined) user.spraak = '';
		
		var infoEgenfornyningLinje1 = '',
			infoEgenfornyningLinje1 = '';
		
		if (user.spraak === 'ENG') {
			// Hvis ikke fjernlån, skriv ut litt ekstra info om fornying:
			if (user.kind === 'person') {
				if (doc.purretype === 'E') {
					if (doc.utlstatus === 'UTL/RES') {
						infoEgenfornyningLinje1 = "Please note:";
						infoEgenfornyningLinje2 = "This document can not be renewed as it has been reserved by someone else.";
					} else {
						infoEgenfornyningLinje1 = "This document can not be renewed online at BIBSYS Ask.";
						infoEgenfornyningLinje2 = "Please visit the library if you want to renew it.";
					}
				} else {
					infoEgenfornyningLinje1 = "Unless requested by someone else, this document can be";
					infoEgenfornyningLinje2 = "renewed online at BIBSYS Ask.";
				}
			}

		} else {
			// Skal boka til et annet bibliotek innad i organisasjonen?
			// Hvis ikke fjernlån, skriv ut litt ekstra info om fornying:
			var infoEgenfornyningLinje1 = '',
				infoEgenfornyningLinje2 = '';
			if (user.kind === 'person') {
				if (doc.purretype === 'E') {
					if (doc.utlstatus === 'UTL/RES') {
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
		
		var forfallVedRes = this.format_date(doc.forfvres, user.spraak);
		if (forfallVedRes) {
			forfallVedRes = 'Ved reservasjoner kan documentet bli innkalt fra ' + forfallVedRes;
		}

        for (; !cells.atEnd(); cells.moveNext()) {
            cell = cells.item();
            if (cell.Value !== undefined && cell.Value !== null) {
                cell.Value = cell.Value
						.replace('{{Navn}}', navn)
						.replace('{{Ltid}}', user.ltid ? user.ltid : '-')
						.replace('{{Tittel}}', doc.tittel ? doc.tittel : '-')
						.replace('{{Dokid}}', doc.dokid ? doc.dokid : '-')
						.replace('{{Bestnr}}', doc.bestnr ? doc.bestnr : '-')
						.replace('{{Utlånsdato}}', this.format_date(doc.utlaansdato, user.spraak))
						.replace('{{Forfallsdato}}', this.format_date(doc.forfallsdato, user.spraak))
						.replace('{{ForfallVedRes}}', forfallVedRes)
						.replace('{{DagensDato}}', this.format_date(this.current_date(), user.spraak))
						//.replace('{{Dato}}', this.format_date(this.current_date()))
						.replace('{{Libnavn}}', libnavn)
						.replace('{{LIBV}}', libv)
						.replace('{{LIBH}}', libh)
						.replace('{{Bestnr}}', doc.bestnr)
						.replace('{{Hentenr}}', doc.hentenr)
						.replace('{{InfoEgenfornyingLinje1}}', infoEgenfornyningLinje1)
						.replace('{{InfoEgenfornyingLinje2}}', infoEgenfornyningLinje2);
            }
        }

        // Hvis forfallsdato ved reservasjon er lik ordinær forfallsdato:
        if (doc.forfvres === doc.forfallsdato) {
            excel.Cells(4, 2).Value = '';
        }
    }

};

(function() {
	var worker,
		client,
		dok = {},
		laaner = {},
		lib = {},
		hjemmebibliotek = '',
		current_date = '',
		config,
		seddel,
		callback,
		working = false,
		siste_bestilling = { active: false },
		fso = new ActiveXObject("Scripting.FileSystemObject"),
			shell = new ActiveXObject("WScript.Shell"),
			appdata = shell.ExpandEnvironmentStrings("%ALLUSERSPROFILE%"),
			stikk_path = appdata + '\\Scriptotek\\Bibduck\\stikk.txt';

	function setWorking(working) {
		working = working;
		//$('#btn-stikkseddel').prop('disabled', working);
	}

	function les_dokstat_skjerm() {

		if (client.get(2, 1, 28) !== 'Utlånsstatus for et dokument') {
			client.alert('Vi er ikke på DOKST-skjermen :(');
			setWorking(false);
			return;
		}

		// Sjekker hvilken linje tittelen står på:
		if (client.get(7, 2, 7) == 'Tittel') {
			// Lån fra egen samling
			dok.tittel = client.get(7, 14, 80).trim();
		} else if (client.get(8, 2, 7) == 'Tittel') {
			// ik...
			dok.tittel = client.get(8, 13, 80).trim();
		} else {
			// Relativt sjelden case? Linje 7-10 er fritekst, og 
			// tittel og forfatter bytter typisk mellom linje 7 og 8.
			// En enkel test, som sikkert vil feile i flere tilfeller:
			var tittel1 = client.get(7, 2, 80).trim(),
				tittel2 = client.get(8, 2, 80).trim();
			if (tittel1.length > tittel2.length) {
				dok.tittel = tittel1;
			} else {
				dok.tittel = tittel2;
			}
		}

		dok.dokid = client.get( 6, 31, 39);

		if (dok.dokid === '') {
			client.alert('Har du husket å trykke enter?');
			setWorking(false);
			return;
		}

		dok.utlstatus    = client.get( 3, 46, 65);   // AVH, RES, UTL, UTL/RES, ...

		// Dokument til avhenting?
		if (dok.utlstatus === 'AVH') {

			dok.hentenr = client.get(1, 44, 50);
			dok.hentefrist = client.get(1, 26, 35);

		} else {

			laaner.ltid      = client.get(14, 11, 20);
			dok.utlaansdato  = client.get(18, 18, 27);   // Utlånsdato
			dok.forfvres     = client.get(20, 18, 27);   // Forfall v./res
			dok.forfallsdato = client.get(21, 18, 27);   // Forfallsdato
			dok.purretype    = client.get(17, 68, 68);
			dok.kommentar    = client.get(23, 17, 80).trim();

			//Tester om låntaker er et bibliotek:
			if (laaner.ltid.substr(0,3) == 'lib') {
				laaner.kind = 'bibliotek';
				laaner.navn = client.get(14, 22, 79).trim();
			} else {
				laaner.kind = 'person';
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

		if (dok.utlstatus === 'AVH') {

			// Vi trenger ikke mer informasjon. 
			// La oss kjøre i gang Excel-helvetet, joho!!
			seddel.avh(dok, laaner, lib);
			emitComplete();

		} else if (dok.utlstatus === 'RES') {

			// Dokument som *kun* er reservert 
			// Finn låneren i reservasjonslista:
			$.bibduck.log('Går til RLIST for å identifisere låneren', 'debug');
			worker.send('rlist,\n');
			worker.wait_for('Hentefrist:', [6,5], function() {
				var resno = -1;
				if (worker.get(3, 63, 71) === dok.dokid) {
					resno = 1;
					laaner.ltid = worker.get(3, 15, 24);
				} else if (worker.get(10, 63, 71) === dok.dokid) {
					resno = 2;
					laaner.ltid = worker.get(10, 15, 24);
				} else if (worker.get(17, 63, 71) === dok.dokid) {
					resno = 3;
					laaner.ltid = worker.get(17, 15, 24);
				}
				$.bibduck.log('Hvilken reservasjon på RLIST-skjermen? Bruker nummer ' + resno + ' fordi den har dokid ' + dok.dokid, 'info');

				// Gå til dokst:
				$.bibduck.log('Sender F12 for å gå til DOkstat', 'debug');
				$.bibduck.sendSpecialKey('F12');
				worker.wait_for('DOkstat', [2,31], function() {
					worker.resetPointer();

					// Vi trenger mer info om låneren:
					$.bibduck.log("Går til LTSØK for å finne ut mer om låneren", 'debug');
					worker.send('ltsø,' + laaner.ltid + '\n');
					worker.wait_for('Fyll ut:', [5,1], function() {
						// Vi sender enter på nytt
						worker.send('\n');
						worker.wait_for('Sist aktiv dato', [22,1], les_ltsok_skjerm);
					});
				});
			});

		} else if (laaner.kind === 'person') {

			// Vi trenger mer info om låneren:
			worker.send('ltsø,' + laaner.ltid + '\n');
			worker.wait_for('Fyll ut:', [5,1], function() {
				// Vi sender enter på nytt
				worker.send('\n');
				worker.wait_for('Sist aktiv dato', [22,1], les_ltsok_skjerm);
			});

		} else {

			// Vi trenger ikke mer informasjon. 
			// La oss kjøre i gang Excel-helvetet, joho!!
			seddel.reg(dok, laaner, lib);
			emitComplete();


		}

	}

	function emitComplete() {
		$.bibduck.log("Stikkseddel ferdig");
		setWorking(false);
		if (callback !== undefined) {
			setTimeout(function() { // a slight delay never hurts
				var data = {
					patron: laaner,
					library: lib,
					document: dok,
					beststed: seddel.beststed
				};
				callback(data);
			}, 200);
		}
	}
	
	/*
	 * Prosedyre for å sende hentebeskjed fra DOKST-skjermen
	 */
	function send_hentebeskjed() {

		if (client.get(2, 1, 28) !== 'Utlånsstatus for et dokument') {
			$.bibduck.log('send_hentebeskjed: Vi er ikke på DOKST-skjermen!', 'error');
			client.alert('send_hentebeskjed: Vi er ikke på DOKST-skjermen!');
			setWorking(false);
			return;
		}

		client.send('\thentb,\n');
		client.wait_for('Hentebrev til låntaker:', [7,15], function() {
			//client.send(dok.dokid + '\n');
			client.send('\t' + laaner.ltid + '\n');


			client.wait_for([

				['Kryss av for ønsket valg', [16,8], function() {
					send_hentb_steg2();
				}],

				['Ønsker du likevel å låne ut boka?', [19,2], function() {
					// Dekker følgende:
					//    'Ugyldig LTID fra dato', [9,2]
					//    'Låntaker har: 1 erstatningskrav, 0 er max.grense' [11,2]
					//    'STOPPMELDING', [6,15]
					// Dekker trolig (ikke testet):
					//    'sistegangspurringer'
					//    'i utestående gebyr'

					if (!client.confirm('Ønsker du å fortsette?', 'Sende hentebeskjed')) {
						return;
					}

					// J og <Enter> for å fortsette
					client.send('J\n');
					client.wait_for('Kryss av for ønsket valg', [16,8], function() {
						send_hentb_steg2();
					});
				}],
								
				['KOMMENTAR', [2,22], function() {
					
					if (!client.confirm('Ønsker du å fortsette?', 'Sende hentebeskjed')) {
						return;
					}

					// <Enter> for å fortsette
					client.send('\n');
					client.wait_for('Kryss av for ønsket valg', [16,8], function() {
						send_hentb_steg2();
					});

				}]

			]);

		});
	
	}

	function send_hentb_steg2() {
		client.send('X\n');
		client.wait_for([
			['Hentebeskjed er sendt', [1,1], function() {
				$.bibduck.log('Hentebeskjed sendt per sms');
				client.resetPointer();
				hentebeskjed_sendt();
			}],
			['Registrer eventuell melding', [8,5], function() {
				$.bibduck.sendSpecialKey('F9');
				client.wait_for('Hentebeskjed er sendt', [1,1], function() {
					$.bibduck.log('Hentebeskjed sendt per epost');
					hentebeskjed_sendt();
				});
			}]
		]);
	}
	
	function hentebeskjed_sendt() {
		var firstline = client.get(1),
			m = firstline.match(/på (sms|Email) til (.+) merket (.+)/);
		
		var	name = m[2],
			nr = m[3].trim();
		if (nr === '') {
			$.bibduck.log('Fant ikke noe hentenr.', 'error');
			setWorking(false);
			return;
		}
		//$.bibduck.log(name + nr + siste_bestilling.laankopi);

		dok.hentenr = nr;
		dok.hentefrist = '-';

		// Vi trenger ikke mer informasjon.
		// La oss kjøre i gang Excel-helvetet, joho!!
		if (siste_bestilling.laankopi == 'K') {
			$.bibduck.log('Hentenr.: ' + nr + ' (kopibestilling)', 'info');
			seddel.avh_copy(dok, laaner, lib);
		} else {
			$.bibduck.log('Hentenr.: ' + nr + ' (lånebestilling)', 'info');
			seddel.avh(dok, laaner, lib);
		}

		emitComplete();

		client.send('dokst,\n'); // for å oppfriske skjermen slik at status endres fra RES til AVH

	}

	function res_sendes() {

		if (worker.get(2, 1, 24) !== 'Opplysninger om låntaker') {
			client.alert("Vi er ikke på LTSØ-skjermen :(");
			setWorking(false);
			return;
		}

		// Vi avslutter med å gå til RLIST igjen for å skrive kommentar
		client.send('\tRLIST,' + dok.dokid + '\n');
		client.wait_for('Hentefrist', [20,5], function () {

			// Hvilket bibliotek skal dokumentet sendes til? Finn signatur for LIBNR
			for (var s in config.sigs) {
				if (config.sigs[s] === lib.ltid) {
					sig = s;
				}
			}

			var checkPage = function(pageNo) {
				var fndEntry = false,
					tabs = '\t';
				for (var line = 3; i <= 17; i+=7) {
					if (client.get(line, 15, 24) == laaner.ltid) {
						client.send(tabs + 'Sendt ' + sig + ' ' + $.bibduck.stikksedler.current_date());
						fndEntry = true;
						break;
					}
					tabs += '\t\t';
				}
				if (fndEntry) {
					// Vi fant dokid på denne siden

					client.alert('Obs! Låner har bestillingssted ' + laaner.beststed +
						', så dokumentet må sendes. Du skal få en stikkseddel.');

					emitComplete();
					seddel.res(dok, laaner, lib);

				} else {
					// Vi fant ikke dokid på denne siden. Vi sjekker neste side hvis den finnes
					if (client.get(25,49,51) !== '') {
						$.bibduck.sendSpecialKey('F8');
						var firstEntryOnNextPage = (pageNo*3 + 1);
						firstEntryOnNextPage = (firstEntryOnNextPage < 10) ? ' ' + firstEntryOnNextPage : '' + firstEntryOnNextPage;
						client.wait_for(firstEntryOnNextPage, [3, 6], function() { checkPage(pageNo+1); });
					} else {
						$.bibduck.log('Fant ikke ltid ' + laaner.ltid + ' i reservasjonslisten!', 'error');
					}
				}

			};

			checkPage(1);

		});

	}

	function les_ltsok_skjerm() {
		var that = this;
		if (worker.get(2, 1, 24) !== 'Opplysninger om låntaker') {
			client.alert("Vi er ikke på LTSØ-skjermen :(");
			setWorking(false);
			return;
		}
		laaner.beststed  = worker.get( 7, 71, 80).trim();
		laaner.etternavn = worker.get( 5, 18, 58).trim();
		laaner.fornavn   = worker.get( 6, 18, 58).trim();
		laaner.spraak    = worker.get(19, 41, 44).trim();

		// DEBUG:
		/*
		$.bibduck.log('Info om låner:');
		$.each(laaner, function(k,v) {
			$.bibduck.log('  ' + k + ': ' + v);
		});*/

		lib.ltid = '';
		lib.navn = '';
		if (laaner.beststed in config.bestillingssteder) {
			lib.ltid = config.bestillingssteder[laaner.beststed];
			$.bibduck.log(lib.ltid);
		} else {
			// En bruker med lånekort fra f.eks. tek (NTNU) 
			// som vi kobler, vil beholde beststed tek.
			$.bibduck.log("Kjenner ikke libnr for bestillingssted: " + laaner.beststed, 'warn');
			// return;
		}
		if (lib.ltid in config.biblnavn) {
			lib.navn = config.biblnavn[lib.ltid];
		} else if (lib.ltid !== '') {
			$.bibduck.log("Kjenner ikke navn for libnr: " + lib.ltid, 'warn');
		}
		
		lib.gangavstand = false;

		if (config.gangavstand[seddel.libnr]) {
			for (var key in config.gangavstand[seddel.libnr]) {
				if (config.gangavstand[seddel.libnr][key] == lib.ltid) {
					lib.gangavstand = true;
					$.bibduck.log('Låner har bestillingssted ' + lib.ltid + ', som er innen gangavstand fra ' + seddel.libnr + ', så vi sender ikke boka.', 'info');
				}
			}
		}

		// DEBUG:
		/*
		$.bibduck.log('Info om bibliotek:');
		$.each(lib, function(k,v) {
			$.bibduck.log('  ' + k + ': ' + v);
		});*/

		// Nå har vi informasjonen vi trenger. La oss kjøre i gang Excel-helvetet, joho!!

		// @TODO: Hva med UTL/RES ?
		if (dok.utlstatus === 'RES') {

			res_sendes();

		} else if (dok.utlstatus === 'AVH') {

			if (laaner.beststed == seddel.beststed || lib.gangavstand) {				
			
				dok.utlstatus = 'AVH';

				client.send('dokst,' + dok.dokid + '\n');
				client.wait_for('Utlkommentar', [23,1], function() {
					// dok.dokid fra rlist kan være et knyttid. Vi overskriver derfor med
					// det virkelige dokid-et.
					dok.dokid = client.get(6,31,39);
					send_hentebeskjed();
				});				

			} else {

				dok.utlstatus = 'RES';

				res_sendes();
			
			}

		} else {

			// Gi beskjed hvis boka skal ut av huset
			if (laaner.kind === 'person' && laaner.beststed !== seddel.beststed && lib.ltid !== '') {
				client.alert('Obs! Låner har bestillingssted: ' + laaner.beststed);
				$.bibduck.log('NB! Låner har et eksternt bestillingssted: ' + laaner.beststed + ' (' + lib.ltid + ')', 'warn');

				// Hvis boken skal sendes, så gå til utlånskommentarfeltet.
				client.send('en,' + dok.dokid + '\n');
				client.wait_for('Utlmkomm:', [8,1], function() {
					client.send('\t\t\t');
					seddel.reg(dok, laaner, lib);
					emitComplete();
				});

			// Hvis ikke går vi tilbake til dokst-skjermen:
			} else {

				if (laaner.kind === 'person' && laaner.beststed !== seddel.beststed) {
					$.bibduck.log('NB! Låner har et eksternt bestillingssted: ' + laaner.beststed, 'warn');
				}

				//result = snt.MessageBox("Vil du gå til REG for å låne ut flere bøker?", "Error", ICON_QUESTION Or BUTTON_YESNO Or DEFBUTTON2)

				//if (result == IDYES) {
				//  // ... tilbake til utlånsskjerm for å registrere flere utlån.
				//  snt.Send("reg,"+ltid)
				//  snt.QuickButton("^M")
				//Else
					// ... tilbake til dokst, for å sende hentebeskjed
					client.send('dokst,' + dok.dokid + '\n');
					client.wait_for('Utlkommentar', [23,1], function() {
						// FINITO, emit
						seddel.reg(dok, laaner, lib);
						emitComplete();
					});
				//}
			}
		}
	}

	function start_from_res() {
		/*
		 * Reservere (RES)                                                    BIBSYS UTLÅN
		 * Gi kommando:                      :                                  2013-06-27
		 * 
		 *      LTID:             :  DOKID/REFID/HEFTID/INNID: 96nf00169 :
		 *      Reskomm:                                                               :
		 *      Resreferanse:                                 :
		 *      Volum:                 År:            Hefte:                           :
		 * ----------------------------- 96nf00169 ---------------------------------------
		 *  Forfatter : Auyang, Sunny Y.
		 *  Tittel    : How is quantum field theory possible? / Sunny Y. Auyang.
		 *  Trykt     : New York : Oxford University Press, 1995.Finnes også som:
		 * 
		 *  Signatur  : UREAL Fys. 0.2 AUY eks. 2
		 * 
		 * 
		 * 
		 * -------------------------------------------------------------------------------
		 *                   ubo0292451  Dan Michael Olsen Heggø
		 *                   Nr. 1 på reserveringslista.
		 */
		laaner = { kind: 'person' };
		lib = {};
		dok = { utlstatus: 'RES' };
		if (client.get(2, 1, 15) !== 'Reservere (RES)') {
			$.bibduck.log('Ikke på reserveringsskjermen', 'error');
			setWorking(false);
			return;
		}

		if (client.get(1, 1, 12) === 'Hentebeskjed') {
			dok.utlstatus = 'AVH';
		}

		if (client.get(1, 1, 12) !== 'Hentebeskjed' && client.get(20, 19, 21) !== 'Nr.') {
			$.bibduck.log('Ingen reservering gjennomført, kan ikke skrive ut stikkseddel', 'error');
			client.alert('Du må gjennomføre en reservering før du kan skrive ut stikkseddel');
			setWorking(false);
			return;
		}

		dok.tittel = '';
		if (dok.utlstatus === 'AVH') {

			laaner.ltid = client.get(5, 12, 22);
			dok.dokid = client.get(5, 53, 61);

			// Gå til dokst:
			$.bibduck.log('Går til DOkstat vha. F12', 'debug');
			$.bibduck.sendSpecialKey('F12');
			client.wait_for('DOkstat', [2,31], function() {

				if (client.get(6,31,39) === dok.dokid) {
					les_dokstat_skjerm();
				} else {
					$.bibduck.log('Feil dokid. Ber om dokstat for dokid ' + dok.dokid, 'debug');
					client.send(dok.dokid + '\n');
					client.wait_for(dok.dokid, [6,31], les_dokstat_skjerm);
				}

			});
		} else {
			laaner.ltid = client.get(19, 19, 28);
			dok.dokid = client.get(9, 31, 39);

			if (client.get(10, 2, 7) == 'Tittel') {
				dok.tittel = client.get(10, 14, 79);
			} else if (client.get(11, 2, 7) == 'Tittel') {
				dok.tittel = client.get(11, 14, 79);
			} else if (client.get(12, 2, 7) == 'Tittel') {
				dok.tittel = client.get(12, 14, 79);
			} else if (client.get(13, 2, 7) == 'Tittel') {
				dok.tittel = client.get(13, 14, 79);
			}

			// Vi trenger mer info om låneren:
			worker.resetPointer();
			$.bibduck.log('Gjør LTSØk for ' + laaner.ltid, 'debug');
			worker.send('ltsø,' + laaner.ltid + '\n');
			worker.wait_for('Fyll ut:', [5,1], function() {
				// Vi sender enter på nytt
				worker.send('\n');
				worker.wait_for('Sist aktiv dato', [22,1], les_ltsok_skjerm);
			});
		}
	}

	function start_from_imo() {
		laaner = { kind: 'person' };
		lib = {};
		dok = { utlstatus: 'AVH', bestnr:  siste_bestilling.bestnr };
		var firstline = client.get(1);
		var tilhvem = firstline.match(/på (sms|Email) til (.+) merket (.+)/);
		if (tilhvem[2] == undefined) {
			client.alert("Oi, BIBSYS har ikke laget noe hentenummer til oss.");
			$.bibduck.log("Ikke noe hentenummer på skjermen", "error");
			return;
		}
		var name = tilhvem[2];
		var nr = tilhvem[3].trim();
		if (nr === '') {
			$.bibduck.log('Fant ikke noe hentenr.', 'error');
			setWorking(false);
			return;
		}

		dok.hentenr = nr;
		dok.hentefrist = '-';

		// Vi trenger ikke mer informasjon.
		// La oss kjøre i gang Excel-helvetet, joho!!
		if (siste_bestilling.laankopi == 'K') {
			$.bibduck.log('Hentenr.: ' + nr + ' (kopibestilling)', 'info');
			seddel.avh_copy(dok, laaner, lib);
		} else {
			$.bibduck.log('Hentenr.: ' + nr + ' (lånebestilling)', 'info');
			seddel.avh(dok, laaner, lib);
		}
		emitComplete();

	}

	function start_from_rlist() {
		laaner = { kind: 'person' };
		lib = {};
		dok = {};
		var resno = -1;

		if (client.get(2, 1, 25) !== 'Reserveringsliste (RLIST)') {
			$.bibduck.log('Ikke på rlist-skjerm', 'error');
			setWorking(false);
			return;
		}

		var firstline = client.get(1);
		if (firstline.indexOf('Hentebeskjed er sendt') !== -1) {
			var tilhvem = firstline.match(/på (sms|Email) til (.+) merket/);
			if (client.get(4).match(tilhvem[2])) {
				resno = 1;
			} else if (client.get(11).match(tilhvem[2])) {
				resno = 2;
			} else if (client.get(18).match(tilhvem[2])) {
				resno = 3;
			} else {
				$.bibduck.log('Fant ikke "' + tilhvem[2] + '" på RLIST-skjermen!', 'error');
				client.alert("Beklager, klarte ikke å identifisere hvilken oppføring vi skal bruke.");
				setWorking(false);
				return;
			}
			$.bibduck.log('Hvilken reservasjon på RLIST-skjermen? Bruker nummer ' + resno + ' fordi hentebeskjed ble sendt til "' + tilhvem[2] + '"', 'info');
		} else {
			var lineno = client.getCurrentLineNumber();
			if (lineno === 8) {
				resno = 1;
			} else if (lineno === 15) {
				resno = 2;
			} else if (lineno === 22) {
				resno = 3;
			} else {
				client.alert("Du må stå i et ref.-felt");
				setWorking(false);
				return;
			}
			$.bibduck.log('Hvilken reservasjon på RLIST-skjermen? Bruker nummer ' + resno + ' basert på hvilket ref.-felt som har fokus.', 'info');
		}
		if (resno === 1) {
			if (client.get(3,1,1) === 'A') {
				dok.utlstatus = 'AVH';
			} else {
				dok.utlstatus = 'RES';
			}
			laaner.ltid = client.get(3, 15, 24);
			laaner.beststed = client.get(3, 47, 54);
			dok.dokid = client.get(3, 63, 71);
		} else if (resno === 2) {
			if (client.get(10,1,1) === 'A') {
				dok.utlstatus = 'AVH';
			} else {
				dok.utlstatus = 'RES';
			}
			laaner.ltid = client.get(10, 15, 24);
			laaner.beststed = client.get(10, 47, 54);
			dok.dokid = client.get(10, 63, 71);

		} else if (resno === 3) {
			if (client.get(17,1,1) === 'A') {
				dok.utlstatus = 'AVH';
			} else {
				dok.utlstatus = 'RES';
			}
			laaner.ltid = client.get(17, 15, 24);
			laaner.beststed = client.get(17, 47, 54);
			dok.dokid = client.get(17, 63, 71);

		} else {
			client.alert("Flytt pekeren til Ref.-feltet for reserveringen du ønsker å skrive ut stikkseddel for og prøv på nytt.");
			setWorking(false);
			return;
		}
		if (dok.dokid === '') {
			client.alert("Reservering nummer " + resno + " er tom. Flytt pekeren til Ref.-feltet for reserveringen du ønsker å skrive ut stikkseddel for og prøv på nytt.");
			setWorking(false);
			return;
		}

		// Dokid i selve lista blir ikke oppdatert før skjermen lastes inn på nytt
		// Vi bruker istedet dokid som er scannet/skrevet inn
		dok.dokid = client.get(2, 31,39);

		$.bibduck.log('  Dokid: ' + dok.dokid + '. Ltid: ' + laaner.ltid, 'info');

		dok.tittel = '';

		// Gå til dokst:
		$.bibduck.log('Går til DOkstat vha. F12');
		$.bibduck.sendSpecialKey('F12');
		client.wait_for('DOkstat', [2,31], function() {
			if (client.get(6,31,39) === dok.dokid) {
				les_dokstat_skjerm();
			} else {
				$.bibduck.log('Feil dokid. Ber om dokstat for dokid ' + dok.dokid, 'debug');
				client.send(dok.dokid + '\n');
				client.wait_for(dok.dokid, [6,31], les_dokstat_skjerm);
			}

		});

	}

	function utlaan() {
		laaner = {};
		lib = {};
		dok = {};
		if (client.get(2, 1, 22) == 'Registrere utlån (REG)') {
			var dokid = client.get(10, 7, 15);
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
		
		$.bibduck.log('start from ltsøk');

		laaner = {};
		lib = {};
		dok = {};
		
		if (info !== undefined) {
			
			// Vi har mottatt en stikkseddelforespørsel fra en annen prosess
			dok.dokid = info.dokid;
			laaner.ltid = info.ltid;
			dok.utlstatus = 'AVH';
			les_ltsok_skjerm();
			
		} else {
		
			// Vi skriver ut en retur-seddel. Nyttig f.eks. hvis 
			// man ikke får stikkseddel fra IRET

			if (client.get(18, 18, 20) !== 'lib') {
				client.alert("Beklager, kan ikke skrive returseddel når låntakeren ikke er et bibliotek.");
				$.bibduck.log("Kan ikke skrive returseddel når låntakeren ikke er et bibliotek.", 'warn');
				setWorking(false);
				return;
			}

			laaner.ltid = client.get(18, 18, 27);
			laaner.navn = client.get(10, 18, 50);
			laaner.kind = 'bibliotek';
			lib.ltid = laaner.ltid;
			lib.navn = laaner.navn;

			seddel.ret(dok, laaner, lib);
			emitComplete();
		}
	}

	function retur() {
		worker.resetPointer();

		laaner = {};
		lib = {};
		dok = {};

		if (client.get(2).indexOf('IRETur') !== -1) {

			dok.dokid = client.get(1, 1, 9);
			dok.bestnr = client.get(4, 49, 57);

			laaner.ltid = client.get(6, 15, 24);
			laaner.navn = client.get(7, 20, 50);
			laaner.kind = 'bibliotek';
			lib.ltid = client.get(6, 15, 24);
			lib.navn = client.get(7, 20, 50);
			if (laaner.navn === 'xxx') {
				laaner.navn = '';
				laaner.navn = '';
				lib.ltid = '';
				lib.navn = '';
			}

		} else {

			// Retur til annet bibliotek innad i organisasjonen

			var sig = client.get(11, 14, 40).split(' ')[0];
			dok.dokid = client.get(6, 31, 39);
			dok.bestnr = '';

			if (sig in config.sigs) {
				lib.ltid = config.sigs[sig];
				lib.navn = config.biblnavn[lib.ltid];
			} else {
				client.alert('Beklager, BIBDUCK kjenner ikke igjen signaturen "' + sig + '".');
				setWorking(false);
				return;
			}

		}

		dok.tittel = '';
		if (client.get(7, 2, 7) == 'Tittel') {
			dok.tittel = client.get(7, 14, 79);
		} else if (client.get(8, 2, 7) == 'Tittel') {
			dok.tittel = client.get(8, 14, 79);
		} else if (client.get(9, 2, 7) == 'Tittel') {
			dok.tittel = client.get(9, 14, 79);
		} else if (client.get(10, 2, 7) == 'Tittel') {
			dok.tittel = client.get(10, 14, 79);
		}

		if (hjemmebibliotek === '') {
			client.alert('Libnr. er ikke satt. Dette setter du under Innstillinger.');
		}
		if (lib.ltid === 'lib' + hjemmebibliotek) {
			client.alert('Boka hører til her. Returseddel trengs ikke.');
			setWorking(false);
			return;
		}

		seddel.ret(dok, laaner, lib);
		emitComplete();

	}

	function start(info) {

		try {
			$('audio#ping').get(0).play();
		} catch (e) {
			// IE8?
		}
		$.bibduck.log('Lager stikkseddel');
		hjemmebibliotek = $.bibduck.config.libnr;
		seddel = $.bibduck.stikksedler;
		seddel.libnr = 'lib' + $.bibduck.config.libnr;
		seddel.template_dir = 'plugins\\stikksedler\\' + config.malmapper[seddel.libnr] + '\\';
		seddel.beststed = '';
		for (var key in config.bestillingssteder) {
			if (config.bestillingssteder[key] == seddel.libnr) {
				seddel.beststed = key;
			}
		}
		if (seddel.libnr === 'lib') {
			client.alert('Obs! Libnr. er ikke satt enda. Dette setter du under Innstillinger i Bibduck.');
			setWorking(false);
			return;
		} else if (seddel.beststed === '') {
			client.alert('Fant ikke et bestillingssted for biblioteksnummeret ' + seddel.libnr + ' i config.json!');
			setWorking(false);
			return;
		}

		if (client.get(2, 1, 22) === 'Registrere utlån (REG)') {
			utlaan();
		} else if (client.get(14, 1, 8) === 'Låntaker') { // DOkstat
			utlaan();
		} else if (client.get(15, 2, 13) === 'Returnert av') {
			retur();
		} else if (client.get(1).indexOf('er returnert') !== -1 && client.get(2).indexOf('IRETur') !== -1) { // Retur innlån (IRETur)
			retur();
		} else if (client.get(2, 1, 32) === 'Opplysninger om låntaker (LTSØk)') {
			start_from_ltsok(info);
		} else if (client.get(2, 1, 15) === 'Reservere (RES)') {
			start_from_res();
		} else if (client.get(2, 1, 25) === 'Reserveringsliste (RLIST)') {
			start_from_rlist();
		} else if (client.get(2, 1, 12) === 'Motta innlån') {
			start_from_imo();
		} else {
			setWorking(false);
			$.bibduck.log('Stikkseddel fra denne skjermen er ikke støttet', 'warn');
			client.alert('Stikkseddel fra denne skjermen er ikke støttet (enda). Ta DOKST og prøv igjen');
		}
	}

	$.bibduck.plugins.push({

		name: 'Stikkseddel-tillegg',

		initialize: function() {
			var that = this;
			/*
			$('#btn-stikkseddel').remove();
			var $btn = $('<button type="button" id="btn-stikkseddel">Stikkseddel</button>');
			$.bibduck.log("Legger til stikkseddelknapp");
			$('#header-inner').append($btn);
			$btn.on('click', function() {
				var bibsys = $.bibduck.getFocused();
				if (bibsys !== undefined) {
					bibsys.bringToFront();
					setTimeout(function() {
						that.lag_stikkseddel(bibsys);
					}, 250);
				}
			});
			*/
			that.listenForNotificationFile();
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

					$.bibduck.log(txt);
					var request = $.parseJSON(txt);

					bibsys = $.bibduck.getFocused();

					fso.DeleteFile(stikk_path);
					
					bibsys = $.bibduck.getFocused();
					
					$.bibduck.log('Fikk forespørsel om stikkseddel.' +
						(request.ltid ? 'Ltid: ' + request.ltid + ', dokid: ' + request.dokid : ''), 'info');

					that.forbered_stikkseddel(bibsys, function() {
						//$.bibduck.log('forbered_stikkseddel callback');
						bibsys.unidle();
						bibsys.update();   // force update
						start(request);
					});
				}
				setTimeout(check, 500);
			};
			setTimeout(check, 500);
		},

		lag_stikkseddel: function(bibsys, cb) {

			if (working) {
				$.bibduck.log("En stikkseddel er allerede under produksjon", "error");
				bibsys.alert("En stikkseddel er allerede under produksjon. Om problemet vedvarer kan du omstarte BIBDUCK.");
				return;
			}
			bibsys.off('waitFailed');
			bibsys.on('waitFailed', function() {
				$.bibduck.log('Stikkseddelutskriften ble avbrutt', 'error');
				setWorking(false);
			});
			callback = cb;

			setWorking(true);
			current_date = bibsys.get(3, 70, 79);
			
			this.forbered_stikkseddel(bibsys, start);
		},
		
		forbered_stikkseddel: function(bibsys, startfn) {

			client = bibsys;
			
			//$.bibduck.log('forbered_stikkseddel');
		
			//$.bibduck.log(current_date);
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
				$.bibduck.log('Load: plugins/stikksedler/config.json');
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
		 * Denne metoden kalles av Bibduck med noen hundre millisekunds mellomrom.
		 */
		update: function(bibsys) {

			// Vi må huske om siste mottatte bestilling var en kopi (K) eller et lån (L)
			// for å skrive ut rett stikkseddel (lån må lånes ut på automat, kopier ikke)
			if (bibsys.get(1, 11, 20) === 'er mottatt') {
				if (!siste_bestilling.active) {
					siste_bestilling = {
						bestnr: bibsys.get(1, 1, 9),
						laankopi: bibsys.get(8, 37, 37)
					};
				}
				siste_bestilling.active = true;
			} else if (siste_bestilling.active) {
				siste_bestilling.active = false;
			}


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
					if (!trigger3) $.bibduck.log('Lager stikkseddel automatisk (stikksedler.js)', 'info');
					that.lag_stikkseddel(bibsys);
				}, 250); // add a small delay
			} else if (this.waiting === true && !trigger3 && !trigger4) {
				this.waiting = false;
			}
		}

	});

})();
