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
$.bibduck.stikksedler = {

	// Settes under Innstillinger i brukergrensesnittet
	beststed: '',

	load_xls: function (filename) {
		var printerStr = $.bibduck.config.printerName + ' on ' + $.bibduck.config.printerPort;
		this.excel = new ActiveXObject('Excel.Application');
		this.excel.Visible = false;
		$.bibduck.log(getCurrentDir() + filename);
		this.excel.Workbooks.Open(getCurrentDir() + filename);
		this.excel.Application.ActivePrinter = printerStr;
		return this.excel;
	},

	print_and_close: function() {
		this.excel.ActiveWorkbook.PrintOut();
		this.excel.ActiveWorkbook.Close(0);
		this.excel.Quit();
		delete this.excel;
		this.excel = undefined;
		//$.bibduck.log('OK', {timestamp: false});
	},

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
		siste_bestilling = { active: false };

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
						worker.wait_for('Sist aktiv dato', [22,1], les_ltst_skjerm);
					});
				});
			});

		} else if (laaner.kind === 'person') {

			// Vi trenger mer info om låneren:
			worker.send('ltsø,' + laaner.ltid + '\n');
			worker.wait_for('Fyll ut:', [5,1], function() {
				// Vi sender enter på nytt
				worker.send('\n');
				worker.wait_for('Sist aktiv dato', [22,1], les_ltst_skjerm);
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

	function les_ltst_skjerm() {
		var that = this;
		$.bibduck.log(client.get(2, 1, 24));
		$.bibduck.log(worker.get(2, 1, 24));
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

		// DEBUG:
		/*
		$.bibduck.log('Info om bibliotek:');
		$.each(lib, function(k,v) {
			$.bibduck.log('  ' + k + ': ' + v);
		});*/

		// Nå har vi informasjonen vi trenger. La oss kjøre i gang Excel-helvetet, joho!!

		// @TODO: Hva med UTL/RES ?
		if (dok.utlstatus === 'RES') {

			var resno,
				kommentar,
				sig = '???';

			if (laaner.beststed == seddel.beststed) {
				client.alert('Obs! Låner har bestillingssted ' + laaner.beststed +
					', så det burde ikke være behov for å sende dokumentet.');
				setWorking(false);
				return;
			}
			client.send('rlist,' + dok.dokid + '\n');
			client.wait_for('Hentefrist:', [6,5], function() {
				if (worker.get(3, 63, 71) === dok.dokid) {
					resno = 1;
					$.bibduck.log("tab once");
					client.send('\t');
					kommentar = worker.get(9, 13, 79);
				} else if (worker.get(10, 63, 71) === dok.dokid) {
					resno = 2;
					$.bibduck.log("tab twice");
					client.send('\t\t\t');
					kommentar = worker.get(16, 13, 79);
				} else if (worker.get(17, 63, 71) === dok.dokid) {
					resno = 3;
					$.bibduck.log("tab thrice");
					client.send('\t\t\t\t\t');
					kommentar = worker.get(23, 13, 79);
				}
					$.bibduck.log('kommentar: "' + kommentar + '"');
				if (kommentar === '') {
					for (var s in config.sigs) {
						if (config.sigs[s] === lib.ltid) {
							sig = s;
						}
					}
					client.send('Sendt ' + sig + ' ' + $.bibduck.stikksedler.current_date());
				}
				setTimeout(function() {
					emitComplete();
					seddel.res(dok, laaner, lib);
				}, 100);
			});

		} else if (dok.utlstatus === 'AVH') {

			if (laaner.beststed == seddel.beststed) {
			
				dok.utlstatus = 'AVH';

				client.send('hentb,\n');
				client.wait_for('Hentebrev til låntaker:', [7,15], function() {
					client.send(dok.dokid + '\n');

					client.wait_for([

						['Kryss av for ønsket valg', [16,8], function() {
							send_hentb_steg2();
						}],

						['Ugyldig LTID fra dato', [9,2], function() {
							//var dt = bibsys.get(9,25,34);
							//$.bibduck.log('NB! Ugyldig LTID fra dato: ' + dt, 'WARN');
							client.send('J\n');
							client.wait_for('Kryss av for ønsket valg', [16,8], function() {
								send_hentb_steg2();
							});
						}]

					]);

				});

				//waitStrs = Array("Hentebeskjed til","*** STOPPMELDING ***","KOMMENTAR:","Ugyldig LTID fra dato",
				// "sistegangspurringer","i utestående gebyr")
				//seddel.reg(dok, laaner, lib);
				//emitComplete();
			} else {

				dok.utlstatus = 'RES';
			
				client.alert('Obs! Låner har bestillingssted ' + laaner.beststed + 
					', så dokumentet må sendes. Du skal få en stikkseddel.');

				setTimeout(function() {
					emitComplete();
					seddel.res(dok, laaner, lib);
				}, 100);
			
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
					client.wait_for('DOkstat', [2,31], function() {
						// FINITO, emit
						seddel.reg(dok, laaner, lib);
						emitComplete();
					});
				//}
			}
		}
	}
	
	function send_hentb_steg2() {
		client.send('X\n');
		client.wait_for([
			['Hentebeskjed er sendt', [1,1], function() {
				$.bibduck.log('Hentebeskjed sendt per sms', 'info');
				client.resetPointer();
				hentebeskjed_sendt();
			}],
			['Registrer eventuell melding', [8,5], function() {
				$.bibduck.sendSpecialKey('F9');
				$.bibduck.log('Hentebeskjed sendt per epost', 'info');
				hentebeskjed_sendt();
			}]
		]);
	}
	
	function hentebeskjed_sendt() {
		var firstline = client.get(1);
		var tilhvem = firstline.match(/på (sms|Email) til (.+) merket (.+)/);
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
				worker.wait_for('Sist aktiv dato', [22,1], les_ltst_skjerm);
			});
		}
	}

	function start_from_imo() {
		laaner = { kind: 'person' };
		lib = {};
		dok = { utlstatus: 'AVH', bestnr:  siste_bestilling.bestnr };
		var firstline = client.get(1);
		var tilhvem = firstline.match(/på (sms|Email) til (.+) merket (.+)/);
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
			les_ltst_skjerm();
			
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
		if (lib.ltid === 'lib'+hjemmebibliotek) {
			client.alert('Boka hører til her. Returseddel trengs ikke.');
			setWorking(false);
			return;
		}

		seddel.ret(dok, laaner, lib);
		emitComplete();

	}

	function checkFormatter(fortsett) {

		// Last inn enhetsspesifikt script
		if (hjemmebibliotek !== $.bibduck.config.libnr) {
			hjemmebibliotek = $.bibduck.config.libnr;
			var f = config.formatters['lib' + hjemmebibliotek];
			$.bibduck.log('Load: plugins/stikksedler/' + f);
			$.getScript('plugins/stikksedler/' + f)
				.done(fortsett)
				.fail(function() {
				$.bibduck.log('Load failed!', 'error');
				setWorking(false);
			});
		} else {
			fortsett();
		}
	}

	function start(info) {

		try {
			$('audio#ping').get(0).play();
		} catch (e) {
			// IE8?
		}
		$.bibduck.log('Skriver ut stikkseddel', 'info');
		seddel = $.bibduck.stikksedler;
		seddel.libnr = 'lib' + $.bibduck.config.libnr;
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
				fso = new ActiveXObject("Scripting.FileSystemObject"),
				shell = new ActiveXObject("WScript.Shell"),
				appdata = shell.ExpandEnvironmentStrings("%APPDATA%"),
				path = appdata + '\\Scriptotek\\Bibduck\\stikk.txt';

			var check = function() {
				if (fso.FileExists(path)){

					var bibsys = $.bibduck.getFocused(),
						txt = readFile(path);

					$.bibduck.log(txt);
					var request = $.parseJSON(txt);
					fso.DeleteFile(path);
					$.bibduck.log('Fikk forespørsel om stikkseddel fra vindu ' + request.window +
						'. Ltid: ' + request.ltid + ', dokid: ' + request.dokid, 'info');
			
					for (var i = 0; i < $.bibduck.instances.length; i++) {
						$.bibduck.log($.bibduck.instances[i].bibsys.index);
						if ($.bibduck.instances[i].bibsys.index === request.window) {
							bibsys = $.bibduck.instances[i].bibsys;
						}
					}
					
					that.forbered_stikkseddel(bibsys, function() {
						$.bibduck.log('forbered_stikkseddel callback');
						bibsys.unidle();
						bibsys.update();   // force update
						start(request);
					});
				}
				setTimeout(check, 1000);
			};
			setTimeout(check, 1000);
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
			
			$.bibduck.log('forbered_stikkseddel');
		
			//$.bibduck.log(current_date);
			if ($.bibduck.config.printerPort === '') {
				client.alert('Sett opp stikkseddelskriver ved å trykke på knappen «Innstillinger» først.');
				setWorking(false);
				return;
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
					checkFormatter(startfn);
				});
			} else {
				checkFormatter(startfn);
			}

		},

		waiting: false,

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

			var trigger1 = ($.bibduck.config.autoStikkEtterRes === true &&
							bibsys.get(1).indexOf('Hentebeskjed er sendt') !== -1 &&
							(bibsys.get(2, 1, 17) === 'Reserveringsliste' || bibsys.get(2, 1, 15) === 'Reservere (RES)')),
				//trigger2 = (bibsys.get(1).indexOf('er returnert') !== -1 && bibsys.get(2).indexOf('IRETur') !== -1),
				trigger3 = (bibsys.getCurrentLine('lower').indexOf('stikk!') !== -1),
				trigger4 = (bibsys.get(1,1,14) === 'Lån registrert' &&
					($.bibduck.config.autoStikkEtterReg === 'autostikk_reg_alle' ||
					$.bibduck.config.autoStikkEtterReg === 'autostikk_reg_lib' && bibsys.get(1, 20, 22) == 'lib')
					),
				trigger5 = ($.bibduck.config.autoStikkEtterRes === true &&
							bibsys.get(1).indexOf('Hentebeskjed er sendt') !== -1 &&
							(bibsys.get(2, 1, 12) === 'Motta innlån'));


			if (this.waiting === false && (trigger1 || trigger3 || trigger4 || trigger5)) {
				this.waiting = true;
				var that = this;
				setTimeout(function(){
					if (trigger3) bibsys.clearInput();
					if (!trigger3) $.bibduck.log('Lager stikkseddel automatisk (stikksedler.js)', 'info');
					that.lag_stikkseddel(bibsys);
				}, 250); // add a small delay
			} else if (this.waiting === true && !trigger1 && !trigger3 && !trigger4 && !trigger5) {
				this.waiting = false;
			}
		}

	});

})();
