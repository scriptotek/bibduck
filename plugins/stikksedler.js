/*****************************************************************************
 * Tillegg for å skrive stikksedler
 *
 * Nye kommandoer:
 *   stikk!     : Skriver stikkseddel
 *****************************************************************************/
$.bibduck.stikksedler = {

    // Settes under Innstillinger i brukergrensesnittet
    beststed: '',

    load_xls: function (filename) {
        var printerStr = window.bibduck.printerName + ' on ' + window.bibduck.printerPort;
        this.excel = new ActiveXObject('Excel.Application');
        this.excel.Visible = false;
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
        excel,
        hjemmebibliotek = '',
        current_date = '',
        config,
        seddel,
		callback;

    function les_dokstat_skjerm() {

        if (client.get(2, 1, 28) !== 'Utlånsstatus for et dokument') {
            alert("Vi er ikke på DOKST-skjermen :(");
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

        dok.dokid        = client.get( 6, 31, 39);
        laaner.ltid      = client.get(14, 11, 20);
        dok.utlaansdato  = client.get(18, 18, 27);   // Utlånsdato
        dok.forfvres     = client.get(20, 18, 27);   // Forfall v./res
        dok.forfallsdato = client.get(21, 18, 27);   // Forfallsdato
        dok.utlstatus    = client.get( 3, 46, 65);
        dok.purretype    = client.get(17, 68, 68);
        dok.kommentar    = client.get(23, 17, 80).trim();

        if (dok.dokid === '') {
            alert('Har du husket å trykke enter?');
            return;
        }

        //Tester om låntaker er et bibliotek:
        if (laaner.ltid.substr(0,3) == 'lib') {
            laaner.kind = 'bibliotek';
            laaner.navn = client.get(10, 18, 28).trim();
        } else {
            laaner.kind = 'person';
        }

        // DEBUG:
        /*
        $.bibduck.log('Info om lånet:');
        $.each(dok, function(k,v) {
            $.bibduck.log('  ' + k + ': ' + v);
        });
*/

        // 1. Vi sender ltsø,<ltid><enter>
        worker.resetPointer();
        worker.send('ltsø,' + laaner.ltid + '\n');
        worker.wait_for('Fyll ut:', [5,1], function() {
            // Vi sender enter på nytt
            worker.send('\n');
            worker.wait_for('Sist aktiv dato', [22,1], les_ltst_skjerm);
        });
    }

    function les_ltst_skjerm() {
        if (worker.get(2, 1, 24) !== 'Opplysninger om låntaker') {
            alert("Vi er ikke på LTSØ-skjermen :(");
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

        if (laaner.beststed in config.bestillingssteder) {
            lib.ltid = config.bestillingssteder[laaner.beststed];
        } else {
            alert("Ukjent bestillingssted: " + laaner.beststed);
            return;
        }
        if (lib.ltid in config.biblnavn) {
            lib.navn = config.biblnavn[lib.ltid];
        } else {
            alert("Ukjent bibliotek: " + lib.ltid);
            return;
        }

        // DEBUG:
        /*
        $.bibduck.log('Info om bibliotek:');
        $.each(lib, function(k,v) {
            $.bibduck.log('  ' + k + ': ' + v);
        });*/

        if (worker !== client) {
            worker.resetPointer();
            worker.send('men,\n');
        } else {

            // Gi beskjed hvis boka skal ut av huset
            if (laaner.kind === 'person' && laaner.beststed !== seddel.beststed) {
                alert('Obs! Låner har bestillingssted: ' + laaner.beststed);

                // Hvis boken skal sendes, så gå til utlånskommentarfeltet.
                client.send('en,' + dok.dokid + '\n');
                client.wait_for('Utlmkomm:', [8,1], function() {
                    client.send('\t\t\t');
					setTimeout(function() {
						// FINITO, emit
						if (callback !== undefined) {
							callback({
								patron: laaner,
								library: lib,
								document: dok,
								beststed: seddel.beststed
							});
						}
					}, 200);
                });

            // Hvis ikke går vi tilbake til dokst-skjermen:
            } else {

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
						if (callback !== undefined) {
							callback({
								patron: laaner,
								library: lib,
								document: dok,
								beststed: seddel.beststed
							});
						}
					});
                //}

            }
        }

        // Nå har vi informasjonen vi trenger. La oss kjøre i gang Excel-helvetet, joho!!
        seddel.reg(dok, laaner, lib);
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

    function retur() {
        worker.resetPointer();

        laaner = {};
        lib = {};

        if (worker.get(2, 16, 21) === 'IRETur') {
            dok = {
                dokid: client.get(1, 2, 10),
                bestnr: client.get(4, 48, 60)
            };
            laaner.ltid = worker.get(6, 16, 25);
            laaner.navn = worker.get(7, 21, 50);
            if (worker.get(9, 3, 8) === 'Tittel') {
                dok.tittel = worker.get(9, 14, 80);
            } else if (worker.get(10, 3, 8) === 'Tittel') {
                dok.tittel = worker.get(10, 14, 80);
            }
            if (laaner.navn === 'xxx') {
                lib.ltid = '';
                lib.navn = '';
            }

        } else {

            // Retur til annet bibliotek innad i organisasjonen
            var sig = client.get(11, 14, 40).split(' ')[0];
            dok = {
                dokid: client.get(6, 31, 39),
                bestnr: ''
            };
            if (sig in config.sigs) {
                lib.ltid = config.sigs[sig];
                lib.navn = config.biblnavn[lib.ltid];
            } else {
                alert('Beklager, BIBDUCK kjenner ikke igjen signaturen "' + sig + '".');
                return;
            }

        }

        if (client.get(7, 2, 7) == 'Tittel') {
            dok.tittel = client.get(7, 14, 79);
        } else if (client.get(8, 2, 7) == 'Tittel') {
            dok.tittel = client.get(8, 14, 79);
        }

        if (hjemmebibliotek === '') {
            alert('Libnr. er ikke satt. Dette setter du under Innstillinger.');
        }
        if (lib.ltid === 'lib'+hjemmebibliotek) {
            alert('Boka hører til her. Returseddel trengs ikke.');
            client.bringToFront();
            return;
        }

        seddel.ret(dok, laaner, lib);

    }

    function checkFormatter() {

        // Last inn enhetsspesifikt script
        if (hjemmebibliotek !== $.bibduck.libnr) {
            hjemmebibliotek = $.bibduck.libnr;
            var f = config.formatters['lib' + hjemmebibliotek];
            $.bibduck.log('Load: plugins/stikksedler/' + f);
            $.getScript('plugins/stikksedler/' + f, function() {
                start();
            });
        } else {
            start();
        }
    }

    function start() {

        seddel = $.bibduck.stikksedler;
        seddel.libnr = 'lib' + $.bibduck.libnr;
        seddel.beststed = '';
        for (var key in config.bestillingssteder) {
            if (config.bestillingssteder[key] == seddel.libnr) {
                seddel.beststed = key;
            }
        }
        if (seddel.libnr === 'lib') {
            alert('Obs! Libnr. er ikke satt enda. Dette setter du under Innstillinger i Bibduck.');
            return;
        } else if (seddel.beststed === '') {
            alert('Fant ikke et bestillingssted for biblioteksnummeret ' + seddel.libnr + ' i config.json!');
            return;
        }

        if (client.get(2, 1, 22) === 'Registrere utlån (REG)') {
            utlaan();
        } else if (client.get(14, 1, 8) === 'Låntaker') {
            utlaan();
        } else if (client.get(15, 2, 13) === 'Returnert av') {
            retur();
        } else if (client.get(2, 16, 21) === 'IRETur') {
            retur();
        } else {
            alert('Stikkseddel fra denne skjermen er ikke støttet (enda). Ta DOKST og prøv igjen');
            client.bringToFront();
        }
    }

    $.bibduck.plugins.push({

        name: 'Stikkseddel-tillegg',
		
		lag_stikkseddel: function(bibsys, cb) {
			callback = cb
            client = bibsys;
		    current_date = client.get(3, 70, 79);
			$.bibduck.log(current_date);
			if ($.bibduck.printerPort === '') {
				alert('Sett opp stikkseddelskriver ved å trykke på knappen «Innstillinger» først.');
				return;
			}

			if ($.bibduck.getBackgroundInstance() !== null) {
				worker = $.bibduck.getBackgroundInstance();
			} else {
				worker = client;
			}

			// Load config if not yet loaded
			if (config === undefined) {
				bibduck.log('Load: plugins/stikksedler/config.json');
				$.getJSON('plugins/stikksedler/config.json', function(json) {
					config = json;
					checkFormatter();
				});
			} else {
				checkFormatter();
			}

		},

        update: function(bibsys) {

            if (bibsys.getCurrentLine().indexOf('stikk!') !== -1) {
                bibsys.clearInput();
				this.lag_stikkseddel(bibsys);
            }
        }

    });

})();