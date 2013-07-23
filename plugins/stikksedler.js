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
        excel,
        hjemmebibliotek = '',
        current_date = '',
        config,
        seddel,
        callback,
        working = false;

    function les_dokstat_skjerm() {

        if (client.get(2, 1, 28) !== 'Utlånsstatus for et dokument') {
            client.alert('Vi er ikke på DOKST-skjermen :(');
            working = false;
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
            working = false;
            return;
        }

        laaner.ltid      = client.get(14, 11, 20);
        dok.utlaansdato  = client.get(18, 18, 27);   // Utlånsdato
        dok.forfvres     = client.get(20, 18, 27);   // Forfall v./res
        dok.forfallsdato = client.get(21, 18, 27);   // Forfallsdato
        dok.utlstatus    = client.get( 3, 46, 65);   // AVH, RES, UTL, UTL/RES, ...
        dok.purretype    = client.get(17, 68, 68);
        dok.kommentar    = client.get(23, 17, 80).trim();

        // Dokument til avhenting?
        if (dok.utlstatus === 'AVH') {
            dok.hentenr = client.get(1, 44, 50);
            dok.hentefrist = client.get(1, 26, 35);

        } else {

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
            $.bibduck.log("Går til RLIST for å identifisere låneren");
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

                $.bibduck.sendSpecialKey('F12');
                worker.wait_for('DOkstat', [2,31], function() {
                    worker.resetPointer();

                    // Vi trenger mer info om låneren:
                    $.bibduck.log("Går til LTSØK for å finne ut mer om låneren");
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
        working = false;
        if (callback !== undefined) {
            setTimeout(function() { // a slight delay never hurts
                callback({
                    patron: laaner,
                    library: lib,
                    document: dok,
                    beststed: seddel.beststed
                });
            }, 200);
        }
    }

    function les_ltst_skjerm() {
        if (worker.get(2, 1, 24) !== 'Opplysninger om låntaker') {
            client.alert("Vi er ikke på LTSØ-skjermen :(");
            working = false;
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

        if (worker !== client) {
            worker.resetPointer();
            worker.send('men,\n');
        } else {

            if (dok.utlstatus === 'RES') {

                if (laaner.beststed == seddel.beststed) {
                    client.alert('Obs! Låner har bestillingssted ' + laaner.beststed + ', så det burde ikke være behov for å sende det.');
                    working = false;
                    return;
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
                            emitComplete();
                        });
                    //}
                }
            }
        }

        // Nå har vi informasjonen vi trenger. La oss kjøre i gang Excel-helvetet, joho!!

        // @TODO: Hva med UTL/RES ?
        if (dok.utlstatus === 'RES') {

            var resno,
                kommentar,
                sig = '???';
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

        } else {
            seddel.reg(dok, laaner, lib);
            emitComplete();
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
            working = false;
            return;
        }

        if (client.get(1, 1, 12) === 'Hentebeskjed') {
            dok.utlstatus = 'AVH';
        }

        if (client.get(1, 1, 12) !== 'Hentebeskjed' && client.get(20, 19, 21) !== 'Nr.') {
            $.bibduck.log('Ingen reservering gjennomført, kan ikke skrive ut stikkseddel', 'error');
            client.alert('Du må gjennomføre en reservering før du kan skrive ut stikkseddel');
            working = false;
            return;
        }

        dok.tittel = '';
        if (dok.utlstatus === 'AVH') {

            laaner.ltid = client.get(5, 12, 22);
            dok.dokid = client.get(5, 53, 61);

            // Gå til dokst:
            $.bibduck.sendSpecialKey('F12');
            client.wait_for('DOkstat', [2,31], function() {
                if (client.get(23,1,12) === 'Utlkommentar') {
                    les_dokstat_skjerm();
                } else {
                    client.send(dok.dokid + '\n');
                    client.wait_for('Utlkommentar', [23,1], les_dokstat_skjerm);
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
            worker.send('ltsø,' + laaner.ltid + '\n');
            worker.wait_for('Fyll ut:', [5,1], function() {
                // Vi sender enter på nytt
                worker.send('\n');
                worker.wait_for('Sist aktiv dato', [22,1], les_ltst_skjerm);
            });
        }
    }

    function start_from_rlist() {
        laaner = { kind: 'person' };
        lib = {};
        dok = {};
        var resno = -1;

        if (client.get(2, 1, 25) !== 'Reserveringsliste (RLIST)') {
            $.bibduck.log('Ikke på rlist-skjerm', 'error');
            working = false;
            return;
        }

        var firstline = client.get(1);
        if (firstline.indexOf('Hentebeskjed er sendt') !== -1) {
            var tilhvem = firstline.match(/på sms til (.+) merket/);
            if (client.get(4).match(tilhvem[1])) {
                resno = 1;
            } else if (client.get(11).match(tilhvem[1])) {
                resno = 2;
            } else if (client.get(18).match(tilhvem[1])) {
                resno = 3;
            } else {
                $.bibduck.log('Fant ikke "' + tilhvem[1] + '" på RLIST-skjermen!', 'error');
                client.alert("Beklager, klarte ikke skrive ut stikkseddel automatisk. Prøv manuelt");
                working = false;
                return;
            }
            $.bibduck.log('Hvilken reservasjon på RLIST-skjermen? Bruker nummer ' + resno + 'fordi hentebeskjed ble sendt til "' + tilhvem[1] + '"', 'info');
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
                working = false;
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
            working = false;
            return;
        }
        if (dok.dokid === '') {
            client.alert("Reservering nummer " + resno + " er tom. Flytt pekeren til Ref.-feltet for reserveringen du ønsker å skrive ut stikkseddel for og prøv på nytt.");
            working = false;
            return;
        }

        dok.tittel = '';

            // Gå til dokst:
            $.bibduck.log('Sender F12');
            $.bibduck.sendSpecialKey('F12');
            client.wait_for('DOkstat', [2,31], function() {
                if (client.get(6,31,39) === dok.dokid) {
                    les_dokstat_skjerm();
                } else {
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

    function retur() {
        worker.resetPointer();

        laaner = {};
        lib = {};
        dok = {};

        if (client.get(2, 27, 31) === 'LTSØk') {

            // Vi skriver ut en retur-seddel. Nyttig f.eks. hvis 
            // man ikke får stikkseddel fra IRET

            if (client.get(18, 18, 20) !== 'lib') {
                client.alert("Feil: Låntakeren er ikke et bibliotek!");
                $.bibduck.log("Feil: Låntakeren er ikke et bibliotek!", 'warn');
                working = false;
                return;
            }

            laaner.ltid = client.get(18, 18, 27);
            laaner.navn = client.get(10, 18, 50);
            laaner.kind = 'bibliotek';
            lib.ltid = laaner.ltid;
            lib.navn = laaner.navn;

            seddel.ret(dok, laaner, lib);
            emitComplete();
            return;

        } else if (client.get(2).indexOf('IRETur') !== -1) {

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
                working = false;
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
            working = false;
            return;
        }

        seddel.ret(dok, laaner, lib);
        emitComplete();

    }

    function checkFormatter() {

        // Last inn enhetsspesifikt script
        if (hjemmebibliotek !== $.bibduck.libnr) {
            hjemmebibliotek = $.bibduck.libnr;
            var f = config.formatters['lib' + hjemmebibliotek];
            $.bibduck.log('Load: plugins/stikksedler/' + f);
            $.getScript('plugins/stikksedler/' + f)
             .done(start)
             .fail(function(jqxhr, settings, exception) {
                $.bibduck.log('Load failed!', 'error');
                working = false;
             });
        } else {
            start();
        }
    }

    function start() {

        $.bibduck.log('Skriver ut stikkseddel', 'info');
        seddel = $.bibduck.stikksedler;
        seddel.libnr = 'lib' + $.bibduck.libnr;
        seddel.beststed = '';
        for (var key in config.bestillingssteder) {
            if (config.bestillingssteder[key] == seddel.libnr) {
                seddel.beststed = key;
            }
        }
        if (seddel.libnr === 'lib') {
            client.alert('Obs! Libnr. er ikke satt enda. Dette setter du under Innstillinger i Bibduck.');
            working = false;
            return;
        } else if (seddel.beststed === '') {
            client.alert('Fant ikke et bestillingssted for biblioteksnummeret ' + seddel.libnr + ' i config.json!');
            working = false;
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
            retur();
        } else if (client.get(2, 1, 15) === 'Reservere (RES)') {
            start_from_res();
        } else if (client.get(2, 1, 25) === 'Reserveringsliste (RLIST)') {
            start_from_rlist();
        } else {
            working = false;
            $.bibduck.log('Stikkseddel fra denne skjermen er ikke støttet', 'warn');
            client.alert('Stikkseddel fra denne skjermen er ikke støttet (enda). Ta DOKST og prøv igjen');
        }
    }

    $.bibduck.plugins.push({

        name: 'Stikkseddel-tillegg',

        initialize: function() {
            var that = this;
            $('#btn-stikkseddel').remove();
            var $btn = $('<button class="btn" id="btn-stikkseddel">Stikkseddel</button>');
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
                working = false;
            });
            working = true;
            callback = cb;
            client = bibsys;
            current_date = client.get(3, 70, 79);
            //$.bibduck.log(current_date);
            if ($.bibduck.printerPort === '') {
                client.alert('Sett opp stikkseddelskriver ved å trykke på knappen «Innstillinger» først.');
                working = false;
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

        waiting: false,

        update: function(bibsys) {

            var trigger1 = (bibsys.get(1).indexOf('Hentebeskjed er sendt') !== -1 && (bibsys.get(2, 1, 17) === 'Reserveringsliste' || bibsys.get(2, 1, 15) === 'Reservere (RES)')),
                trigger2 = (bibsys.get(1).indexOf('er returnert') !== -1 && bibsys.get(2).indexOf('IRETur') !== -1),
                trigger3 = (bibsys.getCurrentLine('lower').indexOf('stikk!') !== -1);

            if (this.waiting === false && (trigger1 || trigger2 || trigger3)) {
                this.waiting = true;
                var that = this;
                setTimeout(function(){
                    if (trigger3) bibsys.clearInput();
                    if (!trigger3) $.bibduck.log('stikksedler.js: Lager stikkseddel automatisk', 'info');
                    that.lag_stikkseddel(bibsys);
                }, 250); // add a small delay
            } else if (this.waiting === true && !trigger1 && !trigger2 && !trigger3) {
                this.waiting = false;
            }

        }

    });

})();