

function Stikkseddel (bibduck, bibsys) {

    var config = {
            beststed: 'ureal',
            stikkseddelfil: '\\\\platon\\ub-umn\\felles\\umn-skript\\stikkseddel-umn.xls'
        },
        worker = null;

    var sigs = {
        'UHS'        : 'lib1030300',
        'UHS/SOPH'   : 'lib1030303',
        'UHS/ETNO'   : 'lib1030010',
        'UHS/ARK'    : 'lib1030011',
        'UJUR'       : 'lib1030300',
        'UJUR/IFP'   : 'lib1030001',
        'UJUR/IKR'   : 'lib1030002',
        'UJUR/IOR'   : 'lib1030003',
        'UJUR/IRI'   : 'lib1030004',
        'UJUR/NIP'   : 'lib1030005',
        'UJUR/NIS'   : 'lib1030006',
        'UJUR/Nifs'  : 'lib1030006',
        'UJUR/NifsP' : 'lib1030006',
        'UJUR/DN'    : 'lib1030009',
        'UJUR/RS'    : 'lib1030015',
        'UJUR/MR'    : 'lib1030048',
        'UMN/NHM'    : 'lib1030500',
        'UREAL/NHM'  : 'lib1030500',
        'UMN/INF'    : 'lib1030317',
        'UREAL/INF'  : 'lib1030317',
        'UMED'       : 'lib1032300',
        'UMED/ODONT' : 'lib1030307',
        'UREAL'      : 'lib1030310'
    };
    var biblnavn = {
        'lib1030300' : 'HumSam-biblioteket',
        'lib1030303' : 'Biblioteket Sophus Bugge',
        'lib1030010' : 'HumSam-biblioteket. Etnografi',
        'lib1030011' : 'HumSam-biblioteket. Arkeologi.',
        'lib1030000' : 'Juridisk bibliotek',
        'lib1030001' : 'Juridisk bibliotek. Privatrett',
        'lib1030002' : 'Juridisk bibliotek. Kriminologi og rettssosiologi',
        'lib1030003' : 'Juridisk bibliotek. Offentlig rett.',
        'lib1030004' : 'Juridisk bibliotek. Rettsinformasjon.',
        'lib1030005' : 'Juridisk bibliotek. Petroleumsrett og europarett',
        'lib1030006' : 'Juridisk bibliotek. Sjørett',
        'lib1030009' : 'Juridisk bibliotek. Læringssenteret',
        'lib1030015' : 'Juridisk bibliotek. Rettshistorisk samling',
        'lib1030048' : 'Juridisk bibliotek. Menneskerettigheter',
        'lib1030500' : 'Realfagsbiblioteket. Naturhistorisk Museum',
        'lib1030317' : 'Realfagsbiblioteket. Informatikk',
        'lib1032300' : 'Medisinsk bibliotek',
        'lib1030307' : 'Medisinsk bibliotek. Odontologi',
        'lib1030310' : 'Realfagsbiblioteket'
    };
    var bestillingssteder = {
        'umh'        : 'lib1032300',
        'umhpsyk'    : 'lib1032300', // Medisinsk, siden umhpsyk er nedlagt, right?
        'uod'        : 'lib1030307',
        'uhs'        : 'lib1030300',
        'uhssoph'    : 'lib1030303',
        'uhsetno'    : 'lib1030010',
        'uhsark'     : 'lib1030011',
        'ujur'       : 'lib1030000',
        'ujurifp'    : 'lib1030001',
        'ujurikr'    : 'lib1030002',
        'ujurior'    : 'lib1030003',
        'ujuriri'    : 'lib1030004',
        'ujurnip'    : 'lib1030005',
        'ujurnif'    : 'lib1030006',
        'ujurdn'     : 'lib1030009',
        'ujurrs'     : 'lib1030015',
        'ujurmr'     : 'lib1030048',
        'umninf'     : 'lib1030317',
        'umnnhm'     : 'lib1030500',
        'ureal'      : 'lib1030310'
    };
    var dok = {};
    var laaner = {};
    var lib = {};

    //Printe ut via formatert Excel-ark:
    var excel = new ActiveXObject('Excel.Application');
    excel.Visible = true;
    excel.Workbooks.Open(config.stikkseddelfil);

    function timeout() {
        alert('Bibsys svarer ikke. Ta DOKSTAT på dokumentet, og prøv igjen.');
    }

    function les_dokstat_skjerm(bibsys) {

        if (bibsys.get(2, 1, 28) !== 'Utlånsstatus for et dokument') {
            alert("Vi er ikke på DOKST-skjermen :(");
            return;
        }

        // Sjekker hvilken linje tittelen står på:
        if (bibsys.get(7, 2, 7) == 'Tittel') {
                // Lån fra egen samling
            dok.tittel = bibsys.Get(7, 14, 80).trim();
        } else if (bibsys.get(8, 2, 7) == 'Tittel') {
            // ik...
            dok.tittel = bibsys.get(8, 13, 80).trim();
        } else {
            // Relativt sjelden case? Linje 7-10 er fritekst, og 
            // tittel og forfatter bytter typisk mellom linje 7 og 8.
            // En enkel test, som sikkert vil feile i flere tilfeller:
            var tittel1 = bibsys.get(7, 2, 80).trim(),
                tittel2 = bibsys.get(8, 2, 80).trim();
            if (tittel1.length > tittel2.length) {
                dok.tittel = tittel1;
            } else {
                dok.tittel = tittel2;
            }
        }

        dok.dokid        = bibsys.get( 6, 31, 39);
        laaner.ltid      = bibsys.get(14, 11, 20);
        dok.utlaansdato  = bibsys.get(18, 18, 27);   // Utlånsdato
        dok.forfvres     = bibsys.get(20, 18, 27);   // Forfall v./res
        dok.forfallsdato = bibsys.get(21, 18, 27);   // Forfallsdato
        dok.utlstatus    = bibsys.get( 3, 46, 65);
        dok.purretype    = bibsys.get(17, 68, 68);
        dok.kommentar    = bibsys.get(23, 17, 80).trim();

        if (dok.dokid === '') {
            alert('Har du husket å trykke enter?');
            return;
        }

        //Tester om låntaker er et bibliotek:
        if (laaner.ltid.substr(0,3) == 'lib') {
            laaner.kind = 'bibliotek';
            laaner.navn = bibsys.get(10, 18, 28).trim();
        } else {
            laaner.kind = 'person';
        }

        // DEBUG:
        bibduck.log('Info om lånet:');
        $.each(dok, function(k,v) {
            bibduck.log('  ' + k + ': ' + v);
        });

        // 1. Vi sender ltsø,<ltid><enter>
        worker.resetPointer();
        worker.send('ltsø,' + laaner.ltid + '\n');
        worker.wait_for('Fyll ut:', {
            success: function () {
                // Vi sender enter på nytt
                worker.send('\n');
                worker.wait_for('LTSØk', {
                    success: les_ltst_skjerm,
                    failure: timeout
                });
            },
            failure: timeout
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
        bibduck.log('Info om låner:');
        $.each(laaner, function(k,v) {
            bibduck.log('  ' + k + ': ' + v);
        });

        if (laaner.beststed in bestillingssteder) {
            lib.ltid = bestillingssteder[laaner.beststed];
        } else {
            alert("Ukjent bestillingssted: " + laaner.beststed);
            return;
        }
        if (lib.ltid in biblnavn) {
            lib.navn = biblnavn[lib.ltid];
        } else {
            alert("Ukjent bibliotek: " + lib.ltid);
            return;
        }

        // DEBUG:
        bibduck.log('Info om bibliotek:');
        $.each(lib, function(k,v) {
            bibduck.log('  ' + k + ': ' + v);
        });

        if (worker !== bibsys) {
            worker.resetPointer();
            worker.send('men,\n');
        } else {

            // Gi beskjed hvis boka skal ut av huset
            if (laaner.kind === 'person' && laaner.beststed !== config.beststed) {
                alert('Obs! Låner har bestillingssted: ' + laaner.beststed);

                // Hvis boken skal sendes, så gå til utlånskommentarfeltet.
                worker.Send("en," + dok.dokid + '\n');
                worker.wait_for('Utlmkomm:', {
                    success: function () {
                        worker.send('\t\t\t');
                    },
                    failure: timeout
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
                    bibsys.send('dokst,' + dok.dokid + '\n');
                //}

            }
        }

        // Nå har vi informasjonen vi trenger. La oss kjøre i gang Excel-helvetet, joho!!
        lag_stikkseddel();
    }


    function lag_stikkseddel() {

        if (bibduck.printerPort === '') {
            alert('Sett opp stikkseddelskriver først.');
            return;
        }

        var printerStr = bibduck.printerName + ' on ' + bibduck.printerPort;
        bibduck.log('Printer: "' + printerStr + '"');

        excel.Application.ActivePrinter = printerStr;

        if (laaner.kind === 'person') {
            excel.Cells(1, 1).Value = laaner.etternavn + ', ' + laaner.fornavn;
        }

        /*utlaansdato = formatDate(utlaansdato, ltspraak)
        forfvres = formatDate(forfvres, ltspraak)
        forfallsdato = formatDate(forfallsdato, ltspraak)
*/
        if (laaner.spraak === 'ENG') {
            excel.Cells(2, 1).Value = " Loan date : " + dok.utlaansdato
            if (dok.forfvres !== dok.forfallsdato) {
                excel.Cells(3, 2).Value = "Due : " + dok.forfallsdato 
                excel.Cells(4, 2).Value = "If required by another loaner, the document may be recalled from: " + dok.forfvres
            } else {
                excel.Cells(3, 2).Value = "Due : " + dok.forfvres 
            }
            excel.Cells(7, 1).Value = "Title:"
        } else {
            excel.Cells(2, 1).Value = " Utlånsdato : " + dok.utlaansdato
            if (dok.forfvres !== dok.forfallsdato) {
                excel.Cells(3, 2).Value = "Lånefrist / Due : " + dok.forfallsdato 
                excel.Cells(4, 2).Value = "Ved reservasjoner kan dokumentet bli innkalt fra: " + dok.forfvres
            } else {
                excel.Cells(3, 2).Value = "Lånefrist / Due : " + dok.forfvres 
            }
            excel.Cells(7, 1).Value = "Tittel:"
        }

        excel.Cells(7, 3).Value = dok.tittel
        excel.Cells(8, 3).Value = dok.dokid

        // Hvis ikke fjernlån, skriv ut litt ekstra info om fornying:
        if (laaner.kind === 'person') {
            if (dok.purretype === "E") {
                if (dok.utlstatus === "UTL/RES") {
                    if (laaner.spraak === "ENG") {
                        excel.Cells(11, 1).Value = "Please note:"
                        excel.Cells(12, 1).Value = "This document can not be renewed as it has been reserved by someone else."
                    } else {
                        excel.Cells(11, 1).Value = "NB:"
                        excel.Cells(12, 1).Value = "Dette dokumentet kan ikke fornyes, da det er reservert for en annen låntaker."
                    }
                } else {
                    if (laaner.spraak === "ENG") {
                        excel.Cells(11, 1).Value = "This document can not be renewed online at BIBSYS Ask."
                        excel.Cells(12, 1).Value = "Please visit the library if you want to renew it."
                    } else {
                        excel.Cells(11, 1).Value = "Dette lånet kan du ikke fornye selv på BIBSYS Ask."
                        excel.Cells(12, 1).Value = "Kom innom biblioteket hvis du ønsker å fornye dette lånet."
                    }
                }
            } else {
                if (laaner.spraak === "ENG") {
                    excel.Cells(11, 1).Value = "Unless requested by someone else, this document can be"
                    excel.Cells(12, 1).Value = "renewed online at BIBSYS Ask."
                } else {
                    excel.Cells(11, 1).Value = "Dette lånet kan du fornye selv på BIBSYS Ask"
                    excel.Cells(12, 1).Value = "hvis det ikke kommer reserveringer."
                }
            }
        }

        //Default skjule engelsk logo
        excel.ActiveSheet.Shapes("Picture 2").Visible = false;

        if (laaner.kind === 'bibliotek') {
            excel.Cells(1, 1).Value  = "Fjernlån til " + laaner.navn;
            excel.Cells(32, 1).Value = laaner.ltid.substr(4,8); //Venstre del av lib-nr.
            excel.Cells(32, 4).Value = laaner.ltid.substr(8); //Høyre del av lib-nr.

        } else if (laaner.beststed !== config.beststed) {
            excel.Cells(32, 1).Value = laaner.ltid.substr(4,8); //Venstre del av lib-nr.
            excel.Cells(32, 4).Value = laaner.ltid.substr(8); //Høyre del av lib-nr.
            excel.Cells(31, 1).Value = biblnavn[laaner.ltid];
        }

        excel.ActiveWorkbook.PrintOut();
        //excel.ActiveWorkbook.Close(0);
        //excel.Quit();

    }

    this.start = function() {

        if (bibduck.getBackgroundInstance() !== null) {
            worker = bibduck.getBackgroundInstance();
            worker.Synchronous = true;
        } else {
            worker = bibsys;
        }

        if (bibsys.get(2, 1, 22) == 'Registrere utlån (REG)') {
            var dokid = bibsys.get(10, 7, 15);

            // Gå til DOKST-skjerm:
            worker.resetPointer();
            worker.send('dokst\n');
            //Kan ikke ta dokst, (med komma) for da blir dokid automatisk valgt og aldri refid, sender separat
            worker.wait_for('Utlånsstatus for et dokument', {
                success: function () {
                    worker.send(dokid + '\n');
                    worker.wait_for('Utlkommentar', {
                        success: function () {
                            les_dokstat_skjerm(worker);
                        },
                        failure: timeout
                    });
                },
                failure: timeout
            });
        } else if (bibsys.get(2, 1, 28) == 'Utlånsstatus for et dokument') {
            les_dokstat_skjerm(bibsys);
        } else {
            alert('Stikkseddel fra denne skjermen er ikke støttet (enda). Ta DOKST og prøv igjen');
            bibsys.bringToFront();
        }

    };

}
