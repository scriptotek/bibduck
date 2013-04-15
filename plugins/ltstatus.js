/*
 * Forsøk på å skrive ut hele låneoversikten for en bruker i ett dokument.
 * NB! dokst,x gir dokst,x mod 15 på ltst, noe som virker som en bug i BIBSYS.
 * Vi leser derfor ut dokid-ene først og bruker så dokst,<dokid>
 */

$.bibduck.plugins.push({

    template: '\\\\platon\\ub-umn\\felles\\umn-skript\\ltstatus.xls',
    name: 'Ltstatus-utskriftstillegg',
    perpage: 15,
    bib: undefined,

    update: function(bibsys) {
        if (bibsys.get(4, 1, 8) === 'Ltstatus' && bibsys.getCurrentLine().indexOf('print!') !== -1) {
            bibsys.clearLine();
            if (bibsys.get(8, 4, 4) !== '1') {
                alert('Vennligst hopp tilbake til første side før du skriver ut');
                return;
            }
            this.bib = bibsys;
            $.bibduck.log('Skriv ut ltstatus');
            this.start();
        }
    },

    start: function() {
        var bib = this.bib;
        this.data = {
            ltid: bib.get(4, 15, 25),
            ltnavn: bib.get(4, 30, 79),
            antutlaan: bib.get(5, 14, 20),
            page: 1,
            item: 0,
            items: []
        };
        this.getPage();
    },

    getPage: function () {
        var lastentry = 1,
            npages = Math.ceil(this.data.antutlaan / this.perpage),
            that = this,
            bib = this.bib,
            line,
            dokid;

        $.bibduck.log('Side ' + this.data.page + ' av ' + npages);

        for (line = 8; line <= 22; line += 1) {
            dokid = bib.get(line, 24, 32);
            if (dokid.length === 9) {
                this.data.items.push({
                    dokid: dokid,
                    abbrtitle: bib.get(line, 34, 54)
                });
            }
        }

        if (this.data.page < npages) {
            this.data.page += 1;
            lastentry = this.perpage * this.data.page;
            line = 22;
            if (lastentry > this.data.antutlaan) {
                line = 22 - (lastentry - this.data.antutlaan);
                lastentry = this.data.antutlaan;
            }
            bib.send('mer\n');
            bib.wait_for(String(lastentry), [line, 4], function() {
                that.getPage();
            });
        } else {
            $.bibduck.log('complete');
            this.dokst();
        }
    },

    lesDokstSkjerm: function() {
        var tittel,
            tittel1,
            tittel2,
            x = this.bib.get(2, 1, 28);
        if (x !== 'Utlånsstatus for et dokument') {
            $.bibduck.log('Vi er ikke på DOKST-skjermen', 'error');
            $.bibduck.log(x);
            return;
        }

        // Sjekker hvilken linje tittelen står på:
        if (this.bib.get(7, 2, 7) === 'Tittel') {
            // Lån fra egen samling
            tittel = this.bib.get(7, 14, 79).trim();
        } else if (this.bib.get(8, 2, 7) === 'Tittel') {
            // Dokument med ik-nummer
            tittel = this.bib.get(8, 13, 79).trim();
        } else {
            /* Relativt sjelden case? Linje 7-10 er fritekst, og 
             * tittel og forfatter bytter typisk mellom linje 7 og 8.
             * En enkel test, som sikkert vil feile i flere tilfeller
             * er å anta at tittelen er den lengste linjen :) 
             */
            tittel1 = this.bib.get(7, 2, 80).trim();
            tittel2 = this.bib.get(8, 2, 80).trim();
            if (tittel1.length > tittel2.length) {
                tittel = tittel1;
            } else {
                tittel = tittel2;
            }
        }

        this.data.items[this.data.item].tittel = tittel;
        this.data.items[this.data.item].forfvres = this.bib.get(20, 18, 27);
        this.data.items[this.data.item].forfall = this.bib.get(21, 18, 27);
        this.data.item += 1;
        this.dokst();

    },

    dokst: function() {
        var i = this.data.item,
            that = this,
            bib = this.bib,
            line;
        if (i < this.data.items.length) {
            $.bibduck.log('Objekt ' + (i + 1) + ' av ' + this.data.items.length);
            if (!bib.resetPointer()) {
                $.bibduck.log('Klarte ikke å finne kommandolinja', 'error');
                return;
            }
            bib.send('dokst,' + this.data.items[i].dokid + '\n');
            bib.wait_for([

                // Dokid på linje 6, kolonne 31:
                [this.data.items[i].dokid, [6, 31], function() {
                    that.lesDokstSkjerm();
                }],

                // Heftevalg:
                ['Kryss av', [2, 3], function() {
                    var abbrtitle = that.data.items[i].abbrtitle,
                        sendstr = '',
                        fnd = false;
                    for (line = 4; line <= 22; line += 1) {
                        if (bib.get(line, 36, 56) === abbrtitle) {
                            $.bibduck.log('Fant riktig hefte på linje ' + line);
                            sendstr += 'X\n';
                            fnd = true;
                            break;
                        }
                        sendstr += '\t';
                    }
                    if (fnd === false) {
                        $.bibduck.log('Fant ikke et hefte "' + abbrtitle + '"', 'error');
                        return;
                    }
                    bib.send(sendstr);
                    bib.wait_for('Er dette korrekt hefte', 12, function() {
                        bib.send('J\n');
                        bib.wait_for(that.data.items[i].dokid, 5, function() {
                            that.lesDokstSkjerm();
                        });
                    });
                }]
            ]);
        } else {
            this.print();
        }

    },

    print: function() {
        // Printe ut via Excel-ark:
        var excel = new ActiveXObject('Excel.Application'),
            j;
        excel.Visible = true;
        excel.Workbooks.Open(this.template);
        excel.Cells(2, 1).Value = " " + this.data.antutlaan + " utlån for " + this.data.ltnavn + " (" + this.data.ltid + ")";

        for (j = 0; j < this.data.items.length; j += 1) {
            excel.Cells(j + 6, 1).Value = j + 1;
            excel.Cells(j + 6, 2).Value = this.data.items[j].forfall;
            excel.Cells(j + 6, 3).Value = this.data.items[j].tittel;
        }

        // excel.ActiveWorkbook.PrintOut();
        // excel.ActiveWorkbook.Close(0);
    }

});