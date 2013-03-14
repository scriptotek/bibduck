
var triggers = [];

/* logger utlån */
triggers.push({
    active: false,
    check: function (bibduck, bibsys) {
        var s = bibsys.get(1, 1,14);
        if (this.active === false && s === 'Lån registrert') {
            var ltid = bibsys.get(1, 20, 29),
                dokid = bibsys.get(10, 7, 15);
            bibduck.log('Utlån registrert: ' + dokid + ' til ' + ltid);
            this.active = true;
            if (confirm("Stikkseddel?") === true) {
                bibsys.bringToFront();
                bibsys.typetext('!stikk');
            }
        } else if (this.active === true && s !== 'Lån registrert') {
            this.active = false;
        }
    }
});

/* logger returer */
triggers.push({
    siste_retur: '',
    check: function (bibduck, bibsys) {
        if (bibsys.get(2, 1, 15) === 'Returnere utlån') {
            var dokid = bibsys.get(6, 31, 39).trim();
            if (dokid.length == 9 && dokid != this.siste_retur) {
                var ltid = bibsys.get(15, 16, 25);
                this.siste_retur = dokid;
                bibduck.log('Retur registrert: ' + dokid + ' fra ' + ltid);
            }
        } else {
            this.siste_retur = '';
        }
    }
});

/* logger ltst-besøk */
triggers.push({
    siste_ltid: '',
    check: function (bibduck, bibsys) {

        // Er vi på LTST-skjermen?
        if (bibsys.get(2, 1, 34) === 'Oversikt over lån og reserveringer') {

            // Finnes det noe som ligner på et LTID på linje 4,
            // (som vi ikke allerede har sett)?
            var ltid = bibsys.get(4, 15, 24).trim();
            if (ltid.length == 10 && ltid != this.siste_ltid) {
                this.siste_ltid = ltid;
                bibduck.log('LTST for: ' + ltid);
            }
        } else {
            this.siste_ltid = '';
        }
    }
});

/*
 * Aksepterer dokid i forfatter-feltet på BIB-skjermen, slik at man
 * slipper å tabbe ned til dokid-feltet 
 */
triggers.push({
    check: function (bibduck, bibsys) {

        // Er vi på BIB-skjermen?
        if (bibsys.get(2, 1, 17) === 'Bibliografisk søk') {

            // Finnes det noe som ligner på et dokid på linje 5?
            var dokid = bibsys.get(5, 17, 26).trim();
            if (dokid.length == 9 && /^\d+$/.test(dokid.substr(0, 2))) {

                // Sjekk hvilken linje vi er på. Hvis dokid er limt inn, 
                // kan det komme med en tab eller enter, slik at vi har hoppet
                // til neste linje før denne rutinen får kjørt
                var c = bibsys.getCursorPos();
                if (c.row == 5) {
                    bibsys.send('\t\t\t\t' + dokid + '\n');
                } else if (c.row == 6) {
                    bibsys.send('\t\t\t' + dokid + '\n');
                }

            }
        }
    }
});

/*
 * Stikkseddel
 */
triggers.push({

    stikkseddel: undefined,

    check: function (bibduck, bibsys) {

        if (this.stikkseddel === undefined) {
            this.stikkseddel = new Stikkseddel(bibduck, bibsys);
        }

        // Sjekk om en quickbutton har endret vindustittelen 
        // (i mangel av en bedre måte å kommunisere på)
        if (bibsys.getCurrentLine().indexOf('!stikk') !== -1) {
            bibduck.log('Skriv ut stikkseddel');
            bibsys.clearLine();
            this.stikkseddel.start();
        }
    }
});
