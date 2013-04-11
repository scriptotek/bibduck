/*****************************************************************************
 * Tillegg for å logge LTID og DOKID fra utlån og retur, 
 * samt LTID fra LTST- og LTSØK-besøk.
 *
 * Funksjonaliteten er ment å hjelpe i situasjoner der man
 * "mister" en bruker eller et dokument mens man holder på
 * å jobbe, ikke for å drive systematisk logging. Loggen tømmes
 * når man avslutter BIBDUCK, eller manuelt ved å trykke på 
 * knappen "Tøm logg".
 *
 * Nye kommandoer:
 *   ltid!     : Setter inn siste LTID
 *   dokid!    : Setter inn siste DOKID
 *****************************************************************************/
window.bibduck.plugins.push({
    siste_ltid: '',
    siste_dokid: '',
    aktiv_ltid: '',
    siste_retur: '',
    utlaansskjerm: false,
    name: 'Logger',

    keypress: function (bibsys, evt) {
        if (bibsys.getTrace() == 'ltid!') {
            bibsys.clearInput();
            bibsys.send(this.siste_ltid);
        }
        if (bibsys.getTrace() == 'dokid!') {
            bibsys.clearInput();
            bibsys.send(this.siste_dokid);
        }
    },

    update: function (bibduck, bibsys) {

        // Er vi på LTST-skjermen?
        if (bibsys.get(2, 1, 34) === 'Oversikt over lån og reserveringer') {

            // Finnes det noe som ligner på et LTID på linje 4,
            // (som vi ikke allerede har sett)?
            var ltid = bibsys.get(4, 15, 24).trim();
            if (ltid.length == 10 && ltid != this.aktiv_ltid) {
                this.aktiv_ltid = ltid;
                this.siste_ltid = ltid;
                bibduck.log('LTST for: ' + ltid);
            }

        // Er vi på LTSØK-skjermen?            
        } else if (bibsys.get(2, 1, 34) === 'Opplysninger om låntaker (LTSØk)') {

            // Finnes det noe som ligner på et LTID på linje 4,
            // (som vi ikke allerede har sett)?
            var ltid = bibsys.get(18, 18, 27).trim();
            if (ltid.length == 10 && ltid != this.aktiv_ltid) {
                this.aktiv_ltid = ltid;
                this.siste_ltid = ltid;
                bibduck.log('LTSØK for: ' + ltid);
            }
        } else {
            this.aktiv_ltid = '';
        }

        // Er vi på en retur-skjerm?
        if (bibsys.get(2, 1, 15) === 'Returnere utlån') {
            var dokid = bibsys.get(6, 31, 39).trim();
            if (dokid.length == 9 && dokid != this.siste_retur) {
                var ltid = bibsys.get(15, 16, 25);
                this.siste_retur = dokid;
                this.siste_dokid = dokid;
                this.siste_ltid = ltid;
                bibduck.log('Retur registrert: ' + dokid + ' fra ' + ltid);
            }
        } else {
            this.siste_retur = '';
        }

        // Er vi på en utlånsskjerm?
        var s = bibsys.get(1, 1,14);
        if (this.utlaansskjerm === false && s === 'Lån registrert') {
            var ltid = bibsys.get(1, 20, 29),
                dokid = bibsys.get(10, 7, 15);
            bibduck.log('Utlån registrert: ' + dokid + ' til ' + ltid);
            this.utlaansskjerm = true;
            this.siste_ltid = ltid;
            this.siste_dokid = dokid;
            /*if (confirm("Stikkseddel?") === true) {
                bibsys.bringToFront();
                bibsys.typetext('stikk!');
            }*/
        } else if (this.utlaansskjerm === true && s !== 'Lån registrert') {
            this.utlaansskjerm = false;
        }

    }
});
