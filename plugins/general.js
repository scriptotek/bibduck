/*****************************************************************************
 * Nye kommandoer:
 *   !!     : Tømmer linjen
 *   #!     : Flytter pekeren til kommandolinja (linje 3)
 *****************************************************************************/
window.bibduck.plugins.push({

    name: 'Linjetømmer',

    keypress: function(bibsys, event) {

        var trace = bibsys.getTrace();
        //window.bibduck.log(trace);

        if (trace.length >= 2 && trace.substr(trace.length-2, trace.length) === "!!") {
            bibsys.clearInput();
        }

        if (trace.length >= 2 && trace.substr(trace.length-2, trace.length) === "#!") {
            window.bibduck.log('RESET POINTER');
            bibsys.clearInput();
            bibsys.resetPointer();
        }
    }

});


/*****************************************************************************
 * Tillegg som aksepterer dokid i forfatter-feltet på BIB-skjermen, 
 * slik at man slipper å tabbe ned til dokid-feltet. 
 *****************************************************************************/
window.bibduck.plugins.push({
    name: 'Dokid i forfatter-feltet på BIB-skjermen',

    update: function (bibduck, bibsys) {

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
