/*****************************************************************************
 * Nye kommandoer:
 *   !!     : Tømmer linjen
 *   #!     : Flytter pekeren til kommandolinja (linje 3)
 *****************************************************************************/
$.bibduck.plugins.push({

    name: 'Linjetømmer',

    keypress: function(bibsys) {

        var trace = bibsys.getTrace();
        //$.bibduck.log(trace);

        if (trace === "hjelp!") {
           bibsys.clearInput();
           bibsys.resetPointer();
		   bibsys.send('bib,\n');
		   bibsys.wait_for('Forfatter', [5,1], function() {
				bibsys.send('Åh nei!\t');
				bibsys.send('Åh nei!\t');
				bibsys.send('kanskje lurt å skru av og på?\t');
				bibsys.send('Stakkars deg!\t');
				bibsys.send('Stakkars deg!\t');
				bibsys.send('Huff, åh huff\t');
				bibsys.send('\t');
				bibsys.send('kommer ikke lappene? er det papir i printern?\t');
				bibsys.send('har du prøvd alt?\t');
				bibsys.send('ja, ja, du får ringe og mase på dan michael da.. \t');
				bibsys.send('902 07 510\t');
				bibsys.send('\t');
			});
		}
		
        if (trace.length >= 2 && trace.substr(trace.length - 2, trace.length) === "!!") {
            bibsys.clearInput();
        }

        if (trace.length >= 2 && trace.substr(trace.length - 2, trace.length) === "#!") {
            bibsys.clearInput();
            bibsys.resetPointer();
        }
		
		
    }

});


/*****************************************************************************
 * Tillegg som aksepterer dokid i forfatter-feltet på BIB-skjermen, 
 * slik at man slipper å tabbe ned til dokid-feltet. 
 *****************************************************************************/
$.bibduck.plugins.push({
    name: 'Dokid i forfatter-feltet på BIB-skjermen',

    update: function (bibsys) {
        var dokid,
            cursorpos;

        // Er vi på BIB-skjermen?
        if (bibsys.get(2, 1, 17) === 'Bibliografisk søk') {

            // Finnes det noe som ligner på et dokid på linje 5?
            dokid = bibsys.get(5, 17, 26).trim();
            if (dokid.length === 9 && /^\d+$/.test(dokid.substr(0, 2))) {

                // Sjekk hvilken linje vi er på. Hvis dokid er limt inn, 
                // kan det komme med en tab eller enter, slik at vi har hoppet
                // til neste linje før denne rutinen får kjørt
                cursorpos = bibsys.getCursorPos();
                if (cursorpos.row === 3) {
                    bibsys.send('\t\t');
                    while (bibsys.getCursorPos().row === 3) {
                        bibsys.microsleep();
                    }
                    bibsys.clearLine();
                    bibsys.send('\t\t\t\t' + dokid + '\n');
                } else if (cursorpos.row === 5) {
                    bibsys.send('\t\t\t\t' + dokid + '\n');
                } else if (cursorpos.row === 6) {
                    bibsys.send('\t\t\t' + dokid + '\n');
                }

            }
        }
		
		// Er vi på LTSØK-skjermen?
        if (bibsys.get(2, 1, 25) === 'Søking etter låntakerdata' && bibsys.get(7, 1, 4) === 'Ltid') {

            // Finnes det noe som ligner på et LTID på linje 13?
            ltid = bibsys.get(13, 20, 29).trim();
            if (ltid.length === 10 && /^\d+$/.test(ltid.substr(3))) {

                // Sjekk hvilken linje vi er på. Hvis dokid er limt inn, 
                // kan det komme med en tab eller enter, slik at vi har hoppet
                // til neste linje før denne rutinen får kjørt
                cursorpos = bibsys.getCursorPos();
                if (cursorpos.row === 14) {
                    bibsys.send('\t\t\t\t\t\t\t\t');
                    while (bibsys.getCursorPos().row === 13) {
                        bibsys.microsleep();
                    }
                    bibsys.clearLine();
                    bibsys.send('\t\t\t\t\t' + ltid + '\n');
                } else if (cursorpos.row === 13) {
                    bibsys.clearLine();
					bibsys.send('\t\t\t\t\t' + ltid);
                }

            }
        }
    }
});
