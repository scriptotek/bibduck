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
	
	ltsokWorking: false,
	last_dokid: '',

    update: function (bibsys) {
        var dokid,
            cursorpos,
			that = this;

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
        if (!that.ltsokWorking && bibsys.get(2, 1, 25) === 'Søking etter låntakerdata' && bibsys.get(7, 1, 4) === 'Ltid') {

            // Finnes det noe som ligner på et LTID på linje 13?
            ltid = bibsys.get(14, 20, 29).trim();
			
            if (ltid.length === 10 && /^\d+$/.test(ltid.substr(3))) {
				
				that.ltsokWorking = true;
				
				setTimeout(function() {

					// Sjekk hvilken linje vi er på. Hvis dokid er limt inn, 
					// kan det komme med en tab eller enter, slik at vi har hoppet
					// til neste linje før denne rutinen får kjørt
					cursorpos = bibsys.getCursorPos();
					if (cursorpos.row !== 14) {
						while (bibsys.getCursorPos().row !== 14) {
							var crow = bibsys.getCursorPos().row,
								ccol = bibsys.getCursorPos().col;
							bibsys.send('\t');
							do {
								//logger('sleep');
								bibsys.microsleep(); // Venter til pekeren faktisk har flyttet seg
							} while (crow == bibsys.getCursorPos().row && ccol == bibsys.getCursorPos().col);
						}
					}
					bibsys.clearLine();
					bibsys.send('\t\t\t\t\t' + ltid + '\n');
					that.ltsokWorking = false;

				}, 100);

            }
        }
		
		// Er vi på LTST-skjermen?
        if (!that.ltsokWorking && bibsys.get(2, 1, 28) === 'Søk på låntaker fra LTSTatus' && bibsys.get(7, 1, 4) === 'Ltid') {

            // Finnes det noe som ligner på et LTID på linje 13?
            ltid = bibsys.get(12, 20, 29).trim();
			
            if (ltid.length === 10 && /^\d+$/.test(ltid.substr(3))) {
				
				that.ltsokWorking = true;
				
				setTimeout(function() {

					// Sjekk hvilken linje vi er på. Hvis dokid er limt inn, 
					// kan det komme med en tab eller enter, slik at vi har hoppet
					// til neste linje før denne rutinen får kjørt
					cursorpos = bibsys.getCursorPos();
					if (cursorpos.row !== 12) {
						while (bibsys.getCursorPos().row !== 12) {
							var crow = bibsys.getCursorPos().row,
								ccol = bibsys.getCursorPos().col;
							bibsys.send('\t');
							do {
								//logger('sleep');
								bibsys.microsleep(); // Venter til pekeren faktisk har flyttet seg
							} while (crow == bibsys.getCursorPos().row && ccol == bibsys.getCursorPos().col);
						}
					}
					bibsys.clearLine();
					bibsys.send('\t\t\t\t\t' + ltid + '\n');
					that.ltsokWorking = false;

				}, 100);

            }
        }
		
		/**
		 * Sjekk for Onkel Toms hylle 
		 */

		 // Er vi på RET-skjermen?
		if (bibsys.get(2, 1, 15) === 'Returnere utlån') {
			var dokid = bibsys.get(6,31,39);

			// Har vi et nytt dokument som matcher "Onkel Toms" i signaturen?
			if (dokid != that.last_dokid && /Onkel Toms/.test(bibsys.get(11))) {
				that.last_dokid = dokid;
				bibsys.alert('Denne boka skal lånes ut på umn1000906 og settes tilbake på utstilling i Onkel Toms hylle. Forvirra? Legg boka til Karoline :)');
			}
		} else {
			that.last_dokid = '';
		}

	}
});
