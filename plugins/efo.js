/*****************************************************************************
 * <efo.js>
 * Modul for å fornye lån med purrestatus E i én operasjon.
 * Av: Hodo Elmi Aden og Dan Michael O. Heggø (c) 2013
 *
 * Nye kommandoer:
 *   efo! : fornyer lån med status E
 *****************************************************************************/
$.bibduck.plugins.push({

    name: 'EFO',

    keypress: function(bibsys) {

        if (bibsys.getTrace() === "efo!") {

            // Gå til ENDRE-skjermen:
            bibsys.clearInput();
            bibsys.resetPointer();
            bibsys.send('endre,\n');
            bibsys.wait_for('Ptyp', [5,1], function() {

                // Endre purrestatus til N og lagre med F9:
                bibsys.send('N');
                $.bibduck.sendSpecialKey('F9');
                bibsys.wait_for('Utlreferanse', [20,1], function() {

                    // Forny og sett purrestatus tilbake til E:
                    bibsys.resetPointer();
                    bibsys.send('forny,\n');
                    bibsys.wait_for('Ptyp', [6,6], function() {
                        bibsys.send('\tE\n');
                        bibsys.wait_for([

                            // Utfall 1: Lånet blir fornyet direkte
                            ['Fornyet', [1,1], function() {
                                $.bibduck.log('Ok, lånet er fornyet');
                            }],

                            // Utfall 2: Fornyelse må bekreftes med J/N
                            ['LÅNET', [16,16], function() {

                                // Bekreft fornying:
                                bibsys.send('J\n');
                                bibsys.wait_for('Fornyet', [1,1], function() {
                                    $.bibduck.log('Ok, lånet er fornyet');
                                });
                            }]
                        ]);
                    });
                });
            });
        }
    }

});
