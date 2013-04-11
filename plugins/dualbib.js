/*****************************************************************************
 * DualBib er et tillegg med eksperimentell funksjonalitet for å 
 * utnytte to eller flere samtidig åpne BIBSYS-vinduer bedre.
 * 
 * Nye kommandoer:
 *   n!        : Hopper mellom de åpne vinduene
 *   skr,x!    : Viser resultat fra BIB-skjerm i neste vindu
 *****************************************************************************/
window.bibduck.plugins.push({
    name: 'DualBib',
    paa_bibskjerm: false,
    bibscreen: '',

    getNextInstance: function(bibsys) {
        var bibsys2,
            instances = window.bibduck.instances(),
            instance,
            cidx;
        for (var idx in instances) {
            instance = instances[idx];
            if (instance.bibsys.index === bibsys.index) {
                cidx = idx-0;
                break;
            }
        }
        //window.bibduck.log(instances.slice(cidx+1).length)
        var theothers = instances.slice(cidx+1).concat(instances.slice(0,cidx));
        if (theothers.length === 0) {
            return undefined;
        }
        var next = theothers[0];
        return next.bibsys;
    },

    refill: function(bibsys, cb) {
        var bs = this.bibscreen.split('\n');

        if (bibsys === undefined) {
            window.bibduck.log('Fant ikke et annet BIBSYS-vindu');
            return;
        }
        window.bibduck.log('use:' + bibsys.index);
        bibsys.bringToFront();
        bibsys.resetPointer();
        bibsys.send('bib\n');
        setTimeout(function() {
            bibsys.send(
                bs[4].substr(16).trim() + '\t' +
                bs[5].substr(16).trim() + '\t' +
                bs[6].substr(16).trim() + '\t' +
                bs[7].substr(16).trim() + '\t' +
                bs[8].substr(16).trim() + '\t' +
                bs[9].substr(16).trim() + '\t' +
                bs[10].substr(16).trim() + '\t' +
                bs[11].substr(16).trim() + '\t' +
                bs[12].substr(16).trim() + '\t' +
                bs[13].substr(16).trim() + '\t' +
                bs[14].substr(16).trim() + '\t' +
                bs[15].substr(16).trim() + '\t' +
                bs[16].substr(16).trim() + '\t' +
                bs[17].substr(16).trim() + '\t' +
                bs[18].substr(16).trim() + '\t' +
                bs[19].substr(16).trim() + '\t' +
                bs[20].substr(16).trim() + '\n');
                if (cb !== undefined) {
                    cb();
                }
        }, 500);
    },

    keypress: function (bibsys, evt) {

        var bibsys2;

        // n! for å bla gjennom vinduene
        var match = bibsys.getTrace().match(/^n\!/);
        if (match !== null) {
            bibsys.clearInput();
            bibsys2 = this.getNextInstance(bibsys);
            if (bibsys2 !== undefined) {
                bibsys2.bringToFront();
            }
        }

        // skr,x! for å vise resultat i annet vindu
        match = bibsys.getTrace().match(/^skr,([0-9]+)\!/);
        if (match !== null) {
            bibsys.clearInput();
            bibsys2 = this.getNextInstance(bibsys);
            if (bibsys2 === undefined) {
                window.bibduck.log('Fant ikke et annet BIBSYS-vindu');
                return;
            }

            if (bibsys2.get(2,1,10) !== 'Treffliste') {
                this.refill(bibsys2, function() {
                    bibsys.bringToFront();
                    window.bibduck.log('skr,' + match[1]);
                    bibsys2.send('skr,' + match[1] + '\n');
                    setTimeout(function() {
                        bibsys.bringToFront();
                    },100);
                });
            } else {
                window.bibduck.log('skr,' + match[1]);
                bibsys2.send('skr,' + match[1] + '\n');
                setTimeout(function() {
                    bibsys.bringToFront();
                },100);
            }

        }

        //window.bibduck.log(evt.wParam);
        if (this.paa_bibskjerm && evt.wParam === 13) {
            window.bibduck.log('save screen');
            this.bibscreen = bibsys.get();
            /*
            bibsys2 = this.getNextInstance(bibsys);
            this.refill(bibsys2, function() {
                bibsys.bringToFront();
            });
            */
        }
       /*
        if (bibsys.getTrace() == '!ltid') {
            bibsys.clearInput();
            bibsys.send(this.siste_ltid);
        }
        if (bibsys.getTrace() == '!dokid') {
            bibsys.clearInput();
            bibsys.send(this.siste_dokid);
        }*/
    },

    update: function (bibduck, bibsys) {

        // Er vi på LTST-skjermen?
        if (bibsys.get(2, 1, 26) === 'Bibliografisk søk (BIBsøk)') {
            this.paa_bibskjerm = true;
        } else {
            this.paa_bibskjerm = false;
        }

    }
});
