
var macros = [];

/* Makro som logger utlån */
macros.push({
    active: false,
    check: function (bibduck, bibsys) {
        var s = bibsys.get(1, 1,14);
        if (this.active === false && s === 'Lån registrert') {
            var ltid = bibsys.get(1, 20, 29),
                dokid = bibsys.get(10, 7, 15);
            bibduck.log('Utlån registrert: ' + dokid + ' til ' + ltid);
            this.active = true;
        } else if (this.active === true && s !== 'Lån registrert') {
            this.active = false;
        }
    }
});

/* Makro som logger returer */
macros.push({
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

/* Makro som logger ltst-besøk */
macros.push({
    siste_ltid: '',
    check: function (bibduck, bibsys) {
        if (bibsys.get(2, 1, 34) === 'Oversikt over lån og reserveringer') {
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
