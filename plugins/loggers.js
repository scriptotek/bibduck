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
$.bibduck.plugins.push({
    siste_ltid: '',
    siste_dokid: '',
    aktiv_ltid: '',
    siste_retur: '',
    utlaansskjerm: false,
    name: 'Logger',
	items: [],

    keypress: function (bibsys) {
        if (bibsys.getTrace() === 'ltid!') {
            bibsys.clearInput();
            bibsys.send(this.siste_ltid);
        } else if (bibsys.getTrace() === 'dokid!') {
            bibsys.clearInput();
            bibsys.send(this.siste_dokid);
        }
    },

	valid_ltid: function(ltid) {
		ltid.replace(' ', '');
		if (ltid.length !== 10) return false;
		if (!isNaN(parseInt(ltid.substring(0,1)))) return false;
		if (!isNaN(parseInt(ltid.substring(1,2)))) return false;
		if (isNaN(parseInt(ltid.substring(8,10)))) return false;
		return true;
	},

	valid_dokid: function(dokid) {
		dokid.replace(' ', '');
		if (dokid.length !== 9) return false;
		if (isNaN(parseInt(dokid.substring(0,2), 10))) return false;
		return true;
	},

	initialize: function() {
		var that = this;

		this.items.push({
			check: function(bibsys) {
				var ltid = bibsys.get(4, 15, 24);
				if ((bibsys.get(2, 1, 34) === 'Oversikt over lån og reserveringer')
					&& (that.valid_ltid(ltid))) {
						return {ltid: ltid};
				}
				return;
			},
			format: function(args) {
                return '[LOG] LTST for: ' + args.ltid;
            },
			active: false
		});

		this.items.push({
			check: function(bibsys) {
				var ltid = bibsys.get(18, 18, 27);
				if ((bibsys.get(2, 1, 34) === 'Opplysninger om låntaker (LTSØk)')
					&& (that.valid_ltid(ltid))) {
						return {ltid: ltid};
				}
				return;
			},
			format: function(args) {
                return '[LOG] LTSØK for: ' + args.ltid;
            },
			active: false
		});
		
		this.items.push({
			check: function(bibsys) {
				var dokst = bibsys.get(6, 31, 39);
				if ((bibsys.get(2, 1, 38) === 'Utlånsstatus for et dokument (DOkstat)')
					&& (that.valid_dokid(dokst))) {
						return {dokst: dokst};
				}
				return;
			},
			format: function(args) {
                return '[LOG] DOKstat for: ' + args.dokst;
            },
			active: false
		});

		this.items.push({
			check: function(bibsys) {
				var dokid = bibsys.get(6, 31, 39),
					ltid = bibsys.get(15, 16, 25);
				if ((bibsys.get(2, 1, 15) === 'Returnere utlån')
					&& (that.valid_ltid(ltid)) && (that.valid_dokid(dokid))) {
						return {ltid: ltid, dokid: dokid};
				}
				return;
			},
			format: function(args) {
                return '[LOG] Retur registrert: ' + args.dokid + ' fra ' + args.ltid;
            },
			active: false
		});

		this.items.push({
			check: function(bibsys) {
				var ltid = bibsys.get(1, 20, 29),
					dokid = bibsys.get(10, 7, 15);
				if ((bibsys.get(1, 1, 14) === 'Lån registrert')
					&& (that.valid_ltid(ltid)) && (that.valid_dokid(dokid))) {
						return {ltid: ltid, dokid: dokid};
				}
				return;
			},
			format: function(args) {
                return '[LOG] Utlån registrert: ' + args.dokid + ' fra ' + args.ltid;
            },
			active: false
		});

		this.items.push({
			check: function(bibsys) {
				var ltid = bibsys.get(19, 19, 29),
					dokid = bibsys.get(9, 31, 39);
				if ((bibsys.get(2, 1, 15) === 'Reservere (RES)')
					&& (that.valid_ltid(ltid)) && (that.valid_dokid(dokid))) {
						return {ltid: ltid, dokid: dokid};
				}
				return;
			},
			format: function(args) {
                return '[LOG] Reservering registrert: ' + args.dokid + ' for ' + args.ltid;
            },
			active: false
		});

		this.items.push({
			check: function(bibsys) {
				var bestnr = bibsys.get(1, 44, 52);
				if (bibsys.get(1, 1, 32) === 'Din kopibestilling er registrert') {
					return {bestnr: bestnr};
				}
				return;
			},
			format: function(args) {
                return '[LOG] Kopibestilling registrert: ' + args.bestnr;
            },
			active: false
		});

		/*this.items.push({
			check: function(bibsys) {
				var m = bibsys.get(1).match(/Hentebeskjed er sendt på (sms|Email) til (.+) merket (.+)/);
				if (m) {
					return {medium: m[1], name: m[2], hentenr: m[3]};
				}
				return;
			},
			format: function(args) {
                return 'Hentebeskjed sendt på ' + args.medium + ' til ' + args.name + ' merket ' + args.hentenr;
            },
			active: false
		});*/

	},

    update: function (bibsys) {
        var ltid,
            dokid,
			item,
			res;

		for (var i = 0; i < this.items.length; i++) {
			item = this.items[i];
			res = item.check(bibsys);
			if (!item.active && res) {
				item.active = true;
				if (!bibsys.busy) {
					$.bibduck.log(item.format(res), 'info');
				}
			} else if (item.active && !res) {
				item.active = false;
			}
		}
    }
});
