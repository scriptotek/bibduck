/*****************************************************************************
 * Tillegg som automatiserer utlån, stikkseddel-utskrift og hentebeskjed
 * for mottatte dokumenter.
 * slik at man slipper å tabbe ned til dokid-feltet. 
 *****************************************************************************/
$.bibduck.plugins.push({
    name: 'IMO/IRET auto-stikk',
	bestnr: '',

	send_hentb: function(bibsys, callback) {
		bibsys.resetPointer();
		bibsys.send('hentb,\n');
		bibsys.wait_for('Kryss av for ønsket valg', [16,8], function() {
			bibsys.send('X\n');
			bibsys.wait_for([
				['Hentebeskjed er sendt', [1,1], function() {
					$.bibduck.log('Hentebeskjed sendt per sms', 'info');
					bibsys.resetPointer();
					if (callback !== undefined) {
						bibsys.bringToFront();
						callback(bibsys);
					}
				}],
				['Registrer eventuell melding', [8,5], function() {
					$.bibduck.sendSpecialKey('F9');
					$.bibduck.log('Hentebeskjed sendt per epost', 'info');
					if (callback !== undefined) {
						bibsys.bringToFront();
						callback(bibsys);
					}
				}]
			]);
		});
	},

	stikkseddel: function(bibsys) {
		var sid = -1,
			splug,
			that = this;
		for (var j = 0; j < $.bibduck.plugins.length; j += 1) {
            if ($.bibduck.plugins[j].hasOwnProperty('name') && $.bibduck.plugins[j].name === 'Stikkseddel-tillegg') {
				sid = j;
				break;
			}
		}
		if (sid === -1) {
			$.bibduck.log('Fant ikke stikkseddel-tillegg', 'error');
			return;
		}
		splug = $.bibduck.plugins[sid];
		//$.bibduck.log('Hello ' + sid);
		splug.lag_stikkseddel(bibsys, function(data) {
			$.bibduck.log('stikkseddel ferdig');
			if (data.patron.kind === 'person' && data.patron.beststed === data.beststed) {
				$.bibduck.log('skal ikke sendes. La oss sente hentb');
				that.send_hentb(bibsys);
			}
		});
	},

	working: false,

    update: function (bibsys) {
        var dokid,
			bestnr,
			innid,
            cursorpos,
			laan,
			that = this;

		if ($.bibduck.libnr === '1030310' || $.bibduck.libnr === '1030317') {
			// Eksperimentelt tillegg, foreløpig skrur vi det bare på for UREAL og UREALINF

			// Har vi returnert noe?
			if ((bibsys.get(1, 11, 45) === 'er returnert både i INNLÅN og UTLÅN') || (bibsys.get(1, 11, 31) === 'er returnert i INNLÅN')) {
				if (this.working === true) return;
				this.working = true;

				that.stikkseddel(bibsys);

			// Har vi mottatt noe?
			} else if (bibsys.get(1, 11, 20) === 'er mottatt') {
				if (this.working === true) return;
				this.working = true;

				//$.bibduck.bringToFront(); // to avoid accidental key presses

				// Hva da?
				bestnr = bibsys.get(1, 1, 9);
				if (bestnr !== this.bestnr) {
					this.bestnr = bestnr;
					innid = bibsys.get(1, 31, 39);
					dokid = bibsys.get(4, 42, 50);
					laan = bibsys.get(8, 37, 37);
					ltid = bibsys.get(5, 13, 22);
					ltnavn = bibsys.get(5, 26, 61);

					if (laan === 'L') {
						$.bibduck.log('------------', 'info');
						$.bibduck.log('Mottok lån', 'info');
						$.bibduck.log('>  Bestnr: ' + bestnr + ', innid: ' + innid + ', dokid: ' + dokid, 'info');
						bibsys.resetPointer();
						bibsys.send('reg,\n');
						bibsys.wait_for('Registrere utlån', 2, function() {
							bibsys.send('\n');
							bibsys.wait_for([
								['Dokumentet er reservert for annen låntaker!', [1,1], function() {
									bibsys.send('\n');
									bibsys.wait_for('Lån registrert', [1,1], function() {
										that.stikkseddel(bibsys);
									});
								}],
								['Lån registrert', [1,1], function() {
									that.stikkseddel(bibsys);
								}],
								['Ugyldig LTID fra dato', [9,2], function() {
									// vi stanser her
								}]
							]);
						});

					} else {
						$.bibduck.log('------------', 'info');
						$.bibduck.log('Mottok kopi', 'info');
						$.bibduck.log('  Bestnr: ' + bestnr + ', innid: ' + innid + ', dokid: ' + dokid, 'info');
						/*
						bibsys.resetPointer();
						that.send_hentb(bibsys, function() {
							var ltnavn = bibsys.get(1, 34, 79);
							// Nå kan vi skrive ut navn og dato
							// Egentlig kan vi gjøre det direkte nå. Vi har både ltid og navn,
							// men inntil videre kan vi sende til ltsøk for å bruke kyrres navn og dato
							bibsys.send('ltsø,' + ltid + '\n');
							bibsys.wait_for('Ltid', [7,1], function() {
								bibsys.send('\n');
								bibsys.bringToFront();
							});
						});
						*/
					}
				}

			} else if (this.working === true) {
				this.working = false;
			}

        }
    }
});
