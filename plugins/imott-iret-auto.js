/*****************************************************************************
 * Tillegg som automatiserer utlån, stikkseddel-utskrift og hentebeskjed
 * for mottatte dokumenter.
 * 
 *****************************************************************************/
$.bibduck.plugins.push({
    name: 'IMO/IRET auto-stikk',
	bestnr: '',

	send_hentb: function(bibsys, callback) {
		var that = this;
		bibsys.resetPointer();
		bibsys.send('hentb,\n');
		bibsys.wait_for([
		
			['Kryss av for ønsket valg', [16,8], function() {
				that.send_hentb_steg2(bibsys, callback);
			}],

			['Ugyldig LTID fra dato', [9,2], function() {
				//var dt = bibsys.get(9,25,34);
				//$.bibduck.log('NB! Ugyldig LTID fra dato: ' + dt, 'WARN');
				bibsys.send('J\n');
				bibsys.wait_for('Kryss av for ønsket valg', [16,8], function() {
					that.send_hentb_steg2(bibsys, callback);
				});
			}]
			
		]);
	},
	
	send_hentb_steg2: function(bibsys, callback) {
		bibsys.send('X\n');
		bibsys.wait_for([
			['Hentebeskjed er sendt', [1,1], function() {
				$.bibduck.log('Hentebeskjed sendt per sms', 'info');
				bibsys.resetPointer();
				if (callback !== undefined) {
					bibsys.bringToFront();
					setTimeout(function() {
						callback(bibsys);
					}, 200);
				}
			}],
			['Registrer eventuell melding', [8,5], function() {
				$.bibduck.sendSpecialKey('F9');
				bibsys.wait_for([
					['Hentebeskjed er sendt', [1,1], function() {
						$.bibduck.log('Hentebeskjed sendt per epost');
						setTimeout(function() {
							//bibsys.resetPointer();
							if (callback !== undefined) callback(bibsys);
						}, 200);
					}],
					['DOKID/REFID/HEFTID/INNID', [5,5], function() {  // av og til får vi bare en blank DOKST-skjerm!
						$.bibduck.log('Hentebeskjed sannsynligvis sendt per epost');
						setTimeout(function() {
							bibsys.resetPointer();
							if (callback !== undefined) callback(bibsys);
						}, 200);
					}]
				]);

			}]
		]);
	},

	stikkseddel: function(bibsys, options) {
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
			if (data.patron.kind === 'person' && data.patron.beststed === data.beststed) {
				$.bibduck.log('Dokumentet skal ikke sendes. La oss sente hentb');
				that.send_hentb(bibsys);
			}
		}, options);
	},

	working: false,
	
	siste_mottak: {},

    update: function (bibsys) {
        var dokid,
			bestnr,
			innid,
            cursorpos,
			laan,
			that = this;

		if ($.bibduck.config.autoImoEnabled) {
			// Eksperimentelt tillegg, foreløpig skrur vi det bare på for UREAL og UREALINF

			// Har vi sendt en bestilling?
			if (bibsys.get(1, 1, 32) === 'Din lånebestilling er registrert') {
				if (this.working === true) return;
				this.working = true;
				var bestnr = bibsys.get(1).match(/BESTNR = (b[0-9]+)/)[1];
				$.bibduck.log('Sendt bestilling med bestnr: ' + bestnr, 'info');
				
			// Har vi returnert noe?
			} else if ((bibsys.get(1, 11, 45) === 'er returnert både i INNLÅN og UTLÅN') || (bibsys.get(1, 11, 31) === 'er returnert i INNLÅN')) {
				if (this.working === true) return;
				this.working = true;

				$.bibduck.log('Lager stikkseddel automatisk for IRETur (imot-iret-auto.js)', 'info');
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
						$.bibduck.log('Mottok innlånt dokument', 'info');
						$.bibduck.log('> Bestnr: ' + bestnr + ', innid: ' + innid + ', dokid: ' + dokid + ', ltid: ' + ltid, 'info');

						var options = { bestnr: bestnr, artikkelkopi: false };

						bibsys.resetPointer();
						bibsys.send('reg,\n');
						bibsys.wait_for('Registrere utlån', 2, function() {
							bibsys.send('\n');
							bibsys.wait_for([
								['Dokumentet er reservert for annen låntaker!', [1,1], function() {
									bibsys.send('\n');
									bibsys.wait_for('Lån registrert', [1,1], function() {
										that.stikkseddel(bibsys, options);
									});
								}],
								['Lån registrert', [1,1], function() {
									that.stikkseddel(bibsys, options);
								}],
								['Ugyldig LTID fra dato', [9,2], function() {
									var dt = bibsys.get(9,25,34);
									$.bibduck.log('NB! Ugyldig LTID fra dato: ' + dt, 'WARN');
									bibsys.send('J\n');
									bibsys.wait_for('Lån registrert', [1,1], function() {
										that.stikkseddel(bibsys, options);
									});
								}]
							]);
						});

					} else {
						$.bibduck.log('------------', 'info');
						$.bibduck.log('Mottok kopi', 'info');
						$.bibduck.log('> Bestnr: ' + bestnr + ', innid: ' + innid + ', dokid: ' + dokid + ', ltid: ' + ltid, 'info');

						var options = { bestnr: bestnr, artikkelkopi: true };

						bibsys.resetPointer();
						that.send_hentb(bibsys, function() {
							that.stikkseddel(bibsys, options);
						});
					}
				}

			} else if (this.working === true) {
				this.working = false;
			}

        }
    },

	initialize: function() {

		$('#settings-form table').append('<tr>' +
		  '<th>' +
		   ' Auto-IMO/IRET' +
			'</th><td>' +
			'<input type="checkbox" id="auto_imo" ' + ($.bibduck.config.autoImoEnabled ? ' checked="checked"' : '') + '>' +
		   '   <label for="auto_imo">Automatisk IMO-behandling (og automatisk stikkseddel ved IRET)</label>' +
			' </td>' +
			'</tr>');

	},

	/**
	 * Kalles av Bibduck når innstillingene skal lagres
	 */
	saveSettings: function(file) {

		$.bibduck.config.autoImoEnabled = $('#auto_imo').is(':checked');
		file.WriteLine('autoImoEnabled=' + ($.bibduck.config.autoImoEnabled ? 'true' : 'false') );

	},

	/**
	 * Kalles av Bibduck når innstillingene skal lastes
	 */
	loadSettings: function(data) {

		// Default
		$.bibduck.config.autoImoEnabled = false;

		var line;
		for (var i = 0; i < data.length; i += 1) {
			line = data[i]
			if (line[0] === 'autoImoEnabled') {
				$.bibduck.config.autoImoEnabled = (line[1] === 'true');
			}
		}

	}

});
