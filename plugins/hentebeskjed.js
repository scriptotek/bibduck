
/*****************************************************************************
 * Klasse for å sende hentebeskjed fra DOKST-skjermen
 * - Gir feilmelding hvis vi ikke er på DOKST-skjermen ved start
 * - callback() kalles når skriptet er ferdig
 *****************************************************************************/
var Hentebeskjed = function(bibsys, ltid, dokid, callback) {

	this.send = function() {
	
		if (bibsys.get(2, 1, 28) !== 'Utlånsstatus for et dokument') {
			$.bibduck.log('send_hentebeskjed: Vi er ikke på DOKST-skjermen!', 'error');
			bibsys.alert('send_hentebeskjed: Vi er ikke på DOKST-skjermen!');
			//setWorking(false);
			bibsys.setBusy(false);
			return;
		}
		//$.bibduck.sendSpecialKey('F3'); // Har opplevd at jeg ikke har fått bekreftelse på hentebeskjed hvis sendt fra DOKST
		//bibsys.wait_for('BIBSYS UTLÅN', [2,63], function() {
		send_hentebeskjed_del2();
		//});
	
	};

	function send_hentebeskjed_del2() {
	
		$.bibduck.log('Sender hentebeskjed til ' + ltid + ' for dokument ' + dokid);

		bibsys.send('\tHENTB\n');
		bibsys.wait_for('Hentebrev til låntaker:', [7,15], function() {
			bibsys.send(dokid + ltid + '\n');

			bibsys.wait_for([
			
				['Hentebeskjed allerede sendt til', [1,1], function() {
					$.bibduck.log('[HENTB] Hentebeskjed er allerede sendt.', 'warn');
					bibsys.alert('Hentebeskjed er allerede sendt.');
					hentebeskjed_sendt();
				}],

				['Kryss av for ønsket valg', [16,8], function() {
					send_hentebeskjed_del3();
				}],

				['Ønsker du likevel å', [19,2], function() {
					  /*************** Eksempel-skjerm: *********************************************
					01:                                                                            
					02:					*** ADVARSEL ***                                                
					03:																					
					04:	Kat.: 1 Låntaker: uoxxxxxxxx Navn Navn                    
					05:																					
					06:																					
					07:																					
					08:																					
					09:																					
					10:	 Låntaker har kr. 200 i utestående gebyr, 199 er max.grense                     
					11:																					
					12:	 Låntaker har: 1 sistegangspurringer, 0 er max.grense                           
					13:																					
					14:																					
					15:																					
					16:																					
					17:																					
					18:																					
					19:	 Ønsker du likevel å sende beskjeden? Svar (J/N):             _                 
					20:																					
					21:																			 
					22:																			 
					23:																			 
					24:																			 
					25:																			 
					   **************** Eksempel-skjerm: *********************************************/
					// 'Ønsker du likevel å låne ut boka?'
					// 'Ønsker du likevel å sende beskjeden'
					// Dekker følgende:
					//    'Ugyldig LTID fra dato', [9,2]
					//    'Låntaker har: 1 erstatningskrav, 0 er max.grense' [11,2]
					//    'STOPPMELDING', [6,15]
					// Dekker trolig (ikke testet):
					//    'sistegangspurringer'
					//    'i utestående gebyr'

					if (!bibsys.confirm('Ønsker du å fortsette?', 'Sende hentebeskjed')) {
						$.bibduck.log('[HENTB] Avbrutt etter ønske.', 'info');
						bibsys.setBusy(false);
						return;
					}
					
					bibsys.send('J');
					setTimeout(function() {

						// J og <Enter> for å fortsette
						bibsys.send('\t\tJ\n');
						bibsys.wait_for('Kryss av for ønsket valg', [16,8], function() {
							send_hentebeskjed_del3();
						});
					
					}, 300);
				}],
								
				['KOMMENTAR', [2,22], function() {
					
					if (!bibsys.confirm('Ønsker du å fortsette?', 'Sende hentebeskjed')) {
						$.bibduck.log('[HENTB] Avbrutt etter ønske.', 'info');
						bibsys.setBusy(false);
						return;
					}

					// <Enter> for å fortsette
					bibsys.send('\n');
					bibsys.wait_for('Kryss av for ønsket valg', [16,8], function() {
						send_hentebeskjed_del3();
					});

				}],
				
				['KOMMENTAR', [7,36], function() {				
					/*************** Eksempel-skjerm: *********************************************
					Utlånsstatus for et dokument (DOkstat)                             BIBSYS UTLÅN 
					Gi kommando: HENTB                :                                  2013-10-21 
																									
						DOKID/REFÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿                 
								 ³                                                ³                 
								 ³                     KOMMENTAR                  ³                 
					 DOKID: 09pf0³                                                ³30               
								 ³Kat.: 1 Låntaker: uo00xxxxxx Navn Navn          ³                 
								 ³                                                ³                 
								 ³Bøker sendes til UHS                            ³                 
								 ³                                                ³                 
								 ³                                                ³                 
								 ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ                 
																									
																									
																									
																									
					 Et annet dokument har reserveringer.                                           
					 Kommando Flere(PF12) gir dokstat for evt andre dokument under dette objekt.    
					**************** Eksempel-skjerm: *********************************************/

					if (!bibsys.confirm('Ønsker du å fortsette?', 'Sende hentebeskjed')) {
						bibsys.setBusy(false);
						$.bibduck.log('[HENTB] Avbrutt etter ønske.', 'info');
						return;
					}

					// <Enter> for å fortsette
					bibsys.send('\n');
					bibsys.wait_for('Kryss av for ønsket valg', [16,8], function() {
						send_hentebeskjed_del3();
					});

				}]

			]);

		});
	
	}

	function send_hentebeskjed_del3() {
		bibsys.send('X\n');
		bibsys.wait_for([
			['Hentebeskjed er sendt', [1,1], function() {
				$.bibduck.log('[HENTB] Hentebeskjed sendt per sms', 'info');
				bibsys.resetPointer();
				hentebeskjed_sendt();
			}],
			['Registrer eventuell melding', [8,5], function() {
				$.bibduck.sendSpecialKey('F9');
				bibsys.wait_for([
					['Hentebeskjed er sendt', [1,1], function() {
						$.bibduck.log('[HENTB] Hentebeskjed sendt per epost', 'info');
						hentebeskjed_sendt();
					}],
					['DOKID/REFID/HEFTID/INNID', [5,5], function() {  // av og til får vi bare en blank DOKST-skjerm!
						$.bibduck.log('[HENTB] Hentebeskjed sannsynligvis sendt per epost', 'info');
						setTimeout(function() {
							bibsys.resetPointer();
							hentebeskjed_sendt();					
						}, 200);
					}]
				]);
			}],
			['DOKID/REFID/HEFTID/INNID', [5,5], function() {  // av og til får vi bare en blank DOKST-skjerm!
				$.bibduck.log('[HENTB] Hentebeskjed sannsynligvis sendt per sms', 'info');
				setTimeout(function() {
					bibsys.resetPointer();
					hentebeskjed_sendt();					
				}, 200);
			}]
		]);
	}

	function hentebeskjed_sendt(secondattempt) {

		/* var firstline = bibsys.get(1),
			m = firstline.match(/på (sms|Email) til (.+) merket (.+)/);
		
		var	name = m[2],
			nr = m[3].trim();
		if (nr === '') {
			$.bibduck.log('Fant ikke noe hentenr.', 'error');
			setWorking(false);
			return;
		}
		//$.bibduck.log(name + nr + siste_bestilling.laankopi);
		*/

		if (bibsys.getCursorPos().row == 3) {
			bibsys.send('dokst,' + dokid + '\n'); // for å oppfriske skjermen slik at status endres fra RES til AVH
		} else {
			bibsys.send('\tdokst,' + dokid + '\n'); // for å oppfriske skjermen slik at status endres fra RES til AVH
		}

		setTimeout(function() {

			var firstline = bibsys.get(1),
				m1 = firstline.match(/med hentedato: (.+) merket (.+)/),
				m2 = firstline.match(/på (sms|Email) til (.+) merket (.+)/),
				m3 = firstline.match(/på (sms|Email) til (.+)/); // mulig for veldig lange navn

			if (m1) {
				callback({
					hentefrist: m1[1],
					hentenr: m1[2]
				});
			// } else if (m2) {
			// 	callback({
			// 		hentenr: m2[2],
			// 		hentefrist: '-'
			// 	});
			} else {
				if (secondattempt) {
					$.bibduck.log('Finner ikke hentenr. på DOKST-skjermen!','error');
					$.bibduck.writeErrorLog(bibsys, 'dokst_hentenr_mangler1');
					bibsys.alert('Det ble sendt hentebeskjed, men finner ikke hentenr. på DOKST-skjermen. Prøv å gjenoppfrisk DOKST-skjermen og skriv ut stikkseddel derfra.');
					bibsys.setBusy(false);

				} else {
					hentebeskjed_sendt(true); // vi prøver en gang til;
				}
			}
		}, 250);

	}

};