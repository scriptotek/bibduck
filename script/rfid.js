
if (!String.prototype.trim) {
	// String.trim() was added natively in JavaScript 1.8.1 / ECMAScript 5 / IE9
	String.prototype.trim=function(){return this.replace(/^\s+|\s+$/g, '');};
}

Array.prototype.contains = function(str) {
	// Array.indexOf was added to JavaScript at some point, and could be used instead, 
	// but it was not available with the runtime used while writing this script.
    for (var i = 0; i < this.length; i++) {
        if (this[i] === str) {
            return true;
        }
    }
    return false;
}

/****************************************************************************
 * RFID
 ****************************************************************************/ 

/* A note on the construct below (which may seem "weird" at first):
 * We could first define a function, `var RFID = function() { ... };`, and then 
 * construct an object using the function as constructor; `var rfid = new RFID();` 
 * (parentheses are optional). But since we only need one instance, we can rather 
 * use an anonymous function as constructor; `var rfid = new (function() { ... })();`, 
 * and skipping the parentheses; `var rfid = new function() { ... };`
 */
var rfid = new function() {
	var possibleStates = ['disabled', 'read', 'reg', 'ret'],
		controllerPath = 'C:\\RFIDIFControl\\RFIDIFControl.exe',
		fso = new ActiveXObject("Scripting.FileSystemObject"),
		statusTexts = { 'disabled': 'Skrudd av', 'read': 'Lesing', 'reg': 'Utlån', 'ret': 'Retur' },
		objShell = new ActiveXObject("WScript.Shell");
	
	this.state = '';

	this.status = function() {
		return statusTexts[this.state];
	}
	
	this.setState = function (state) {
		if (possibleStates.contains(state) === -1) {
			alert(state + ' er ikke en gyldig RFID-status!');
			return;
		}
		if (this.state === state) { 
			return;
		}
		// Note to self: Best would be to add a small delay here, that could be
		// cancelled, to avoid rapid flickering, but setTimeout is not available.
		// wsh.sleep is, but it cannot be cancelled.
		if (!this.enabled) {
			//snt.MessageBox(this.state + ' -> ' + state);
		}
		this.state = state;
		if (!this.enabled) {
			//snt.Caption = "BIBSYS - RFID (simulert): " + this.status();
		} else {
			//snt.Caption = "BIBSYS - RFID: " + this.status();
			switch (this.state) {
				case 'reg':
					objShell.Run('C:\\RFIDIFControl\\RFIDIFControl.exe SelectDeactivate', 7, false);
					break;
				case 'ret':
					objShell.Run('C:\\RFIDIFControl\\RFIDIFControl.exe SelectActivate', 7, false);
					break;
				case 'read':
					objShell.Run('C:\\RFIDIFControl\\RFIDIFControl.exe None', 7, false);
					break;
				case 'disabled':
					objShell.Run('C:\\RFIDIFControl\\RFIDIFControl.exe DisableInput', 7, false);
					break;
			}
		}
	};
	
	this.check = function (snt) {
		var line1 = snt.Get(1,1,1,14).trim(),
			line2 = snt.Get(2,1,2,30).trim(),
			line4 = snt.Get(4,1,4,32).trim();
		
		if (line2 === 'Registrere utlån (REG)') {
			rfid.setState('reg');
		} else if (line2 === 'Returnere utlån (RETur)') {
			rfid.setState('ret');
		} else if (line2 === 'Returnere innlån (IRETur)') {
			rfid.setState('ret');
		} else if (line2 === 'Utlånsstatus for et dokument (') {
			rfid.setState('read');
		} else if (line2 === 'Bibliografisk søk (BIBsøk)') {
			rfid.setState('read');
		} else if (line4 === 'Reserveringsliste (RLIST)') {
			rfid.setState('read');
		} else {
			rfid.setState('disabled');
		}
	};
	
	if (fso.FileExists(controllerPath)) {
		this.enabled = true;
	} else {
		this.enabled = false;
		//snt.MessageBox("Finner ikke RFIDIFControl.exe. RFID-støtte skrus av");
	}
	this.setState('disabled');	
	
}();