
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
 * RFID class
 ****************************************************************************/ 

var RFID = function(bibduck) {
    var possibleStates = ['disabled', 'read', 'reg', 'ret'],
        controllerPath = 'C:\\RFIDIFControl\\RFIDIFControl.exe',
        guiPath = 'C:\\RFIDIFControl\\RFIDIFControl.exe',
        fso = new ActiveXObject('Scripting.FileSystemObject'),
        objShell = new ActiveXObject('WScript.Shell'),
        statusStrings = {   // What is shown in the display
            'disabled': 'Skrudd av', 
            'read': 'Lesing', 
            'reg': 'Utlån', 
            'ret': 'Retur' 
        };
    
    this.state = '';
    this.enabled = false;

    this.status = function() {
        return statusStrings[this.state];
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

        bibduck.log('RFID status endret fra ' + this.state + ' til ' + state);
        this.state = state;
        $('#rfidstatus').html('RFID: ' + this.status());
        if (!this.enabled) {
            //snt.Caption = "BIBSYS - RFID (simulert): " + this.status();
        } else {
            //snt.Caption = "BIBSYS - RFID: " + this.status();
            switch (this.state) {
                case 'reg':
                    objShell.Run(controllerPath + ' SelectDeactivate', 7, false);
                    break;
                case 'ret':
                    objShell.Run(controllerPath + ' SelectActivate', 7, false);
                    break;
                case 'read':
                    objShell.Run(controllerPath + ' None', 7, false);
                    break;
                case 'disabled':
                    objShell.Run(controllerPath + ' DisableInput', 7, false);
                    break;
            }
        }
    };
    
    /* 
        Returns the state of a Bibsys instance, or false if the 
        associated SecureNetTerm windows was closed by the user.
    */
    this.checkBibsysState = function (bibsys) {
        try {
            var line1 = bibsys.get(1, 1, 14),
                line2 = bibsys.get(2, 1, 28),
                line4 = bibsys.get(4, 1, 32);           
        } catch (err) {
            return false;
        }
        
        if (line2 === 'Registrere utlån (REG)') {
            return 'reg';
        } else if (line2 === 'Fornye utlån (FORNy)') {
            return 'reg';
        } else if (line2 === 'Returnere utlån (RETur)') {
            return 'ret';
        } else if (line2 === 'Returnere innlån (IRETur)') {
            return 'ret';
        } else if (line2 === 'Utlånsstatus for et dokument') {
            return 'read';
        } else if (line2 === 'Bibliografisk søk (BIBsøk)') {
            return 'read';
        } else if (line4 === 'Reserveringsliste (RLIST)') {
            return 'read';
        } else {
            return 'disabled';
        }
    };    
        
    if (fso.FileExists(controllerPath)) {
        this.enabled = true;
        objShell.Run(guiPath, 1, false);
        bibduck.log('RFID OK');
    } else {
        this.enabled = false;
        bibduck.log('Fant ikke RFID-controlleren: ' + controllerPath + '. RFID-støtte vil bare bli simulert.');
    }
    this.setState('disabled');
    bibduck.attachRFID(this);
    
};