/*****************************************************************************
 * Tillegg for å sette RFID-status basert på hvilken skjerm
 * som vises i BIBSYS
 *****************************************************************************/

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
};

window.RFID = {
    possibleStates: ['disabled', 'read', 'reg', 'ret'],
    controllerPath: 'C:\\RFIDIFControl\\RFIDIFControl.exe',
    guiPath: '"C:\\Program Files (x86)\\Bibliotheca RFID\\RFIDIF\\RFIDIF.exe"',
    fso: new ActiveXObject('Scripting.FileSystemObject'),
    objShell: new ActiveXObject('WScript.Shell'),
    statusStrings: {   // What is shown in the display
        'disabled': 'Skrudd av',
        'read': 'Lesing',
        'reg': 'Utlån',
        'ret': 'Retur'
    },
    state: 'na',
    enabled: false,

    status: function() {
        return this.statusStrings[this.state];
    },

    setState: function (state, force) {
        if (this.possibleStates.contains(state) === -1) {
            alert(state + ' er ikke en gyldig RFID-status!');
            return;
        }
        if (force !== true && this.state === state) {
            return;
        }
        // Note to self: Best would be to add a small delay here, that could be
        // cancelled, to avoid rapid flickering, but setTimeout is not available.
        // wsh.sleep is, but it cannot be cancelled.
        if (!this.enabled) {
            //snt.MessageBox(this.state + ' -> ' + state);
        }

        //window.bibduck.log('RFID status endret fra ' + this.state + ' til ' + state);
        this.state = state;
        $('#rfidstatus').html('RFID: ' + this.status());
        if (!this.enabled) {
            //snt.Caption = "BIBSYS - RFID (simulert): " + this.status();
        } else {
            //snt.Caption = "BIBSYS - RFID: " + this.status();
            switch (this.state) {
                case 'reg':
                    this.objShell.Run(this.controllerPath + ' SelectDeactivate', 7, false);
                    break;
                case 'ret':
                    this.objShell.Run(this.controllerPath + ' SelectActivate', 7, false);
                    break;
                case 'read':
                    this.objShell.Run(this.controllerPath + ' None', 7, false);
                    break;
                case 'disabled':
                    this.objShell.Run(this.controllerPath + ' DisableInput', 7, false);
                    break;
            }
        }
    },

    initialize: function () {
        if (this.fso.FileExists(this.controllerPath)) {
            this.enabled = true;

            var strComputer = '.',
                wmi = GetObject("winmgmts:" + "{impersonationLevel=impersonate}!\\\\" + strComputer + "\\root\\cimv2"),
                processes = new Enumerator(wmi.ExecQuery("Select * from Win32_Process")),
                foundProcess = false;
            processes.moveFirst();
            while (processes.atEnd() === false) {
                process = processes.item();
                if (process.Name == 'RFIDIF.exe') {
                    foundProcess = true;
                }
                processes.moveNext();
            }
            if (!foundProcess) {
                window.bibduck.log('Starter RFIDIF.exe', 'debug');
                this.objShell.Run(this.guiPath, 1, false);
            } else {
                window.bibduck.log('RFIDIF.exe kjører allerede: ' + process.ProcessId, 'debug');
            }

            //  if objProcess.Name = 'RFIDIF.exe'
            //window.bibduck.log('RFID OK');
        } else {
            this.enabled = false;
            //window.bibduck.log('Fant ikke RFID-controlleren: ' + this.controllerPath + '. RFID-støtte vil kun bli simulert.', 'warn');
        }
        this.setState('disabled');
    }

    //bibduck.attachRFID(this);

};

window.bibduck.plugins.push({

    name: 'RFID-plugin',

    initialize: function() {
        window.RFID.initialize();
    },

    keypress: function (bibsys, evt) {
        if (evt.type === 'WM_KEYDOWN' && evt.wParam === 114) {
            // pressed F3
            //window.bibduck.log("set state");
            window.RFID.setState(window.RFID.state);
        }
    },

    update: function (bibduck, bibsys) {
        var state = 'disabled';
        try {
            var line1 = bibsys.get(1, 1, 14),
                line2 = bibsys.get(2, 1, 28),
                line4 = bibsys.get(4, 1, 32);
            if (line2 === 'Registrere utlån (REG)') {
                state = 'reg';
            } else if (line2 === 'Fornye utlån (FORNy)') {
                state = 'reg';
            } else if (line2 === 'Returnere utlån (RETur)') {
                state = 'ret';
            } else if (line2 === 'Returnere innlån (IRETur)') {
                state = 'ret';
            } else if (line2 === 'Utlånsstatus for et dokument') {
                state = 'read';
            } else if (line2 === 'Bibliografisk søk (BIBsøk)') {
                state = 'read';
            } else if (line4 === 'Reserveringsliste (RLIST)') {
                state = 'read';
            } else if (line2 === 'Endre utlånsdata (ENdre)') {
                state = 'read';
            }
        } catch (err) {
            // pass
        }

        // Check if RFID state of the focused instance has changed
        if (state !== window.RFID.state) {
            window.RFID.setState(state);
            $('.instance').each(function(key, val) {
                var bib = $.data(val, 'bibsys');
                bib.setCaption('RFID: ' + window.RFID.status());
            });
        }

    }

});