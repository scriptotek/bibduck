/****************************************************************************
 * Bibsys class
 * Wrapper for a SecureNetTerm instance to BIBSYS
 ****************************************************************************/ 

function Bibsys(visible, index, bibduck, profile, instanceDiv) {

    var snt = new ActiveXObject('SecureNetTerm.Document'),
        sink = new ActiveXObject('EventMapper.SecureNetTerm'),
        ready_cbs = [],
        logger = bibduck.log,
        caption = 'BIBSYS ' + index,
        that = this,
        hist = '',
        trace = '';
    
    this.index = index;
    this.connected = false;
    
    this.ready = function (cb) {
        ready_cbs.push(cb);
    };
    
    this.get = function (y1, x1, y2, x2) {
        if (this.connected === true) {
            return snt.Get(y1, x1, y2, x2);
        } else {
            return '';
        }
    };

    this.quit = function () {
        snt.QuitApp();
    }

    // Checks if the instance is alive or it has been closed
    this.ping = function () {
        try {
            var pong = snt.WindowState;
        } catch (err) {
            return false;
        }
        return true;
    };

    this.setCaption = function(subcaption) {
        if (this.connected) {
            snt.Caption = caption + ' - ' + subcaption;
        }
    }

    this.onKeyDown = function(eventType, wParam, lParam) {
        switch (eventType.toString(16)) { // convert number to hex string
            case '102':
                eventTypeText = "WM_CHAR";
                break;
            case '104':
                eventTypeText = "WM_SYSKEYDOWN";
                break;
            case '100':
                eventTypeText = "WM_KEYDOWN";
                break;
        }

        switch (wParam) {
            // Return
            case 13:
                hist = hist + trace + '<br />';
                trace = ''
                break;
            
            // Tab
            case 9:
                if (trace != '') {
                    hist = hist + '&nbsp;&nbsp;&nbsp;&nbsp;' + trace + '<br />';
                    trace = ''
                }
                break;
            
            // Escape
            case 27:
                trace = "";
                break;
            
            // Space
            case 32:
                trace = trace + " ";
                break;
            
            // Backspace
            case 8:
                if (trace.length > 0) {
                    trace = trace.substr(0, trace.length-1);
                }
                break;
            
            // Forward-delete
            case 46:
                //if (trace.length > 0) {
                //  trace = trace.substr(0, trace.length-1);
                //}
                break;

            default:
                if (eventTypeText === "WM_CHAR") {
                    //snt.MessageBox(wParam);
                    trace = trace + String.fromCharCode(wParam);
                }
                //status = trace
        }
        if (trace.length >= 2 && trace.substr(trace.length-2, trace.length) === "!!") {
            s=""
            for (var i = 0; i < trace.length; i++) {
                s = s & "^H" //backspace
            }
            snt.QuickButton(s) 
            trace = ""
        }

        $('#statusbar').html(trace);
    }

    function wait_for(str, cb, delay) {
        var matchedstr;
        if (typeof(str) === 'string') str = [str]; // make array
        logger('Venter på: ' + str.join(' eller '));
        n = VBWaitForStrings(snt, str.join('|'));
        if (n === 0) {
            alert('Did not receive string "' + str + '"');
            return;
        }
        matchedstr = str[n-1];
        logger('Mottok: ' + matchedstr);
        if (delay == undefined) delay = 200;
        setTimeout(function() { cb(matchedstr); }, delay); // add a small delay
    }
    
    
    snt.Visible = visible;
    snt.WindowState = 1  //Normal (SW_SHOW)
    //snt.Synchronous = true;
    
    sink.Init(snt, 'OnKeyDown', function(eventType, wParam, lParam) {
        that.onKeyDown(eventType, wParam, lParam);
        bibduck.setFocus(that);
    });
    sink.Advise('OnMouseLDown', function(eventType, wParam, lParam) {
        bibduck.setFocus(that);
    });
    sink.Advise('OnConnected', function() {
        that.connected = true;
        bibduck.log('Connected');
    });
    sink.Advise('OnDisconnected', function() {
        that.connected = false;
        bibduck.log('Disconnected');
    });
    
    function klargjor() {
        snt.Send('s');
        snt.QuickButton('^M'); 
        wait_for('HJELP', function() {
            //snt.Synchronous = false;
            $.each(ready_cbs, function(k, cb) {
                if (ready_cbs.hasOwnProperty(k)) {
                    cb();
                }
            });
        });
    }
    
    if (snt.Connect(profile) == true) {
        snt.Caption = caption;
        wait_for('Terminaltype', function() {
            snt.QuickButton('^M'); 
            wait_for( ['Gi kode', 'Bytt ut'] , function(s) {
                if (s == 'Bytt ut') {
                    snt.QuickButton('^M'); 
                    wait_for('Gi kode', function() {
                        klargjor();
                    });
                } else {
                    klargjor();
                }
            });             
        });
    }
    
}