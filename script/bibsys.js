if(typeof String.prototype.trim !== 'function') {
  String.prototype.trim = function() {
    return this.replace(/^\s+|\s+$/g, ''); 
  }
}

/****************************************************************************
 * Bibsys class
 * Wrapper for a SecureNetTerm instance to BIBSYS
 ****************************************************************************/ 

function Bibsys(visible, index, logger, profile) {
    var snt = new ActiveXObject('SecureNetTerm.Document'),
        sink = new ActiveXObject('EventMapper.SecureNetTerm'),
        shell = new ActiveXObject('WScript.Shell'),
        // lists of callback functions for events:
        cbs = { 
            ready: [],
            keypress: [],
            click: []
        },
        caption = 'BIBSYS ' + index,
        user = '',
        that = this,
        hist = '',
        trace = '',
        currentscreen = [];
    this.index = index;
    this.connected = false;

    this.on = function(eventName, cb) {
        if ($.inArray(eventName, Object.keys(cbs)) === -1) {
            alert("Unknown event '" + eventName + "'");
        } else {
            cbs[eventName].push(cb);
        }
    };

    this.numlock_enabled = function () {
        // Silly, but seems to be only way to get numlock state??
        var word = new ActiveXObject('Word.Application'),
            nml_on = word.NumLock;
        word.Quit();
        return nml_on;
        var shell = new ActiveXObject('WScript.Shell'),
            cd = getCurrentDir(),
            tmpFile = cd + 'tmp.txt',
            exc = shell.Exec('"' + getCurrentDir() + 'klocks.exe"'),
            //exc = shell.Run('"' + cd + 'klocks.exe" > "' + tmpFile + '"', 0, true),
            //status = readFile(tmpFile),
            status = exc.StdOut.ReadLine(),
            // split by whitespace:
            status = status.split(/\s/),
            nml_on = (status[0].split(':')[1] == 1);
        return nml_on;
    };

    function trigger(eventName, obj) {
        if (obj === undefined) {
            obj = {}
        }
        obj.instance = that;
        if ($.inArray(eventName, Object.keys(cbs)) === -1) {
            alert("Unknown event " + eventName);
        }
        $.each(cbs[eventName], function(k, cb) {
            if (cbs[eventName].hasOwnProperty(k)) {
                cb(obj);
            }
        });
    }

    this.update = function () {
        if (this.connected === true) {
            // grab the whole screen, and split into lines of 80 chars each
            currentscreen = snt.Get(1, 1, 25, 80).match(/.{80}/g);
            if (currentscreen === null) {
                currentscreen = [];
            }
        } else {
            currentscreen = [];
        }
    };
    
    /*
        Returns content from the current Bibsys screen
        Line numbers and column numbers (start, end) start with index 1 (not 0)
        Examples: 
        * get(2, 1, 10) returns the content of line 2, from column 1 to 10.
        * get(2) returns the whole line 2
        * get() returns the whole screen (25 lines, 80 columns)
    */ 
    this.get = function (line, start, end) {
        if (line == 0 || line > currentscreen.length) {
            return '';
        }
        if (line === undefined) {  
            return currentscreen.join('\n');
        }
        if (start === undefined && end === undefined) {  
            return currentscreen[line-1].trim();
        }
        if (end === undefined) {  
            return currentscreen[line-1].substr(start - 1, 81 - start).trim();
        }
        return currentscreen[line-1].substr(start - 1, end - start + 1).trim();
    }

    this.quit = function () {
        snt.QuitApp();
    };

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
            snt.Caption = caption + ' : ' + this.user + ' - ' + subcaption;
        }
    };

    // private method
    function onKeyDown(eventType, wParam, lParam) {
        
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
                trace = '';
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
                if (eventType === "WM_CHAR") {
                    //snt.MessageBox(wParam);
                    trace = trace + String.fromCharCode(wParam);
                } else if (wParam >= 112 && wParam <= 123) {
                    // Function keys
                    trace = '';
                    if (wParam == 114) {
                        // F3 : Force-update RFID?
                    }
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
    };

    this.send = function(str) {
        var strs = str.split('\n'),
            strss = '';
        for (var i = 0; i < strs.length; i++) {
            strss = strs[i].split('\t');
            for (var j = 0; j < strss.length; j++) {
                if (strss[j] !== '') {
                    snt.Send(strss[j]);
                }
                if (j < strss.length-1) {
                    snt.QuickButton('^I');
                }
            }
            if (i < strs.length-1) {
                snt.QuickButton('^M');
            }
        }
    }

    this.wait_for = function(str, cb, delay) {
        var matchedstr;
        if (typeof(str) === 'string') str = [str]; // make array
        logger('Venter på: ' + str.join(' eller ') + '... ', { linebreak: false });
        n = VBWaitForStrings(snt, str.join('|'));
        if (n === 0) {
            logger('Tidsavbrudd!', { timestamp: false });
            if ((typeof(cb) === 'object') && (cb.failure !== undefined)) {
                cb.failure();
            }
            return;
        }
        matchedstr = str[n-1];
        logger('OK', { timestamp: false });
        if (delay == undefined) delay = 200;
        setTimeout(function() { 
          if ((typeof(cb) === 'object') && (cb.success !== undefined)) {
            cb.success(matchedstr);
          } else {
            cb(matchedstr); 
          }
        }, delay); // add a small delay
    }

    function klargjor() {
        that.send('u\n');
        that.wait_for('HJELP', function() {
            //snt.Synchronous = false;

            logger('Numlock på? ' + (nml ? 'ja' : 'nei'));
            if (nml) {
                // Turn numlock back on (it is disabled by SNetTerm when setting keyboard layout)
                shell.SendKeys('{numlock}');
            }
            trigger('ready');
        });
    }

    if (visible) {
        snt.Visible = visible;
        snt.WindowState = 1  //Normal (SW_SHOW)
    }
    //snt.Synchronous = true;

    sink.Init(snt, 'OnKeyDown', function(eventType, wParam, lParam) {
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
        onKeyDown(eventTypeText, wParam, lParam);
        trigger('keypress', { type: eventTypeText, wParam: wParam, lParam: lParam });
    });
    sink.Advise('OnMouseLDown', function(eventType, wParam, lParam) {
        trigger('click');
    });
    sink.Advise('OnConnected', function() {
        that.connected = true;
        that.user = snt.User;
        logger('Connected as ' + that.user);
        that.wait_for('Terminaltype', function() {
            nml = that.numlock_enabled();
            that.send('\n'); 
            that.wait_for( ['Gi kode', 'Bytt ut'] , function(s) {
                if (s == 'Bytt ut') {
                    that.send('\n'); 
                    that.wait_for('Gi kode', function() {
                        klargjor();
                    });
                } else {
                    klargjor();
                }
            });             
        });

    });
    sink.Advise('OnDisconnected', function() {
        that.connected = false;
        logger('Disconnected');
    });

    this.timer = function () {
        that.update();
        setTimeout(that.timer, 100);
    }    

    function init() {
        // Bring window to front
        setTimeout(that.timer, 100);
        shell.AppActivate('BIBSYS');
        logger('Starter ny instans: ' + profile);

        if (snt.Connect(profile) == true) {
            snt.Caption = caption;
        }
    }
    setTimeout(init, 200); // a slight timeout is nice to give the GUI time to update

}