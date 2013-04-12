if(typeof String.prototype.trim !== 'function') {
  String.prototype.trim = function() {
    return this.replace(/^\s+|\s+$/g, '');
  };
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
            click: [],
            disconnected: [],
            connected: []
        },
        caption = 'BIBSYS ' + index,
        user = '',
        that = this,
        hist = '',
        trace = '',
        currentscreen = '',
        prevscreen = '',
        currentscreenlines = [],
        waiters = [];

    //if (visible) {
        snt.WindowState = 2;  // Minimized
    //}
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
        /*
        var shell = new ActiveXObject('WScript.Shell'),
            cd = getCurrentDir(),
            tmpFile = cd + 'tmp.txt',
            exc = shell.Exec('"' + getCurrentDir() + 'klocks.exe"'),
            //exc = shell.Run('"' + cd + 'klocks.exe" > "' + tmpFile + '"', 0, true),
            //status = readFile(tmpFile),
            // split by whitespace:
            status = exc.StdOut.ReadLine().split(/\s/),
            nml_on = (status[0].split(':')[1] == 1);
        return nml_on;
        */
    };

    function trigger(eventName, obj) {
        if (obj === undefined) {
            obj = {};
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

    function exec_cb(itm, j) {
        logger(' OK', { timestamp: false });
        //logger('Got string: "' + itm.items[j].str + '" after ' + itm.attempts + ' iterations. ' + waiters.length + ' waiters left');
        setTimeout(function() {
            itm.items[j].cb();
        }, 50);
    }

    this.update = function () {
        prevscreen = currentscreen;
        if (this.connected === true) {
            // grab the whole screen, and split into lines of 80 chars each
            currentscreen = snt.Get(1, 1, 25, 80);
            currentscreenlines = currentscreen.match(/.{80}/g);
            if (currentscreenlines === null) {
                currentscreenlines = [];
            }
        } else {
            currentscreen = '';
            currentscreenlines = [];
        }
        for (i = 0; i < waiters.length; i++) {
            waiters[i].attempts += 1;
            if (waiters[i].attempts > 100) {
                logger('GIR OPP', {timestamp: false});
                logger('Mottok ikke den ventede responsen', 'error');
                waiters.splice(i, 1);
                return;
            }
            for (j = 0; j < waiters[i].items.length; j++) {
                if (waiters[i].items[j].col !== -1) {
                    if (currentscreenlines[waiters[i].items[j].line-1].indexOf(waiters[i].items[j].str) + 1 === waiters[i].items[j].col) {
                        exec_cb(waiters.splice(i, 1)[0], j);
                        return;
                    }
                } else {
                    if (currentscreenlines[waiters[i].items[j].line-1].indexOf(waiters[i].items[j].str) !== -1) {
                        exec_cb(waiters.splice(i, 1)[0], j);
                        return;
                    }
                }
            }
        }
        if (waiters.length > 0) {
            logger('.', { linebreak: false, timestamp: false });
        }
    };

    /*
        Returns an object with the current row and column of the cursor.
        The first row/column is 1, not 0.
     */
    this.getCursorPos = function () {
        if (this.connected === false) {
            return { row: 1, col: 1 };
        }
        return { row: snt.CurrentRow, col: snt.CurrentColumn};
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
        try {
            if (line === 0 || line > currentscreenlines.length) {
                return '';
            }
            if (line === undefined) {
                return currentscreenlines.join('\n');
            }
            if (start === undefined && end === undefined) {
                return currentscreenlines[line-1].trim();
            }
            if (end === undefined) {
                return currentscreenlines[line-1].substr(start - 1, 81 - start).trim();
            }
            return currentscreenlines[line-1].substr(start - 1, end - start + 1).trim();
        } catch (e) {
            return '';
        }
    };

    this.getCurrentLine = function() {
        return this.get(snt.CurrentRow);
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
    };

    this.typetext = function(str) {
        shell.SendKeys(str);
    };

    this.wait_for = function(str, cb, delay) {
        var matchedstr;
        if (typeof(str) === 'string') str = [str]; // make array
        logger('Venter på: ' + str.join(' eller ') + '... ', { linebreak: false, level: 'debug' });
        n = VBWaitForStrings(snt, str.join('|'));
        if (n === 0) {
            logger('Trondheim svarer ikke :(', { timestamp: false, level: 'error' });
            if ((typeof(cb) === 'object') && (cb.failure !== undefined)) {
                cb.failure();
            }
            return;
        }
        matchedstr = str[n-1];
        logger('OK', { timestamp: false });
        if (delay === undefined) delay = 200;
        setTimeout(function() {
            if ((typeof(cb) === 'object') && (cb.success !== undefined)) {
                cb.success(matchedstr);
            } else {
                cb(matchedstr);
            }
        }, delay); // add a small delay
    };

    this.wait_for2 = function(str, line, cb) {
        var col = -1,
            waiter = [];
        if (typeof(str) == 'string') {
            if (typeof(line) == 'object') {
                col = line[1];
                line = line[0];
            }
            waiter.push([str, [line, col], cb]);
        } else {
            waiter = str;
        }
        // objectify:

        var s = [];
        for (var j = 0; j < waiter.length; j++) {
            if (typeof(waiter[j][1]) == 'object') {
                col = waiter[j][1][1];
                line = waiter[j][1][0];
            } else {
                col = -1;
                line = waiter[j][1];
            }
            waiter[j] = {
                str: waiter[j][0],
                line: line,
                col: col,
                cb: waiter[j][2]
            };
            s.push(waiter[j].str + '(' + waiter[j].line + ',' + waiter[j].col + ')');
        }
        logger('Venter på ' + s.join(' eller ') + '.', { linebreak: false, level: 'debug' });
        waiters.push({attempts: 0, items: waiter});
        //     waiters.push({
        //         str: str,
        //         line: line,
        //         col: col,
        //         cb: cb,
        //         attempts: 0
        //     });
        // } else {

        // }
    };


    function getforwardchars(cr, cc) {
        var line = snt.Get(cr,1,cr,79);
        var endpos = line.indexOf("  ", cc);
        var todelete = endpos - cc;
        if (line.charAt(cc-1) != " ") todelete++;
        return todelete;
    }

    this.clearLine = function() {
        /* Prøver å tømme en linje med kolon i seg */
        sink.sleep(1);

        var count = 0,
            cr = snt.CurrentRow,
            cc = snt.CurrentColumn,
            line = snt.Get(cr,1,cr,79),
            startpos = line.lastIndexOf(":", cc) + 3,
            todelete = cc - startpos;

        //logger('current: ' + snt.CurrentRow + ',' + snt.CurrentColumn);
        //logger("Characters to back-delete: " + todelete, 'debug');

        while (todelete-- > 0) {
            if (count++ > 70) return;
            //logger('backspace');
            snt.QuickButton("^H");
        }
        count = 0;
        while (snt.CurrentColumn != startpos) {
            if (count++ > 100) break;
            sink.sleep(1);
        }
        cc = snt.CurrentColumn;


        todelete = getforwardchars(cr, cc);
        //logger("Characters to forward-delete: " + todelete, 'debug');

        while (todelete > 0) {
            if (count++ > 70) return;
            //logger('delete key');
            shell.SendKeys('{DEL}');
            //sink.sleep(1000);
            count = 0;
            while (getforwardchars(cr, cc) == todelete) {
                if (count++ > 30) {
                    logger('Delete key did not work!','error');
                    return;
                }
                sink.sleep(1);
            }
            todelete = getforwardchars(cr, cc);
        }

    };

    this.resetPointer = function () {
        /* Flytter pekeren til kommandolinja (linje 3) gjennom suksessiv tabbing */
        var cr = snt.CurrentRow,
            cc = snt.CurrentColumn,
            count = 0;
        while (snt.CurrentRow != 3) {
            if (count++ > 30) return false;
            //logger('tab from: ' + snt.CurrentRow + ',' + snt.CurrentColumn);
            snt.QuickButton("^I");
            do {
                //logger('sleep');
                sink.sleep(1); // Venter til pekeren faktisk har flyttet seg
            } while (cr == snt.CurrentRow && (cc == snt.CurrentColumn || cc+1 == snt.CurrentColumn));
            cr = snt.CurrentRow;
            cc = snt.CurrentColumn;
        }
        //logger('CLEAR LINE');
        this.clearLine();
        return true;
    };

    this.clearInput = function(length) {
        var cc = snt.CurrentColumn,
            count = 0,
            s = '';
        if (length === undefined) length = trace.length;
        for (var i = 0; i < length; i++) {
            s = s + '^H';    // backspace
        }
        snt.QuickButton(s);
        count = 0;
        do {
            if (count++>100) break;
            //logger(snt.CurrentRow + ',' + snt.CurrentColumn);
            sink.sleep(1); // Venter til pekeren faktisk har flyttet seg
        } while (snt.CurrentColumn > cc - trace.length + 1);
        trace = '';
    };

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

    /*
    This really doesn't work. The default title is always returned from SNetTerm :(
    this.getCaption = function () {
        if (this.connected) {
            return snt.Caption;
        } else {
            return '';
        }
    };
    */

    this.setCaption = function(subcaption) {
        if (this.connected) {
            if (subcaption === undefined) subcaption = '';
            snt.Caption = caption + ' : ' + this.user + ' - ' + subcaption;
        }
    };

    // private method
    function onKeyDown(eventType, wParam, lParam) {

        switch (wParam) {
            // Return
            case 13:
                hist = hist + trace + '<br />';
                trace = '';
                break;

            // Tab
            case 9:
                if (trace !== '') {
                    hist = hist + '&nbsp;&nbsp;&nbsp;&nbsp;' + trace + '<br />';
                    trace = '';
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

        $('#statusbar').html(trace);
    }

    this.getTrace = function() {
        return trace;
    };

    function klargjor() {
        that.send('u\n');
        that.wait_for2('HJELP', 25, function() {

            //snt.Synchronous = false;
            //logger('Numlock på? ' + (nml ? 'ja' : 'nei'), 'debug');
            if (nml) {
                // Turn numlock back on (it is disabled by SNetTerm when setting keyboard layout)
                shell.SendKeys('{numlock}');
            }
            if (visible) {
                snt.WindowState = 1;
                that.bringToFront();
            }
            trigger('ready');
        });
    }

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
        //SetSynchronous(snt);
        //snt.Synchronous = true;
        that.connected = true;
        that.user = snt.User;
        shell.AppActivate('BIBDUCK');
        logger('Tilkobla som "' + that.user + '"');
        that.wait_for2('Terminaltype', [25, 1], function() {
            nml = that.numlock_enabled();
            that.send('\n');
            that.wait_for2([
                ['Bytt ut', [23,1], function() {
                    that.send('\n');
                    that.wait_for2('Gi kode', [22, 6], function() {
                        klargjor();
                    });
                }],
                ['Gi kode', [22,6], function() {
                    klargjor();
                }]
            ]);
        });

    });
    sink.Advise('OnDisconnected', function() {
        that.connected = false;
        logger('Frakoblet');
        trigger('disconnected');
    });

    this.timer = function () {
        that.update();
        setTimeout(that.timer, 100);
    };

    this.bringToFront = function () {
        //logger('CAPTION:'+ caption);
        shell.AppActivate(caption);
    };

    function init() {
        // Bring window to front
        setTimeout(that.timer, 100);
        shell.AppActivate('BIBSYS');
        logger('Starter ny instans: ' + profile);

        if (snt.Connect(profile) === true) {
            snt.Caption = caption;
        }
    }
    setTimeout(init, 200); // a slight timeout is nice to give the GUI time to update

}