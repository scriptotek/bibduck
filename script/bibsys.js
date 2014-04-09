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
            connected: [],
            cancelled: [],
            captionChange: [],
            waitFailed: []
        },
        nml = true,
        caption = 'BIBSYS #' + index,
        user = '',
        that = this,
        hist = '',
        trace = '',
        currentscreen = '',
        prevscreen = '',
        currentscreenlines = [],
        waiters = [],
        waiters2 = [], // new promise-based waiters
        last_activity,
        busy_since;

    //if (visible) {
        snt.WindowState = 2;  // Minimized
    //}
    this.index = index;
    this.idle = false;
    this.busy = false;
    this.silent = false; // supress alerts
    this.connected = false;
    this.idletime = 3.0;
    this.waitattempts_warn = 60;
    this.waitattempts = 300;
    this.wait_before_cb_exec = 50;

    this.alert = function(msg, title) {
        snt.MessageBox(msg, title);
    };
    
    this.setBusy = function(busy) {
        if (busy) {         
            $('#instance' + index).addClass('busy');
            busy_since = new Date();
        } else if (!busy && that.busy) {
            $('#instance' + index).removeClass('busy');

            var now = new Date(), diff2 = (now.getTime() - busy_since.getTime())/100.;
            $.bibduck.log('Det tok ' + (Math.round(diff2)/10) + ' sekunder');
            
            busy_since = 0;
        }
        that.busy = busy;
        $.bibduck.checkBusyStates();
    };
    
    this.confirm = function(msg, title) {
        var BUTTON_CANCEL = 1,     // OK and Cancel buttons
                    IDOK = 1,              // OK button clicked
                    IDCANCEL = 2;          // Cancel button clicked
        return (snt.MessageBox(msg, title, BUTTON_CANCEL) !== IDCANCEL);
    };
        
    this.on = function(eventName, cb) {
        if ($.inArray(eventName, Object.keys(cbs)) === -1) {
            alert("Unknown event '" + eventName + "'");
        } else {
            cbs[eventName].push(cb);
        }
    };

    this.off = function(eventName) {
        // removes *all* callbacks for the given event
        if ($.inArray(eventName, Object.keys(cbs)) === -1) {
            alert("Unknown event '" + eventName + "'");
        } else {
            cbs[eventName] = [];
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
        logger(' OK', { timestamp: false, level: 'debug' });
        //logger('Got string: "' + itm.items[j].str + '" after ' + itm.attempts + ' iterations. ' + waiters.length + ' waiters left');
        setTimeout(function() {
            var cb = itm.items[j].cb;
            cb();
        }, this.wait_before_cb_exec);
    }
    
    this.postError = function () {
        // http://blogs.msdn.com/b/ieinternals/archive/2010/05/13/xdomainrequest-restrictions-limitations-and-workarounds.aspx
        // http://blogs.msdn.com/b/ie/archive/2012/02/09/cors-for-xhr-in-ie10.aspx
        // http://stackoverflow.com/questions/9160123/no-transport-error-w-jquery-ajax-call-in-ie
        /*$.support.cors = true;
        $.ajax({
            type: 'POST',
            url: 'http://labs.biblionaut.net/bibduck/logg2.php',
            crossDomain: true,
            data: {
                log: $('#log').html(),
                screen: currentscreenlines.join("\n")
            }
        }).success(function(response) {
            alert(response);
        }).fail(function(jqxhr, textStatus, error) {
            var err = textStatus + ', ' + error;
            alert(err);
        });*/
        /*
        
         var XHR = new ActiveXObject("Msxml2.XMLHTTP");
         function callAjax(url){
           XHR.onreadystatechange=(callback);
           XHR.open("POST",url,true); //"POST" also works
           XHR.send("log=hello&screen=hello2"); // XHR.send("name1=value1&name2=value2");
         }

         function callback(){
           if(XHR.readystate == 4) alert("DONE\n" + XHR.responseText);
         }       
         callAjax('http://labs.biblionaut.net/bibduck/logg2.php');
        */
    };

    $(document).bind('keydown', 'ctrl+b', function() {
        that.postError();
    });
    
    this.unidle = function () {
        last_activity = new Date();
    };

    this.update = function () {
        prevscreen = currentscreen;
        if (this.connected === true) {
            // grab the whole screen, and split into lines of 80 chars each
            if (last_activity === undefined) {
                last_activity = new Date();
            }
            var now = new Date(),
                diff = (now.getTime() - last_activity.getTime())/1000.;
            // Idle for more than one second and not waiting for anything
            if (diff > this.idletime && waiters.length === 0 && waiters2.length === 0) {
                this.idle = true;
            } else {
                this.idle = false;
                currentscreen = snt.Get(1, 1, 25, 80); // Kan føre til feilmld. "No more threads can be created in the system."
                currentscreenlines = currentscreen.match(/.{80}/g);
                if (currentscreenlines === null) {
                    currentscreenlines = [];
                }
                if (currentscreen !== prevscreen) {
                    last_activity = new Date();
                }
            }
            
            if (that.busy) {
                var diff2 = (now.getTime() - busy_since.getTime())/100.;
                that.setSubCaption( (Math.round(diff2)/10) + ' s');
            }


        } else {
            currentscreen = '';
            currentscreenlines = [];
        }
        for (i = 0; i < waiters.length; i++) {
            waiters[i].attempts += 1;
            if (waiters[i].attempts == this.waitattempts_warn) {
                logger('(dette tar lang tid)', {timestamp: false});         
                $.bibduck.writeErrorLog(this, 'warn');
            }
            if (waiters[i].attempts > this.waitattempts) {
                logger('GIR OPP', {timestamp: false, level: 'error'});
                logger('Mottok ikke den ventede responsen', 'error');
                trigger('waitFailed', waiters[i]);

                waiters.splice(i, 1)

                //that.postError();
                $.bibduck.writeErrorLog(this, 'fail');
                if (!that.silent) {
                    that.alert('BIBSYS har gitt oss en uventet respons som Bibduck ikke forstår.');
                }
                this.setBusy(false);
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
        for (i = 0; i < waiters2.length; i++) {
            waiters2[i].attempts += 1;
            if (waiters2[i].attempts == this.waitattempts_warn) {
                logger('(dette tar lang tid)', {timestamp: false});         
                $.bibduck.writeErrorLog(this, 'warn');
            }
            if (waiters2[i].attempts > this.waitattempts) {
                logger('GIR OPP', {timestamp: false, level: 'error'});
                logger('Mottok ikke den ventede responsen', 'error');
                var waiter = waiters2.splice(i, 1)[0];
                waiter.promise.reject();
                this.setBusy(false);
                return;
            }
            for (j = 0; j < waiters2[i].items.length; j++) {
                var itm = waiters2[i].items[j];
                if (itm.col !== -1) {
                    if (currentscreenlines[itm.line-1].indexOf(itm.str) + 1 === itm.col) {
                        var waiter = waiters2.splice(i, 1)[0];
                        waiter.promise.resolve(itm);
                        return;
                    }
                } else {
                    if (currentscreenlines[itm.line-1].indexOf(itm.str) !== -1) {
                        exec_cb(waiters2.splice(i, 1)[0], j);
                        waiter.promise.resolve(itm);
                        return;
                    }
                }
            }
        }
        if (waiters.length > 0 || waiters2.length > 0) {
            logger('.', { linebreak: false, timestamp: false, level: 'debug' });
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

    this.getCurrentLine = function(format) {
        var s = this.get(snt.CurrentRow);
        if (format === 'lower') return s.toLowerCase();
        else return s;
    };

    this.getCurrentLineNumber = function() {
        return snt.CurrentRow;
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

    this.wait_for = function(str, line, cb) {
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
            s.push(waiter[j].str + '"(' + waiter[j].line + ',' + waiter[j].col + ')');
        }
        logger('Venter på "' + s.join(' eller "') + '.', { linebreak: false, level: 'debug' });
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
    
    /**
     * New wait_for method using promises
     */
    this.wait_for2 = function(str, line) {
        var deferred = Q.defer();
        var col = -1,
            waiter = [];
        if (typeof(str) == 'string') {
            if (typeof(line) == 'object') {
                col = line[1];
                line = line[0];
            }
            waiter.push([str, [line, col]]);
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
            };
            s.push(waiter[j].str + '"(' + waiter[j].line + ',' + waiter[j].col + ')');
        }
        logger('Venter på "' + s.join(' eller "') + '.', { linebreak: false, level: 'debug' });
        waiters2.push({attempts: 0, items: waiter, promise: deferred});
        //     waiters.push({
        //         str: str,
        //         line: line,
        //         col: col,
        //         cb: cb,
        //         attempts: 0
        //     });
        // } else {

        // }
        return deferred.promise;
    };


    function getforwardchars(cr, cc) {
        var line = snt.Get(cr,1,cr,79);
        var endpos = line.indexOf("  ", cc);
        var todelete = endpos - cc;
        if (line.charAt(cc-1) != " ") todelete++;
        return todelete;
    }

    this.microsleep = function() {
        // Sleep function for use in short while loops. Use with care, 
        // since loong sleep sessions will appear as freezes.
        sink.sleep(1);
    };

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

    this.getCaption = function() {
        if (this.user !== '') {
            return this.user + '@' + caption;
        } else {
            return caption;
        }
    };

    this.setSubCaption = function(subcaption) {
        if (this.connected) {
            if (subcaption === undefined) subcaption = '';
            snt.Caption = this.getCaption() + ' - ' + subcaption;
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

    function ready() {
        //snt.Synchronous = false;
        //logger('Numlock på? ' + (nml ? 'ja' : 'nei'), 'debug');
        if ($.bibduck.config.numLockFix &&  nml === true) {
            // Turn numlock back on (it is disabled by SNetTerm when setting keyboard layout)
            shell.SendKeys('{numlock}');
        }
        if (visible) {
            snt.WindowState = 1;
            that.bringToFront();
        }
        that.setBusy(false);
        trigger('ready');
    }

    function klargjor() {
        that.send('u\n');
        that.wait_for([
            ['Gi kommando', [3,1], ready],
            ['Gi kommando', [3,2], ready]
        ]);
    }

    sink.Init(snt, 'OnKeyDown', function(eventType, wParam, lParam) {
        last_activity = new Date();
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
        
        // Av en eller annen grunn går det ikke an å lese snt.Caption, kun
        // endre den... Vi må på en eller annen måte få gitt snt-objektet
        // informasjon om hvilken index det har. Det er ikke akkurat flust
        // av muligheter, så som et hack bruker vi snt.User. Dette er en 
        // variabel som vi kan endre og det ser ikke ut til å skape problemer
        // at vi endrer den.
        //snt.User = "" + that.index;
        
        shell.AppActivate('BIBDUCK');
        logger('Tilkobla som "' + that.user + '"');
        snt.Caption = that.getCaption();
        trigger('captionChange', that.getCaption());
        that.wait_for([
            ['Terminaltype', [25, 1], function() {

                if ($.bibduck.config.numLockFix) {
                    nml = that.numlock_enabled();
                }

                that.send('\n');
                that.wait_for([
                    ['Bytt ut', [23,1], function() {
                        that.send('\n');
                        that.wait_for([
                            ['Gi kode', [22, 6], klargjor],
                            ['Gi kommando', [3,1], ready],
                            ['Gi kommando', [3,2], ready],
                            ['Rutinesjekk', [9,18], function() {
                                  // En gang iblant (årlig?) får man denne meldingen:

                                  // ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
                                  // ³                                                     ³
                                  // ³  Rutinesjekk:                                       ³
                                  // ³  Kan du sjekke at epostadressen din er riktig?      ³
                                  // ³                                                     ³
                                  // ³  (Opplysningene kan også endres under valget        ³
                                  // ³  Brukerprofil/-opplysninger på hovedmenyen.)        ³
                                  // ³                                                     ³
                                  // ³  Rett eventuelt her og nå. Avslutt med PF2:         ³
                                  // ³                                                     ³
                                  // ³  d.m.heggo@ub.uio.no............................... ³
                                  // ³                                                     ³
                                  // ³                                                     ³
                                  // ³                                                     ³
                                  // ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
                                $.bibduck.log("BIBSYS ber om rutinesjekk av e-post","warn");
                                snt.WindowState = 1;
                                that.bringToFront();
                                trigger('ready');
                            }]
                        ]);
                    }],
                    ['Gi kode', [22,6], function() {
                        klargjor();
                    }],
                    ['Gi kommando', [3,1], function() {
                        ready();
                    }],
                    ['Gi kommando', [3,2], function() {
                        ready();
                    }],
                    ['Rutinesjekk', [9,18], function() {
                        $.bibduck.log("BIBSYS ber om rutinesjekk av e-post","warn");
                        snt.WindowState = 1;
                        that.bringToFront();
                        trigger('ready');
                    }]
                ]);
            }],
            ['Gi kommando:', [3,1], function() {
                ready();
            }]
        ]);
    });
    sink.Advise('OnDisconnected', function() {
        that.connected = false;
        that.user = '';
        logger('Frakoblet');
        trigger('disconnected');
        snt.Caption = that.getCaption();
        trigger('captionChange', that.getCaption());
    });

    this.timer = function () {
        that.update();
        setTimeout(that.timer, 100);
    };

    this.bringToFront = function () {
        //logger('CAPTION:'+ caption);
        shell.AppActivate(that.getCaption());
    };

    function init() {
        
        that.setBusy(true);
        
        // Bring window to front
        setTimeout(that.timer, 100);
        shell.AppActivate('BIBSYS');
        logger('Starter ny instans: ' + profile);

        if (snt.Connect(profile) === true) {
            //snt.Caption = caption;
            if (!snt.connected) {
                logger('Pålogging avbrutt');
                that.setBusy(false);
                trigger('cancelled');
            }
        } else {
            logger('Tilkobling avbrutt', 'warn');
        }
    }
    
    setTimeout(init, 200); // a slight timeout is nice to give the GUI time to update

}
