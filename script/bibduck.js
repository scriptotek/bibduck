
var BibDuck = function () {

    var that = this,
        shell = new ActiveXObject('WScript.Shell'),
        fso = new ActiveXObject('Scripting.FileSystemObject'),
        reg = new Registry(Registry.HKEY_CURRENT_USER),
        profiles = [], // SNetTerm profiles
        printers = [], // available printers (read from registry)
        backgroundInstance = null,
        mem_usage = 0.0,
        mem_warning_shown = false,
        date_started = new Date(),
        loglevel = 0,
        loglevels = ['debug','info','warn','error'];

    this.plugins = [];
    this.libnr = '';
    this.autoProfilePath = '';
    this.activeProfilePath = '';
    this.printerName = '\\\\winprint64\\ole';
    this.printerPort = '';

    function getAutoProfile() {
        var j;
        for (j = 0; j < profiles.length; j += 1) {
            if (profiles[j].path === that.autoProfilePath) {
                return profiles[j];
            }
        }
        return null;
    }

    function getActiveProfile() {
        var j;
        for (j = 0; j < profiles.length; j += 1) {
            if (profiles[j].path === that.activeProfilePath) {
                return profiles[j];
            }
        }
        return null;
    }

    this.sendSpecialKey = function(key) {
        // Use to send e.g. function keys
        // $.bibduck.sendSpecialKey('F9');
        shell.SendKeys('{' + key + '}');
    };

	this.bringToFront = function () {
        //logger('CAPTION:'+ caption);
        shell.AppActivate('BIBDUCK');
    };

    this.getBackgroundInstance = function() {
        return backgroundInstance;
    };

    this.removeFocus = function() {
        $('.instance').removeClass('focused');
    };

    this.setFocus = function(instance) {
        this.removeFocus();
        $('#instance' + instance.index).addClass('focused');
    };

    this.log = function(str, options) {
        /* options can be either 
            - a string specifying the log level ('DEBUG', 'INFO', 'ERROR', ...)
            - an object containing various options
         */
        var d = new Date(),
            ts = toSiffer(d.getHours()) + ':' + toSiffer(d.getMinutes()), // + ':' + toSiffer(d.getSeconds()) + '.' + d.getMilliseconds(),
            linebreak = true,
            timestamp = true,
            level = 'debug';
        if (typeof options === 'object') {
            if (options.hasOwnProperty('linebreak')) {
                linebreak = options.linebreak;
            }
            if (options.hasOwnProperty('timestamp')) {
                timestamp = options.timestamp;
            }
            if (options.hasOwnProperty('level')) {
                level = options.level;
            }
        } else if (typeof options === 'string') {
            level = options;
        }
        level = level.toLowerCase();
        var $s;
        if (timestamp) {
			$s = $('<div class="' + level + '"></div>');
            if (loglevels.indexOf(level) < loglevel) {
                 $s.hide();
            }

			$s.append('<span class="time">' + ts + '</span> ');
            switch (level.toLowerCase()) {
            case 'warn':
                $s.append('<span class="level">MERK</span> ');
                break;
            case 'error':
                $s.append('<span class="level">FEIL</span> ');
                break;
            case 'debug':
                $s.append('<span class="level">DBUG</span> ');
                break;
            default:
                $s.append('<span class="level">INFO</span> ');
                break;
            }
        } else {
			$s = $('#log div:last-child');
        }
		$s.append(str);
		//s += str + (linebreak ? '</div>' : '');
        $('#log').append($s);
        //$('#log').scrollTop($('#log')[0].scrollHeight);
        //$('#log-outer').stop().animate({ scrollTop: $("#log-outer")[0].scrollHeight }, 800);
        $('#log-outer').scrollTop($("#log-outer")[0].scrollHeight);
    };

    this.setLogLevel = function(level) {
        loglevel = level;
        for (var i = 0; i < loglevels.length; i++) {
            if (i < level) {
                $('.' + loglevels[i]).hide();
            } else {
                $('.' + loglevels[i]).show();
            }
        }
    };

    this.log('BIBDUCK is alive and quacking', 'info');
    var head = getCurrentDir() + '.git\\refs\\heads\\stable',
        headFile = fso.GetFile(head),
        headDate = new Date(Date.parse(headFile.DateLastModified)),
        sha = readFile(head);
    $('#statusbar').html('BIBDUCK, oppdatert <a href="https://github.com/scriptotek/bibduck/commit/' + sha + '" target="_blank">' + headDate.getDate() + '. ' + month_names[headDate.getMonth()] + ' ' + headDate.getFullYear() + ', kl. ' + headDate.getHours() + '.' + headDate.getMinutes() + '</a>.');

    this.newBibsysInstance = function () {
        var inst = $('#instances .instance'),
            n = inst.length + 1,
            caption = 'BIBSYS ' + n,
            instanceDiv = $('<div class="instance" id="instance' + n + '"><a href="#" class="ui-icon ui-icon-close close"></a>' + caption + '</div>'),
            termLink = instanceDiv.find('a.close'),
            bib,
            activeProfile = getActiveProfile();

        $('#loader-anim').show();

        //$('#instances button.new').prop('disabled', true);
        bib = new Bibsys(true, n, that.log, activeProfile.path); //\\BIBSYS-auto');

        //$('#instances button.new').prop('disabled', false);

        // Attach the Bibsys instance to the div
        $.data(instanceDiv[0], 'bibsys', bib);

        // and insert it into the DOM:
        $('#instances').append(instanceDiv);

        // Destroy on clicking the close button
        termLink.click(function(e) {
            var instanceDiv = $(e.target).parent(),
                bib = $.data(instanceDiv[0], 'bibsys');
            bib.quit();
            e.preventDefault();
            termLink.remove();
            //instanceDiv.remove();
        });
        instanceDiv.click(function(e) {
            var bib = $.data(instanceDiv[0], 'bibsys');
            e.preventDefault();
            bib.bringToFront();
            that.setFocus(bib);
        });

        bib.on('keypress', function (evt) {
            var j;
            that.setFocus(bib);
            for (j = 0; j < that.plugins.length; j += 1) {
                if (that.plugins[j].hasOwnProperty('keypress')) {
                    try {
                        that.plugins[j].keypress(bib, evt);
                    } catch (err) {
                        that.log('Plugin ' + j + ' keypress: ' + err.message, 'error');
                    }
                }
            }
        });
        bib.on('click', function () {
            that.setFocus(bib);
        });
		
		bib.on('captionChange', function(newCaption) {
			instanceDiv.text(newCaption);
		});

        bib.on('ready', function () {
            that.log('Instans klar');
			//bib.setSubCaption('');
            that.setFocus(bib);
            $('#loader-anim').hide();

        /*
            bib2 = new Bibsys(false);
            bib2.ready(function() {
                alert('Great, BIBSYS 2 is ready');
            });
            */
        });

        bib.on('disconnected', function () {
            var focused = $('.instance.focused');
            if (focused.length === 1 && focused.attr('id') === 'instance' + bib.index) {
                that.removeFocus();
            }
        });
		
		bib.on('cancelled', function () {
			$('#loader-anim').hide();
			setTimeout(function() {
				termLink.click();
			}, 500);
		});

    };

    // Returns the total memory usage of all the SNetTerm processes
    // in megabytes 
    this.getMemoryUsage = function() {
        var wmi = GetObject('winmgmts:\\\\.\\root\\cimv2'),
            processes = new Enumerator(wmi.ExecQuery('Select * From Win32_Process')),
            totmem = 0.0;

        for (; !processes.atEnd(); processes.moveNext()) {
            var proc = processes.item();
            if (proc.Name == 'SecureNetTerm.exe') {
                totmem += proc.WorkingSetSize / 1048576;
            }
        }
        return totmem;
    };

    /************************************************************
     * Innstillinger 
     ************************************************************/

    this.saveSettings = function() {
        var forWriting = 2,
            homeFolder = shell.ExpandEnvironmentStrings('%APPDATA%'),
            dir = homeFolder + '\\Bibduck',
            newlibnr = $('#settings-form input').val(),
            actp = parseInt($('#active_profile').val(), 10),
            autop = parseInt($('#auto_profile').val(), 10),
            stikkp = parseInt($('#stikk_skriver').val(), 10),
            file;

        if (autop === -1) {
            that.autoProfilePath = 'none';
        } else {
            that.autoProfilePath = profiles[autop].path;
        }
        that.activeProfilePath = profiles[actp].path;
        that.printerName = printers[stikkp].name;
        that.findPrinter();

        if (that.libnr !== newlibnr) {
            that.libnr = newlibnr;
			$('#libnr').text(newlibnr);
            $('#libnr').show();
			that.log('Nytt libnr. lagret: ' + newlibnr);
        }

        if (!fso.FolderExists(dir)) {
            fso.CreateFolder(dir);
        }

        file = fso.OpenTextFile(dir + '\\settings.txt', forWriting, true);
        file.WriteLine('libnr=' + that.libnr);
        file.WriteLine('activeProfilePath=' + that.activeProfilePath);
        file.WriteLine('autoProfilePath=' + that.autoProfilePath);
        file.WriteLine('printerName=' + that.printerName);
        file.close();

    };

    this.loadSettings = function() {
        var homeFolder = shell.ExpandEnvironmentStrings('%APPDATA%'),
            path = homeFolder + '\\Bibduck\\settings.txt',
            data = readFile(path).split(/\r\n|\r|\n/),
            line,
            i;

        for (i = 0; i < data.length; i += 1) {
            line = data[i].split('=');
            if (line[0] === 'libnr') {
                this.libnr = line[1];
                this.log('Vårt libnr. er ' + this.libnr);
				$('#libnr').text(this.libnr);
                $('#settings-form input').val(this.libnr);
            } else if (line[0] === 'autoProfilePath') {
                this.autoProfilePath = line[1];
            } else if (line[0] === 'printerName') {
                this.printerName = line[1];
            }
        }

        if (this.libnr === '') {
			$('#libnr').hide();
            this.log('Libnr. ikke satt! Velg innstillinger for å sette libnr.', 'warn');
        }
    };

    this.loadPlugins = function() {
        var path = getCurrentDir() + 'plugins\\',
            folder = fso.GetFolder(path),
            files = new Enumerator(folder.files),
            waitingfor = [];

        function allLoaded() {
            var j;
            that.log(that.plugins.length + ' plugins loaded', 'debug');
            for (j = 0; j < that.plugins.length; j += 1) {
                if (that.plugins[j].hasOwnProperty('initialize')) {
                    try {
                        that.plugins[j].initialize();
                    } catch (e) {
                        that.log('Plugin ' + j + ': ' + e.message, 'error');
                    }
                }
            }
        }

        this.plugins = [];
        for (; !files.atEnd(); files.moveNext()) {
            (function() {
                var relpath = files.item().path.substr(path.length);
                if (relpath.substr(relpath.length - 3) === '.js') {
                    waitingfor.push(relpath);
                    $.getScript('plugins/' + relpath, function() {
                        that.log('Loaded ' + relpath, 'debug');
                        waitingfor.splice($.inArray(relpath, waitingfor), 1);
                        if (waitingfor.length === 0) {
                            allLoaded();
                        }
                    }).fail(function() {
                        that.log('Failed to load plugin "' + relpath + '"', 'error');
                        waitingfor.splice($.inArray(relpath, waitingfor), 1);
                        if (waitingfor.length === 0) {
                            allLoaded();
                        }
                    });
                }
            })();
        }
    };

    $('#clear-btn').on('click', function () {
        $('#log').html('');
    });

    $('#reload-btn').on('click', function () {
        window.bibduck.loadPlugins();
    });

    $('#settings-btn').on('click', function () {
        $('#settings-form').slideDown();
    });
    $('#settings-form form').on('submit', function () {
        that.saveSettings();
        $('#settings-form').slideUp();
        return false;
    });
    $('#settings-form button').on('click', function () {
        if ($(this).attr('type') !== 'submit') {
            $('#settings-form').slideUp();
        }
    });


    /************************************************************
     * Tilbakemeldinger 
     ************************************************************/

    $('#kvakk-btn').on('click', function () {
        if (that.libnr === '') {
            alert("Du må sette libnummeret ditt først.");
            $('#settings-form').slideDown();
        } else {
            window.open('http://kvakk.biblionaut.net/?bib=' + that.libnr);
            /*$('#kvakk-form').slideDown();
            $('#kvakk-form iframe').attr('src', 'http://kvakk.biblionaut.net/?bib=' + that.libnr);   
            window.resizeTo(900, 800);*/
        }
    });

    $('.modal .close-btn button').on('click', function () {
        $('#kvakk-form').slideUp();
        window.resizeTo(600, 250);
    });

    /************************************************************
     * SNetTerm settings and profiles 
     ************************************************************/

    function readSNetTermProfileFile(filename) {
        var xmltext = readFile(filename),
            xml = $.parseXML(xmltext);
        $(xml).find('Site').each(function () {
            var $this = $(this),
                profile = {
                    name: $this.attr('Name'),
                    user: $this.attr('User'),
                    pass: $this.attr('Pass'),
                    path: $this.attr('Path'),
                    node: $this
                };
            profiles.push(profile);
        });
        return xml;
    }

    function writeSNetTermProfileFile(filename, data) {
        var forWriting = 2,
            file = fso.OpenTextFile(filename, forWriting, true);
        file.Write(data);
        file.Close();
    }

    function readSNetTermIniFile(filename) {
        var lines = readFile(filename).split(/\r\n|\r|\n/);
        $.each(lines, function(n, line) {
            line = line.split('=');
            /*
            if (line[0] === 'ActivePath') {
                $.each(profiles, function(idx, profile) {
                    if (profile.path === line[1]) {
                        that.activeProfilePath = profile.path;
                    }
                });
            }
            */
        });
    }

    this.readSNetTermSettings = function() {
        var userSiteFile = '', // Path to SecureCommon.xml
            userIniFile = '',  // Path to SecureCommon.ini
            regBase = 'Software\\InterSoft International, Inc.\\SecureNetTerm';

        reg.find(regBase, function(path, value) {
            var p = path.split('\\'),
                keyName = p[p.length-1];
            if (keyName === 'UserSiteFile') {
                userSiteFile = value;
            }
            if (keyName === 'UserIniFile') {
                userIniFile = value;
            }
            return true;
        });

        if (userSiteFile === '') {
            this.log('Registerfeil: Fant ikke ' + regBase + '\\UserSiteFile', 'error');
            return false;
        }

        var xmlobj = readSNetTermProfileFile(userSiteFile),            act_html = '',
            sel = '',
            bg_html = '<option value="-1">Ikke bruk bakgrunnsinstans</option>';
        readSNetTermIniFile(userIniFile);
        if (profiles.length === 0) {
            this.log('Fant ingen BIBSYS-profiler. Er SNetTerm installert riktig?', 'error');
            return false;
        }

        //this.log('Antall profiler: ' + profiles.length, 'debug');

        // Check if activeProfile has been set
        if (this.activeProfilePath === '') {
            // set default to first profile
            this.activeProfilePath = profiles[0].path;
        }

        // Check if autoProfile has been set
        if (this.autoProfilePath !== 'none') {
            /*
            Disable this option for now
            if (confirm('Vil du opprette en bakgrunnsprofil?')) {
                var newSite = getActiveProfile().node.clone();
                newSite.attr('Name', 'BIBSYS-bakgrunn');
                newSite.attr('Path', '\\BIBSYS-bakgrunn');
                $(xmlobj).find('Sites').append(newSite);
                var xmlstr = xmlobj.xml;
                writeSNetTermProfileFile(userSiteFile, xmlstr);
                this.autoProfilePath = newSite.attr('Path');
                this.log('Cloned active profile into '+  this.autoProfilePath);

                // reload profiles
                readSNetTermProfileFile(userSiteFile);
                this.log('Antall profiler: ' + profiles.length);
            } else {
                this.autoProfilePath = 'none';
            }
            */
            this.autoProfilePath = 'none';
        }


        // Update settings
        for (var j = 0; j < profiles.length; j++) {
            sel = (profiles[j].path === this.activeProfilePath) ? ' selected="selected"' : '';
            act_html += '<option value="' + j + '"'+sel+'>' + profiles[j].name + '</option>';
            sel = (profiles[j].path === this.autoProfilePath) ? ' selected="selected"' : '';
            bg_html += '<option value="' + j + '"'+sel+'>' + profiles[j].name + '</option>';
        }
        $('#active_profile').html(act_html);
        $('#auto_profile').html(bg_html);

        this.findPrinter();

        this.saveSettings();

        // Start autoProfile if set, and activeProfile
        var autoProfile = getAutoProfile();
        if (autoProfile !== null) {
            if (autoProfile.user === '' || autoProfile.pass === '') {
                alert('Profilen "' + autoProfile.name + '" er konfigurert som bakgrunnsinstans, men siden en bakgrunnsinstans ikke kan be om innlogginsopplysninger må du legge dette inn i profilen. Det gjør du i SNetTerms Profile Manager (husk å velge profilen "' + autoProfile.name + '"). Husk å trykke "Save & Exit" etterpå. Hvis du ikke ønsker å gjøre dette, kan du skru av bruk av bakgrunninstans i BIBDUCK-innstillingene.');
            } else {
                this.log('Starter bakgrunnsinstans...');
                backgroundInstance = new Bibsys(false, 999, this.log, autoProfile.path); //\\BIBSYS-auto');
                backgroundInstance.on('ready', function () {
                    that.log('Bakgrunnsinstans er klar');
                    // Auto-start a BIBSYS instance (after a little delay to let the screen update)
                    setTimeout(function() {
                        $('button.new').click();
                    }, 500);
                });
            }
        } else {

            // Auto-start a BIBSYS instance (after a little delay to let the screen update)
            setTimeout(function() {
                $('button.new').click();
            }, 500);
        }
    };

    this.findPrinter = function () {
        if (this.printerName === '') {
            this.log('Ingen stikkseddelskriver konfigurert.', 'warn');
            return false;
        }
        var basepath = 'Software\\Microsoft\\Windows NT\\CurrentVersion\\Devices',
            opt_html = '';
        printers = [];
        reg.find(basepath, function(path, value) {
            var keyName = path.substr(basepath.length + 1),
                port = value.split(',')[1],
                sel = '';
            //that.log(keyName);
            if (keyName === that.printerName) {
                that.printerPort = port;
                sel = ' selected="selected"';
            }
            opt_html += '<option value="' + printers.length + '"'+sel+'>' + keyName + '</option>';
            printers.push({ name: keyName, port: port });
            return true;
        });
        $('#stikk_skriver').html(opt_html);
        if (this.printerPort === '') {
            alert('Fant ikke stikkseddelskriveren "' + this.printerName + '"!');
            return false;
        }
        return true;
    };

    this.instances = function () {
        var inst = [];
        $('.instance').each(function(key, val) {
            inst.push({
                element: val,
                bibsys: $.data(val, 'bibsys')
            });
        });
        return inst;
    };

    $(window).on('unload', function() {
        if (backgroundInstance !== null) {
            backgroundInstance.quit();
        }
        $.each(that.instances(), function(k, instance) {
            instance.bibsys.quit();
        });

    });

    /************************************************************
     * Start urverket
     ************************************************************/

    this.update = function () {
        // Remember to use that instead of this, since we are in 
        // the window scope when called by SetTimeout

        // Check if all instances are alive
        $.each(that.instances(), function(idx, instance) {
            if (instance.bibsys.ping() === false) {
                // The instance has been closed. Let's remove from DOM
                that.log('Instans avsluttet');
                $(instance.element).remove();
            }
        });

        // Get the focused instance
        var focused = $('.instance.focused');
        if (focused.length !== 1) {
            setTimeout(that.update, 100);
            return;
        }

        //$('#statusbar').html($('#instances .instance').length + ' vinduer, '+ focused[0].id + ' i fokus');        
        var bib = $.data(focused[0], 'bibsys');
        bib.update();
        if (bib.idle) {
            $(focused[0]).addClass('idle');

            if (mem_usage > 1000 && !mem_warning_shown) {
                mem_warning_shown = true;
                var now = new Date(),
                    diff = (now.getTime() - date_started.getTime()) / 1000.0;  // seconds

                // Log it in order to get statistics on how often excessive memory usage occurs
                $.getJSON('http://labs.biblionaut.net/bibduck/logg.php?runtime=' + diff + '&mem=' + mem_usage + '&callback=?', function(json) {
                    //alert(json.msg);
                });

                alert('BIBDUCK har kjørt i ' + Math.round(diff / 360)*10 + ' timer. SNetTerm bruker nå ' + mem_usage + ' MB minne. Det anbefales at du omstarter BIBDUCK for å frigjøre minne.');
            }
        } else {
            $(focused[0]).removeClass('idle');
        }

        for (var j = 0; j < that.plugins.length; j++) {
            if (that.plugins[j].hasOwnProperty('update')) {
                try {
                    that.plugins[j].update(bib);
                } catch (e) {
                    if (that.plugins[j].hasOwnProperty('name')) {
                        that.log('Plugin "' + that.plugins[j].name + '" failed', 'error');
                        that.log(e.message, 'error');
                    } else {
                        that.log('Plugin ' + j + ' failed', 'error');
                        that.log(e.message, 'error');
                    }
                    that.plugins.splice(j, 1);
                    that.log(that.plugins.length + ' plugins left', 'debug');
                    break;
                }
            }
        }

        setTimeout(that.update, 100);
    };

    this.update_memory_usage = function() {
        var mem = that.getMemoryUsage();
        mem_usage = Math.round(mem*10)/10;
        $('#status-right').html(mem_usage + ' MB');
        setTimeout(that.update_memory_usage, 1000);
    };

    window.onerror = function(errorMsg, url, lineNumber) {
        that.log('<span title="' + url + ', line:' + lineNumber + '">' + errorMsg + '</span>', 'error'); // + '", line ' + lineNumber + ' : ' + url, 'error');
        return true;
    };

    this.loadSettings();
    if (this.readSNetTermSettings() === false) {
        this.log('Beklager, BIBDUCK kan ikke fortsette. Nå er det på tide å rope etter hjelp!', 'error');
        $('#loader-anim').hide();
        return;
    }

    // Clicking on the "new" button creates a new Bibsys instance 
    $('button.new').click(this.newBibsysInstance);

    $(document).bind('keydown', 'ctrl+r', function() {
        that.loadPlugins();
    });
    $(document).bind('keydown', 'ctrl+0', function() {
        that.setLogLevel(0);
    });
    $(document).bind('keydown', 'ctrl+1', function() {
        that.setLogLevel(1);
    });
    $(document).bind('keydown', 'ctrl+2', function() {
        that.setLogLevel(2);
    });
    $(document).bind('keydown', 'ctrl+3', function() {
        that.setLogLevel(3);
    });
    that.loadPlugins();

    setTimeout(this.update, 100);
    setTimeout(this.update_memory_usage, 1000);

};


$(document).ready(function() {

    window.bibduck = new BibDuck();
    $.bibduck = window.bibduck;

});
