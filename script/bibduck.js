

var BibDuck = function (macros) {

    var that = this,
        shell = new ActiveXObject('WScript.Shell'),
        fso = new ActiveXObject('Scripting.FileSystemObject'),
        reg = new Registry(Registry.HKEY_CURRENT_USER),
        profiles = [], // SNetTerm profiles
        backgroundInstance = null;

    this.libnr = '';
    this.autoProfilePath = '';
    this.activeProfilePath = '';

    function getAutoProfile() {
        for (var j = 0; j < profiles.length; j++) {
            if (profiles[j].path == that.autoProfilePath) {
                return profiles[j];
            }
        }
        return null;
    }

    function getActiveProfile() {
        for (var j = 0; j < profiles.length; j++) {
            if (profiles[j].path == that.activeProfilePath) {
                return profiles[j];
            }
        }
        return null;
    }

    this.getBackgroundInstance = function() {
        return backgroundInstance;
    }

    this.setFocus = function(instance) {
        $('.instance').removeClass('focused');
        $('#instance' + instance.index).addClass('focused');
    };

    this.log = function(str, options) {
        var d = new Date(),
            ts = toSiffer(d.getHours()) + ':' + toSiffer(d.getMinutes()) + ':' + toSiffer(d.getSeconds()) + '.' + d.getMilliseconds(),
            linebreak = true,
            timestamp = true;
        if (typeof(options) === 'object') {
            if (options.hasOwnProperty('linebreak')) {
                linebreak = options['linebreak'];
            }
            if (options.hasOwnProperty('timestamp')) {
                timestamp = options['timestamp'];
            }
        }
        $('#log').append((timestamp?'[' + ts + '] ':'') + str + (linebreak?'<br />':''));
        //$('#log').scrollTop($('#log')[0].scrollHeight);
        $('#log-outer').stop().animate({ scrollTop: $("#log-outer")[0].scrollHeight }, 800);
    };

    this.log('BIBDUCK is alive and quacking, ' + macros.length + ' macros loaded.');
    var head = getCurrentDir() + '.git\\refs\\heads\\stable',
        headFile = fso.GetFile(head),
        headDate = new Date(Date.parse(headFile.DateLastModified)),
        sha = readFile(head);
    $('#statusbar').html('BIBDUCK, oppdatert <a href="https://github.com/scriptotek/bibduck/commit/' + sha + '" target="_blank">' + headDate.getDate() + '. ' + month_names[headDate.getMonth()] + ' '+ headDate.getFullYear() + ', kl. ' + headDate.getHours() + '.' + headDate.getMinutes() + '</a>.');

    this.newBibsysInstance = function () {
        var inst = $('#instances .instance'),
            n = inst.length + 1,
            caption = 'BIBSYS ' + n,
            instanceDiv = $('<div class="instance" id="instance' + n + '"><a href="#" class="ui-icon ui-icon-close close"></a>' + caption + '</div>'),
            termLink = instanceDiv.find('a.close'),
            bib,
            activeProfile = getActiveProfile();
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
                bib = $.data(instanceDiv[0], 'bibsys'),
                snt = bib.quit();
            e.preventDefault();
            instanceDiv.remove();
        });

        bib.on('keypress', function (e) {
            that.setFocus(bib);
            if (that.rfid !== undefined) {
                that.rfid.onKeyPress(e);
            }
        });
        bib.on('click', function (e) {
            that.setFocus(bib);
        });

        bib.on('ready', function (e) {
            that.log('BIBSYS instance is ready');
            bib.setCaption('RFID: ' + that.rfid.status());

        /*
            bib2 = new Bibsys(false);
            bib2.ready(function() {
                alert('Great, BIBSYS 2 is ready');
            });
            */
        });

    };

    this.instances = function () {
        return $('.instance');
    };

    this.rfid = undefined;
    this.attachRFID = function (rfid) {
        this.rfid = rfid;
    };


    /************************************************************
     * Innstillinger 
     ************************************************************/

    this.saveSettings = function() {
        var forWriting = 2,
            homeFolder = shell.ExpandEnvironmentStrings('%APPDATA%'),
            dir = homeFolder + '\\Bibduck',
            newlibnr = $('#settings-form input').val(),
            actp = parseInt($('#active_profile').val()),
            autop = parseInt($('#auto_profile').val()),
            file;

        if (autop === -1) {
            that.autoProfilePath = 'none';
        } else {
            that.autoProfilePath = profiles[autop].path;
        }
        that.activeProfilePath = profiles[actp].path;

        if (that.libnr != newlibnr) {
            that.libnr = newlibnr;
            that.log('Nytt libnr. lagret: ' + newlibnr);
        }

        if (!fso.FolderExists(dir)) {
            fso.CreateFolder(dir);
        }

        file = fso.OpenTextFile(dir + '\\settings.txt', forWriting, true),
        file.WriteLine('libnr=' + that.libnr);
        file.WriteLine('activeProfilePath=' + that.activeProfilePath );
        file.WriteLine('autoProfilePath=' + that.autoProfilePath );
        file.close();

    }

    this.loadSettings = function() {
        var homeFolder = shell.ExpandEnvironmentStrings('%APPDATA%'),
            path = homeFolder + '\\Bibduck\\settings.txt',
            data = readFile(path).split(/\r\n|\r|\n/),
            line;

        for (var i = 0; i < data.length; i++) {
            line = data[i].split('=');
            if (line[0] == 'libnr') {
                this.libnr = line[1];
                this.log('Vårt libnr. er ' + this.libnr);
                $('#settings-form input').val(this.libnr);
            } else if (line[0] === 'autoProfilePath') {
                this.autoProfilePath = line[1];
            }
        }

        if (this.libnr == '') {
            this.log('Libnr. ikke satt! Velg innstillinger for å sette libnr.');
        }
    }

    $('#clear-btn').on('click', function () {
        $('#log').html('');
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
        if ($(this).attr('type') != 'submit') {
            $('#settings-form').slideUp();
        }
    });


    /************************************************************
     * Tilbakemeldinger 
     ************************************************************/

    $('#kvakk-btn').on('click', function () {
        if (that.libnr == '') {
            alert("Du må sette biblioteksnr. ditt først.");
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
        var xmltext = readFile(filename);
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
            userIniFile = '';  // Path to SecureCommon.ini

        reg.find('Software\\InterSoft International, Inc.\\SecureNetTerm', function(path, value) {
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

        var xmlobj = readSNetTermProfileFile(userSiteFile),
            act_html = '',
            sel = '',
            bg_html = '<option value="-1">Ikke bruk bakgrunnsinstans</option>';
        readSNetTermIniFile(userIniFile);
        this.log('Antall profiler: ' + profiles.length);

        // Check if activeProfile has been set
        if (this.activeProfilePath === '') {
            // set default to first profile
            this.activeProfilePath = profiles[0].path;
        }

        // Check if autoProfile has been set
        if (this.autoProfilePath === 'none') {
            // pass
        } else if (this.autoProfilePath === '') {
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
    }

    $(window).on('unload', function() {
        if (backgroundInstance !== null) {
            backgroundInstance.quit();
        }
        $.each(that.instances(), function(k, instance) {
            $.data(instance, 'bibsys').quit();
        });

    });

    /************************************************************
     * Start urverket
     ************************************************************/

    this.update = function () {
        // Remember to use that instead of this, since we are in 
        // the window scope when called by SetTimeout

        if (that.rfid === undefined) {
            setTimeout(that.update, 100);
            return;
        }

        // Check if all instances are alive
        $('.instance').each(function(key, val) {
            var bib = $.data(val, 'bibsys');
            if (bib.ping() === false) {
                // The instance has been closed. Let's remove from DOM
                that.log('Instance killed');
                $(val).remove();
            }
        });

        // Get the focused instance
        var focused = $('.instance.focused');
        if (focused.length != 1) {
            setTimeout(that.update, 100);
            return;
        }

        //$('#statusbar').html($('#instances .instance').length + ' vinduer, '+ focused[0].id + ' i fokus');        
        var bib = $.data(focused[0], 'bibsys');
        bib.update();

        for (var j = 0; j < macros.length; j++) {
            macros[j].check(that, bib);
        }

        //$('#statusbar').html(bib.get(2, 1, 28));
        var state = that.rfid.checkBibsysState(bib);

        if (state === false) {
            setTimeout(that.update, 100);
            return; // Instance killed. We remove it on next iteration
        }
        // Check if RFID state of the focused instance has changed
        if (state !== that.rfid.state) {
            that.rfid.setState(state);
            $('.instance').each(function(key, val) {
                var bib = $.data(val, 'bibsys');
                bib.setCaption('RFID: ' + that.rfid.status());
            });
        }

        setTimeout(that.update, 100);
    };

    // Clicking on the "new" button creates a new Bibsys instance 
    $('button.new').click(this.newBibsysInstance);

    this.loadSettings();
    this.readSNetTermSettings();

    setTimeout(this.update, 100);

};
