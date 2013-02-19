

var BibDuck = function (macros) {

    var that = this,
        word = new ActiveXObject('Word.Application'),
        shell = new ActiveXObject('WScript.Shell'),
        fso = new ActiveXObject('Scripting.FileSystemObject'),
        reg = new Registry(Registry.HKEY_CURRENT_USER),
        profiles = [], // SNetTerm profiles
        activeProfile = -1,
        autoProfile = -1;

    this.numlock_enabled = function () {
        return word.NumLock; // Silly, but seems to be only way to get numlock state
    };

    this.libnr = '';
    this.autoProfileName = '';
    
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
            bib;
        //$('#instances button.new').prop('disabled', true);
        bib = new Bibsys(true, n, that, 'Active'); //\\BIBSYS-auto');

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

        bib.ready(function () {
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
    }

    this.rfid = undefined;
    this.attachRFID = function (rfid) {
        this.rfid = rfid;
    }


    /************************************************************
     * Innstillinger 
     ************************************************************/

    this.saveSettings = function() {
        var forWriting = 2,
            homeFolder = shell.ExpandEnvironmentStrings('%APPDATA%'),
            dir = homeFolder + '\\Bibduck',
            newlibnr = $('#settings-form input').val(),
            file;

        if (that.libnr != newlibnr) {
            that.libnr = newlibnr;
            that.log('Nytt libnr. lagret: ' + newlibnr);
        }

        if (!fso.FolderExists(dir)) {
            fso.CreateFolder(dir);
        }

        file = fso.OpenTextFile(dir + '\\settings.txt', forWriting, true),
        file.WriteLine('libnr=' + that.libnr  );
        file.WriteLine('autoProfileName=' + that.autoProfileName );
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
            } else if (line[0] === 'autoProfileName') {               
                this.autoProfileName = line[1];
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
            if (line[0] === 'ActivePath') {
                $.each(profiles, function(idx, profile) {
                    if (profile.path === line[1]) {
                        activeProfile = idx;
                    }
                });
            }
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

        var xmlobj = readSNetTermProfileFile(userSiteFile);
        readSNetTermIniFile(userIniFile);
        this.log('Antall profiler: ' + profiles.length);
        if (activeProfile !== -1) {
            this.log('Aktiv profil: ' + profiles[activeProfile].name);
        }
        if (this.autoProfileName === '') {
            prompt('Vil du opprette en bakgrunnsprofil?');
            var newSite = profiles[activeProfile].node.clone();
            newSite.attr('Name', 'BIBSYS-bakgrunn');
            newSite.attr('Path', '\\BIBSYS-bakgrunn');
            $(xmlobj).find('Sites').append(newSite);
            var xmlstr = xmlobj.xml;
            writeSNetTermProfileFile(userSiteFile, xmlstr);
            this.autoProfileName = 'BIBSYS-bakgrunn';
            this.saveSettings();
        } else {
            $.each(profiles, function(idx, profile) {
                that.log(profile.name +' == '+that.autoProfileName);
                if (profile.name == that.autoProfileName) {
                    autoProfile = idx;
                }
            });
            if (autoProfile === -1) {
                this.log('Feil, fant ikke bakgrunnsprofilen!');
            } else {
                this.log('Auto-profil: ' + profiles[autoProfile].name);                
            }
        }
        if (autoProfile !== -1) {
            this.log('Starter autoprofil...');
            var bakgrunnsbib = new Bibsys(false, 999, this, profiles[autoProfile].path); //\\BIBSYS-auto');
            bakgrunnsbib.ready(function () {
                that.log('BIBSYS instance is ready');
                // Auto-start a BIBSYS instance
                $('button.new').click();
            });
        } else {
            // Auto-start a BIBSYS instance
            $('button.new').click();
        }
    }

    $(window).on('unload', function() {
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
    }

    this.loadSettings();
    this.readSNetTermSettings();

    // Clicking on the "new" button creates a new Bibsys instance 
    $('button.new').click(this.newBibsysInstance);

    
    setTimeout(this.update, 100);

};
