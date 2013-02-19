

var BibDuck = function (macros) {

    var that = this;

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

    this.libnr = '';
    
    this.setFocus = function(instance) {
        $('.instance').removeClass('focused'); 
        $('#instance' + instance.index).addClass('focused'); 
    };

    this.log = function(str) {
        var d = new Date();
        var ts = toSiffer(d.getHours()) + ':' + toSiffer(d.getMinutes()) + ':' + toSiffer(d.getSeconds()) + '.' + d.getMilliseconds();
        $('#log').append('[' + ts + '] ' + str + '<br />');
        //$('#log').scrollTop($('#log')[0].scrollHeight);
        $('#log-outer').stop().animate({ scrollTop: $("#log-outer")[0].scrollHeight }, 800);
    };

    this.log('BIBDUCK is alive and quacking, ' + macros.length + ' macros loaded.');
    
    this.newInstance = function () {
        var inst = $('#instances .instance'),
            n = inst.length + 1,
            caption = 'BIBSYS ' + n,
            instanceDiv = $('<div class="instance" id="instance' + n + '"><a href="#" class="ui-icon ui-icon-close close"></a>' + caption + '</div>'),
            termLink = instanceDiv.find('a.close'),
            bib;
        //$('#instances button.new').prop('disabled', true);
        bib = new Bibsys(true, n, that, 'Active', instanceDiv); //\\BIBSYS-auto');

        //$('#instances button.new').prop('disabled', false);

        // Attach the Bibsys instance to the div
        $.data(instanceDiv[0], 'bibsys', bib); 

        // and insert it into the DOM:
        $('#instances button.new').before(instanceDiv);

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
            shell = new ActiveXObject('WScript.Shell'),
            fso = new ActiveXObject('Scripting.FileSystemObject'),
            homeFolder = shell.ExpandEnvironmentStrings('%APPDATA%'),
            dir = homeFolder + '\\Bibduck',
            newlibnr = $('#settings-form input').val(),
            file;

        if (that.libnr != newlibnr) {

            if (!fso.FolderExists(dir)) {
                fso.CreateFolder(dir);
            }

            file = fso.OpenTextFile(dir + '\\settings.txt', forWriting, true),
            file.WriteLine('libnr=' + newlibnr);
            file.close();

            that.libnr = newlibnr;
            that.log('Nytt libnr. lagret: ' + newlibnr);
        }

    }

    this.loadSettings = function() {
        var shell = new ActiveXObject('WScript.Shell'),
            homeFolder = shell.ExpandEnvironmentStrings('%APPDATA%'),
            path = homeFolder + '\\Bibduck\\settings.txt',
            data = readFile(path).split('\r\n'),
            line;

        for (var i = 0; i < data.length; i++) {
            line = data[i].split('=');
            if (line[0] == 'libnr') {
                this.libnr = line[1];
                this.log('Vårt libnr. er ' + this.libnr);
                $('#settings-form input').val(this.libnr);
            }
        }

        if (this.libnr == '') {
            this.log('Libnr. ikke satt! Velg innstillinger for å sette libnr.');
        }
    }

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

    this.loadSettings();

    var newWindow;


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

    setTimeout(this.update, 100);

};
