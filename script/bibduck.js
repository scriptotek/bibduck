

var BibDuck = function () {

    var that = this;
    
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
    
    this.newInstance = function () {
        var inst = $('#instances .instance'),
            n = inst.length + 1,
            caption = 'BIBSYS ' + n,
            instanceDiv = $('<div class="instance" id="instance' + n + '"><a href="#" class="ui-icon ui-icon-close close"></a>' + caption + '</div>'),
            termLink = instanceDiv.find('a.close'),
            bib;
        $('#instances button.new').prop('disabled', true);
        bib = new Bibsys(true, n, that, 'Active', instanceDiv); //\\BIBSYS-auto');
        bib.ready(function () {
            $('#instances button.new').prop('disabled', false);

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

        /*
            bib2 = new Bibsys(false);
            bib2.ready(function() {
                alert('Great, BIBSYS 2 is ready');
            });
            */
        });
    };

    this.rfid = undefined;
    this.attachRFID = function (rfid) {
        this.rfid = rfid;
    }

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
        var bib = $.data(focused[0], 'bibsys'),
            state = that.rfid.checkBibsysState(bib);

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
    this.log('BIBDUCK is alive and quacking');
};
