
function Bibsys(visible, index, bibduck, rfid) {

    var snt = new ActiveXObject('SecureNetTerm.Document'),
		sink = new ActiveXObject('EventMapper.SecureNetTerm'),
		ready_cbs = [],
		logger = bibduck.log,
		caption = 'BIBSYS ' + index;
	
	this.index = index;
	
    this.ready = function(cb) {
	    ready_cbs.push(cb);
    };
	
	this.get = function(y1, x1, y2, x2) {
		return snt.Get(y1, x1, y2, x2);
	};
	
	this.getSnt = function() {
		return snt;
	};
	
	this.update = function() {
		rfid.check(snt);
	};
    
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
	
	var self = this;
	sink.Init(snt, 'OnKeyDown', function(eventType, wParam, lParam) {
		bibduck.setFocus(self);
    });
	sink.Advise('OnMouseLDown', function(eventType, wParam, lParam) {
		bibduck.setFocus(self);
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
    
	if (snt.Connect('Active') == true) {
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