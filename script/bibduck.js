/*
var snt, snt_bg, sink;

var trace = '',
    hist = ''
function snt_OnKeyDown(eventType, wParam, lParam) {
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
			//	trace = trace.substr(0, trace.length-1);
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
	
	
    
	
    $('#div1').html(hist + trace);
}
*/
var bibduck = new function () {
	
	this.setFocus = function(instance) {
		$('.instance').removeClass('focused'); 
		$('#instance' + instance.index).addClass('focused'); 
	};

	this.log = function(str) {
		$('#log').append(str + '<br />');
		//$('#log').scrollTop($('#log')[0].scrollHeight);
		$('#log').stop().animate({ scrollTop: $("#log")[0].scrollHeight }, 800);
	};
	
	this.newInstance = function () {
		var inst = $('#instances .instance'),
			n = inst.length + 1,
			caption = 'BIBSYS ' + n,
			instanceDiv = $('<div class="instance" id="instance' + n + '"><a href="#" class="ui-icon ui-icon-close close"></a>' + caption + '</div>'),
			termLink = instanceDiv.find('a.close'),
			bib;
		$('#instances button.new').prop('disabled', true);
		bib = new Bibsys(true, n, bibduck, rfid);
		bib.ready(function () {
			$('#instances button.new').prop('disabled', false);
			$.data(instanceDiv[0], 'bibsys', bib); 
			$('#instances button.new').before(instanceDiv);
			termLink.click(function(e) {
				var instanceDiv = $(e.target).parent(),
					bib = $.data(instanceDiv[0], 'bibsys'),
					snt = bib.getSnt();
				e.preventDefault();
				snt.QuitApp();
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
	
};

$(document).ready(function() {
	
	$('#instances button.new').click(bibduck.newInstance);
	
	var rfidStatus = '';
        
    function update() {
    	var instances = $('#instances .instance');
		$.each(instances, function(key, val) {
			var bib = $.data(val, 'bibsys');
			bib.update();
		});
		var rfidStatus2 = rfid.status();
		if (rfidStatus != rfidStatus2) {
			rfidStatus = rfidStatus2;
			$('#rfidstatus').html('RFID: ' + rfidStatus);
		}
		setTimeout(update, 100);
    }
	setTimeout(update, 100);

});