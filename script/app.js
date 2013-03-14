$(document).ready(function() {

    var bibduck = new BibDuck(triggers),
        rfid = new RFID(bibduck);
    

});

      /*
      reg.find('Software\\Microsoft\\Windows NT\\CurrentVersion\\Devices', function(path, value) {
        try {
          path = path.split('\\');
          value = value.split(',');
          var printer = path[path.length-1],
            port = value[value.length-1];
        } catch (e) { 
          // ignore
        }
        return true;
      });
      */
