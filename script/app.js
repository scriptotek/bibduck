$(document).ready(function() {

    var bibduck = new BibDuck(),
        rfid = new RFID(bibduck),
        profiles = [],
        activeProfile = '';
    
    // Clicking on the "new" button creates a new Bibsys instance 
    $('#instances button.new').click(bibduck.newInstance);

    // Locate the profile file
    var reg = new Registry(Registry.HKEY_CURRENT_USER),
        UserSiteFile = '';

    function readProfiles(filename) {
        var xmltext = readFile(filename);
            xml = $.parseXML(xmltext);
        $(xml).find('Site').each(function () {
            var $this = $(this),
                name = $this.attr('Name'),
                user = $this.attr('User'),
                path = $this.attr('path'),
                logEntry = 'Fant profil: ' + name;
            profiles.push(path)
            if (user !== '') {
                logEntry += ', user: ' + user;
            }
            bibduck.log(logEntry);
        });
        //bibduck.log('Parsed ' + filename);
    }

    reg.find('Software\\InterSoft International, Inc.\\SecureNetTerm', function(path, value) {
          var p = path.split('\\'),
            keyName = p[p.length-1];
          //alert(keyName);
          if (keyName === 'ActiveProfile') {
            activeProfile = value;
            bibduck.log('Aktiv profil er: ' + value);
          }
          if (keyName === 'UserSiteFile') {
            UserSiteFile = value;
            readProfiles(value);
          }
          return true;
      });

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
