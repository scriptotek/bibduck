/*****************************************************************************
 * Nye kommandoer:
 *   roald! : Åpner Roald
 *****************************************************************************/
$.bibduck.plugins.push({

    name: 'Roald',

    roaldPath: 'U:\\Prosjekt687\\programmer\\Registrering\\Roald.jar',

    keypress: function(bibsys) {

        var trace = bibsys.getTrace();
        //$.bibduck.log(trace);

        if (trace === "roald!") {
            bibsys.clearInput();
            var objShell = new ActiveXObject('WScript.Shell');
            objShell.Run(this.roaldPath, 1, false);
        }

    }

});
