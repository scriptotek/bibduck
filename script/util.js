month_names = ['januar','februar','mars','april','mai','juni','juli','august','september','oktober','november','desember'];

// Implementation of Object.keys for Internet Explorer 8
// Source: http://stackoverflow.com/a/3937321
Object.keys = Object.keys || (function () {
    var hasOwnProperty = Object.prototype.hasOwnProperty,
        hasDontEnumBug = !{toString:null}.propertyIsEnumerable("toString"),
        DontEnums = [ 
            'toString', 'toLocaleString', 'valueOf', 'hasOwnProperty',
            'isPrototypeOf', 'propertyIsEnumerable', 'constructor'
        ],
        DontEnumsLength = DontEnums.length;

    return function (o) {
        if (typeof o != "object" && typeof o != "function" || o === null)
            throw new TypeError("Object.keys called on a non-object");

        var result = [];
        for (var name in o) {
            if (hasOwnProperty.call(o, name))
                result.push(name);
        }

        if (hasDontEnumBug) {
            for (var i = 0; i < DontEnumsLength; i++) {
                if (hasOwnProperty.call(o, DontEnums[i]))
                    result.push(DontEnums[i]);
            }   
        }

        return result;
    };
})();

function readFile(path) {
    var forReading = 1,
        fso = new ActiveXObject('Scripting.FileSystemObject'),
        data = '';
    if (fso.FileExists(path)) {
        var file = fso.OpenTextFile(path, forReading);
        while (!file.AtEndOfStream) {
            data = data + file.ReadLine() + '\n';
        }
        file.close()
    } else {
        data = '';
    }
    return data;
}

function toSiffer(n) {
    if (n < 10) return '0' + n;
    else return String(n);
}

function treSiffer(n) {
    if (n < 10) return '00' + n;
    else if (n < 100) return '0' + n;
    else return String(n);
}

function getCurrentDir() {
    var fso = new ActiveXObject("Scripting.FileSystemObject"),
        shell = new ActiveXObject("WScript.Shell"),
        href = unescape(document.location.href.substr(5).replace(/\//g, '\\')),
        file = fso.GetFile(href),
        parentDir = file.ParentFolder + '\\';
    return parentDir;
    //folder = fso.GetFolder(parentDir),
/*
   '### To preclude problems with folder names that contain spaces, e.g. 
   '### "Documents and Settings" or "Program Files", return the 8.3 name. 
   objCurrDir = objFolder.ShortPath 
   WshShell.CurrentDirectory = objCurrDir */
}