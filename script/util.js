month_names = ['januar','februar','mars','april','mai','juni','juli','august','september','oktober','november','desember'];

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
    href = unescape(document.location.href.substr(8).replace(/\//g, '\\')),
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
