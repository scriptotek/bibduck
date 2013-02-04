
function readFile(path) {
    var forReading = 1,
        fso = new ActiveXObject('Scripting.FileSystemObject'),
        data = '';
    if (fso.FileExists(path)) {
        var file = fso.OpenTextFile(path, forReading);
        while (!file.AtEndOfStream) {
            data = data + file.ReadLine();
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