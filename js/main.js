var ExcelToJSON = function() {
    this.parseExcel = function(file) {
        var reader = new FileReader();

        reader.onload = function(e) {
            var data = e.target.result;
            var workbook = XLSX.read(data, {
                type: 'binary'
            });
            
            workbook.SheetNames.forEach(function(sheetName) {
                if (sheetName.includes('Ausdrücke')) {
                    var csv_object = XLSX.utils.sheet_to_csv(workbook.Sheets[sheetName], {FS: ';'});
                    download(sheetName + ".exp", csv_object);
                }
            })
        };

        reader.onerror = function(ex) {
            console.error(ex);
        };

        reader.readAsBinaryString(file);
    };
};

function download(filename, text) {
    var element = document.createElement('a');
    element.setAttribute('href', 'data:text/plain;charset=utf-8,' + encodeURIComponent(text));
    element.setAttribute('download', filename);
  
    element.style.display = 'none';
    document.body.appendChild(element);
  
    element.click();
  
    document.body.removeChild(element);
}  

function handleFileSelect(evt) {
    var files = evt.target.files; // FileList object
    var xl2json = new ExcelToJSON();
    xl2json.parseExcel(files[0]);
}