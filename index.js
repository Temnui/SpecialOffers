let XLSX = require('xlsx');
let electron = require('electron').remote;

let global = [];

let process_wb = (function () {
    let HTMLOUT = document.getElementById('htmlout');
    let XPORT = document.getElementById('xport');
    document.getElementById('beginWork').disabled = false;
    return function process_wb(wb) {
        XPORT.disabled = false;
        HTMLOUT.innerHTML = "";
        let i = 0;
        wb.SheetNames.forEach(function (sheetName) {
            global[i] = XLSX.utils.sheet_to_json(wb.Sheets[sheetName]);
            HTMLOUT.innerHTML += XLSX.utils.sheet_to_html(wb.Sheets[sheetName], {editable: true});
            i++;
        });
    };
})();

let do_file = (function () {
    return function do_file(files) {
        let f = files[0];
        let reader = new FileReader();
        reader.onload = function (e) {
            // noinspection JSUnresolvedVariable
            let data = e.target.result;
            data = new Uint8Array(data);
            process_wb(XLSX.read(data, {type: 'array'}));
        };
        reader.readAsArrayBuffer(f);
    };
})();

(function () {
    let drop = document.getElementById('drop');

    function handleDrop(e) {
        e.stopPropagation();
        e.preventDefault();
        do_file(e.dataTransfer.files);
    }

    function handleDragover(e) {
        e.stopPropagation();
        e.preventDefault();
        e.dataTransfer.dropEffect = 'copy';
    }

    drop.addEventListener('dragenter', handleDragover, false);
    drop.addEventListener('dragover', handleDragover, false);
    drop.addEventListener('drop', handleDrop, false);
})();

(function () {
    let readf = document.getElementById('readf');

    function handleF(/*e*/) {
        let o = electron.dialog.showOpenDialog({
            title: 'Select a file',
            filters: [{
                name: "Spreadsheets",
                extensions: "xls|xlsx|xlsm|xlsb|xml|xlw|xlc|csv|txt|dif|sylk|slk|prn|ods|fods|uos|dbf|wks|123|wq1|qpw|htm|html".split(
                    "|")
            }],
            properties: ['openFile']
        });
        if (o.length > 0) process_wb(XLSX.readFile(o[0]));
    }

    readf.addEventListener('click', handleF, false);
})();

(function () {
    let xlf = document.getElementById('xlf');

    function handleFile(e) {
        do_file(e.target.files);
    }

    xlf.addEventListener('change', handleFile, false);
})();

let export_xlsx = (function () {
    let HTMLOUT = document.getElementById('htmlout');
    let XTENSION = "xls|xlsx|xlsm|xlsb|xml|csv|txt|dif|sylk|slk|prn|ods|fods|htm|html".split("|");
    return function () {
        let wb = XLSX.utils.table_to_book(HTMLOUT);
        let o = electron.dialog.showSaveDialog({
            title: 'Save file as',
            filters: [{
                name: "Spreadsheets",
                extensions: XTENSION
            }]
        });
        console.log(o);
        XLSX.writeFile(wb, o);
        electron.dialog.showMessageBox({message: "Exported data to " + o, buttons: ["OK"]});
    };
})();
void export_xlsx;

let sheet = null;
let tOffers = [];
let specialOffers = {};

function beginWork() {
    // todo make book sheet navigation.
    // temporarily we chose second sheet by default
    let sheetNum = 1;
    sheet = global[sheetNum];
    console.log(sheet);
    // todo add brochure selection
    // use BROCH by default
    for (let i = 0; i < sheet.length; i++) {
        if (sheet[i]["Offer source"] === 'BROCH') {
            tOffers.push(sheet[i]);
        }
    }
    console.log(tOffers);
    for (let i = 0; i < tOffers.length; i++) {

    }
}

// section for test:

// noinspection JSUnusedLocalSymbols
function go() {
    for (let key in global[19]) {
        if (global[19].hasOwnProperty(key)) {
            if (key.toString().includes(" ", 0)) {

            }
        }
    }
}