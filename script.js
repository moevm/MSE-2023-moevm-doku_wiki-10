/* DOKUWIKI:include  packages/exceljs/exceljs.js */
/* DOKUWIKI:include  packages/xlsx/xlsx.mjs */
/* DOKUWIKI:include  packages/xlsx/cpexcel.full.mjs */

XLSX.set_cptable({
    cptable,
    utils
});

function xlsx2dwButtonOnClick() {
    let input = document.createElement('input');
    input.type = 'file';
    input.accept = ".xls,.xlsx,.ods";
    input.onchange = (e) => parseTableFile(e);
    input.click();
}

function parseTableFile(e) {
    var file = e.target.files?.[0];
    if(!file)
        throw "File is undefined or unselected.";
    let fileName = file.name;

    let reader = new FileReader();
    reader.onload = async function(e) {
        let formattedTable;
        switch(fileName.slice(fileName.lastIndexOf(".")+1).toLowerCase()) {
            case "xlsx":
                formattedTable = await getFormattedTableFromXLSX(e.target.result);
                break;
            case "xls":
                formattedTable = await getFormattedTableFromXLS(e.target.result);
                break;
            case "ods":
                formattedTable = await getFormattedTableFromODS(e.target.result);
                break;
            default:
                throw "Wrong file format.";
        }
        let text = getTextFromFormattedTable(formattedTable);
        insertTextToDokuWiki(text);
    };
    reader.readAsArrayBuffer(file);
}

class Cell {
    constructor(cell) {
        switch(cell.constructor.name) {
            case "Cell":   // ExcelJS Cell Object
                this.address = cell._address || undefined;
                this.value = cell.value ?? "";

                this.fontName = cell.style.font?.name;
                this.fontSize = cell.style.font?.size;
                this.fontColor = "#" + (cell.style.font?.color?.argb?.slice(2) || "000000");

                this.isBold = cell.style.font?.bold || false;
                this.isItalic = cell.style.font?.italic || false;
                this.isUnderline = cell.style.font?.underline || false;
                this.isStrike = cell.style.font?.strike || false;

                this.alignmentHorizontal = cell.style?.alignment?.horizontal || "left";

                this.isMerged = cell.isMerged || false;
                this.mergeCount = cell._mergeCount || 0;
                this.mergedBottom = false;      // не определено
                this.mergedRight = false;       // не определено
                break;
            case "xls":     // WIP
            case "ods":     // WIP
            default:        // ???
                this.address = undefined;
                break;
        }
    }
}

// Из ODS в XLSX при помощи библиотеки xlsx
async function getFormattedTableFromODS(file) {
    /*
    Здесь будет очень большая функция, которая
    превращает ODS в форматированную таблицу
    со стилями. Пока что здесь будет вызов уже
    существующей функции для XLSX.
    */
    let xlsxWorkbook = XLSX.read(file);
    let xlsxRawTable = XLSX.write(xlsxWorkbook, {type: 'binary', bookType: 'xlsx'});
    return await getFormattedTableFromXLSX(xlsxRawTable);
}

// Из XLS в XLSX при помощи библиотеки xlsx
async function getFormattedTableFromXLS(file) {
    /*
    Здесь будет очень большая функция, которая
    превращает XLS в форматированную таблицу
    со стилями. Пока что здесь будет вызов уже
    существующей функции для XLSX.
    */
    let xlsxWorkbook = XLSX.read(file);
    let xlsxRawTable = XLSX.write(xlsxWorkbook, {type: 'binary', bookType: 'xlsx'});
    return await getFormattedTableFromXLSX(xlsxRawTable);
}

// Работает при помощи библиотеки ExcelJS
async function getFormattedTableFromXLSX(file) {
    let workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(file);
    let worksheet = workbook.worksheets[0];
    let formattedTable = [];
    worksheet.eachRow(function(row, rowNumber) {
        let formattedRow = [];
        row.eachCell(function(cell, colNumber) {
            formattedRow.push(new Cell(cell));
        });
        formattedTable.push(formattedRow);
    });
    return formattedTable;
}


// Реализовать вывод стилей в этой функции.
function getTextFromFormattedTable(formattedTable) {
    return formattedTable
        .map(formattedRow => 
            "| " +
            formattedRow
                .map(cell => cell.value)
                .join(" | ") + 
            " |"
        ).join("\n");
}

function insertTextToDokuWiki(text) {
    let textArea = jQuery('#wiki__text');
    let cursorPosition = textArea[0].selectionStart || 0;
    let sourceText = textArea.val();
    textArea.val(
        sourceText.slice(0, cursorPosition) + 
        text + 
        sourceText.slice(cursorPosition)
    );
    return;
}

jQuery(document).ready(() => {
    jQuery('#xlsx2dw_btn').click(() => {
        xlsx2dwButtonOnClick();
    });
});
