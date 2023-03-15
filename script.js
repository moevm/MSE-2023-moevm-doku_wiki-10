/* DOKUWIKI:include  xlsx/xlsx.mjs */
/* DOKUWIKI:include  xlsx/cpexcel.full.mjs */

XLSX.set_cptable(cptable);

function xlsx2dwButtonOnClick() {
    let input = document.createElement('input');
    input.type = 'file';
    input.accept = ".xls,.xlsx,.ods";
    input.onchange = (e) => parseTableFile(e);
    input.click();
}

function parseTableFile(e) {
    let file = e.target.files?.[0];
    if(!file)
        return;
    let reader = new FileReader();
    reader.onload = function(e) {
        let text = "";
        try {
            let workbook = XLSX.read(e.target.result);
            let sheets = workbook.Sheets;
            let sheet = Object.values(sheets)[0];
            text = getDokuWikiTableSyntaxFromSheet(sheet);
        } catch (e) {
            return;     // Something wrong
        }
        let textArea = document.getElementById('wiki__text');
        let cursorPosition = textArea.selectionStart || 0;
        textArea.value = textArea.value.slice(0, cursorPosition) + text + textArea.value.slice(cursorPosition+1);
    };
    reader.readAsArrayBuffer(file);
}

function getDokuWikiTableSyntaxFromSheet(sheet){
    const options = {
        FS: " | ",
        RS: " |\n| ",
        forceQuotes: true
    };
    let text = (("\n\n" + "| " + XLSX.utils.sheet_to_csv(sheet, options)).trim() + " |")
        .replaceAll("| \"", "| ")
        .replaceAll("\" |", " |") + "\n\n";
    // Возможная дальнейшая обработка text
    // будет выполняться в этой функции.
    return text;
}

jQuery(document).ready(() => {
    jQuery('#xlsx2dw_btn').click(() => {
        xlsx2dwButtonOnClick();
    });
});