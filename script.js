/* DOKUWIKI:include  packages/exceljs/exceljs.js */
/* DOKUWIKI:include  packages/xlsx/xlsx.mjs */
/* DOKUWIKI:include  packages/xlsx/cpexcel.full.mjs */
/* DOKUWIKI:include  packages/jszip/jszip.js */
/* DOKUWIKI:include  packages/xmltojson/xmltojson.js */

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

function formattedTablePostroutine(table) {
    let maxCellsPerRow = Math.max(...table.map(row => row.length));
    table.forEach(row => {
        while(row.length < maxCellsPerRow)
            row.push({isEmpty: true});
    });
    table.forEach(row => {
        row.forEach((cell, index) => {
            row[index] = {
                value: cell?.value ?? "",
                isEmpty: cell?.isEmpty ?? false,
                isMerged: cell?.isMerged ?? false,
                isMergedFirstColumn: cell?.isMergedFirstColumn ?? false,

                isBold: cell?.isBold ?? false,
                isItalic: cell?.isItalic ?? false,
                isUnderline: cell?.isUnderline ?? false,
                isStrike: cell?.isStrike ?? false,
                alignmentHorizontal: cell?.alignmentHorizontal ?? "left"
            };
        });
    });
}

// Из ODS в XLSX при помощи библиотеки xlsx
async function getFormattedTableFromODS(file) {
    // 1. Представить ODS-файл как ZIP-архив и получить из него файл content.xml.
    let zip = new JSZip();
    await zip.loadAsync(file);
    let tableStringXML = await zip.files["content.xml"].async('text');

    // 2. Конвертировать строку с XML в JSON-объект.
    let tableJSON = JSON.parse(xmlToJson(tableStringXML));

    // 3. Получить стили клеток с необходимыми полями в виде map-объекта.
    let styleMap = new Map(tableJSON["office:document-content"]["office:automatic-styles"]["style:style"].map(style => [style["-style:name"], {
        isBold: (style["style:text-properties"]?.["-fo:font-weight"] === "bold"),
        isItalic: (style["style:text-properties"]?.["-fo:font-style"] === "italic"),
        isUnderline: (style["style:text-properties"]?.["-style:text-underline-style"] === "solid"),
        isStrike: (style["style:text-properties"]?.["-style:text-line-through-style"] === "solid"),
        alignmentHorizontal: (style["style:paragraph-properties"]?.["-fo:text-align"] === "start") ? "left"
            : (style["style:paragraph-properties"]?.["-fo:text-align"] === "end") ? "right"
            : style["style:paragraph-properties"]?.["-fo:text-align"] ?? "left"
    }]));

    // 4. Получить данные клеток из JSON-объекта в виде двумерного массива.
    // Для этого необходимо Пройтись и переделать из неравномерного массива
    // (возможна бесконечная вложенность через поле "#item") в двумерный.
    let table = [];
    let recursiveCellExportODS = function(row, exportParsedRow) {
        // a) Проверить ряд на корректность - это обычный ряд ненулевой длины,
        // и не индикатор в конце рядов, показывающий стиль на "бесконечности"
        // (обычно показатель строк более 10^6).
        if((Number(row["-table:number-rows-repeated"]) || 1) > 10**6)
            return;
        row = [].concat(...[row["table:covered-table-cell"], row["table:table-cell"]].map(row => {
            if(row === undefined)
                return [];
            if(!Array.isArray(row))
                return [row];
            return row;
        }));
        if(row.length === 0)
            return;
        // b) Для каждой клетки в массиве.
        row.forEach(cell => {
            if(cell["#item"] !== undefined) {
                // c) Для "#item" возможна любая глубина, поэтому
                // рекурсивно вызываем функцию для этого объекта.
                recursiveCellExportODS(cell["#item"], exportParsedRow);
            } else if(
                // d) Если это обычная клетка, а не индикатор в конце массива (см. пункт а).
                ((cell["-self-closing"] !== undefined) || (cell["-office:value-type"] !== undefined)) &&
                ((Number(cell["-table:number-columns-repeated"]) || 1) < 10**3)
            ) {
                exportParsedRow.push(cell);
            }
        });
    };
    tableJSON["office:document-content"]["office:body"]["office:spreadsheet"]["table:table"]["table:table-row"].forEach((row => {
        let parsedRow = [];
        recursiveCellExportODS(row, parsedRow);
        table.push(parsedRow);
    }));
    // 5. Удалить крайние пустые строки.
    while((table.length > 0) && (table[table.length-1].length === 0))
        table.pop();

    // 6. Выделить необходимые данные для Dokuwiki.
    let formattedTable = [];
    table.forEach((row, rowIndex) => {
        // Существует ли уже такой ряд? Если нет, то внести в массив рядов.
        // Поскольку номера рядов не уменьшаются, то считаю допустимым
        // использовать push НОВОГО пустого массива.
        if(formattedTable[rowIndex] === undefined) 
            formattedTable.push(new Array(0));
        row.forEach((cell, columnIndex) => {
            if(formattedTable[rowIndex]?.[columnIndex] !== undefined) {
                // a) На этой позиции уже записана клетка в результате добавлений.
                // (см. реализацию ниже в пункте "c").
                return;
            }
            let mergedColumns = Number(cell["-table:number-columns-spanned"]) || 1;
            let mergedRows = Number(cell["-table:number-rows-spanned"]) || 1;
            if((mergedColumns === 1) && (mergedRows === 1)) {
                // b) Одиночная клетка.
                formattedTable[rowIndex][columnIndex] = {
                    ...{
                        value: cell["text:p"] ?? "",
                        isEmpty: !cell["text:p"],
                        isMerged: false,
                        isMergedFirstColumn: false
                    }, 
                    ...(styleMap.get(cell["-table:style-name"]) ?? {})
                };
            } else {
                // c) Оставшийся вариант - объединённая клетка.
                // Необходимо записать объединённые клетки по соответствующим индексам
                // справа и снизу от главной клетки.
                for(let i = 0; i < mergedRows; i++) {
                    if(formattedTable[rowIndex+i] === undefined)
                        formattedTable.push(new Array(0));
                    for(let j = 0; j < mergedColumns; j++) {
                        formattedTable[rowIndex+i][columnIndex+j] = {
                            ...{
                                value: (i+j === 0) ? (cell["text:p"] ?? "") : "",
                                isEmpty: (i+j === 0) ? !cell["text:p"] : true,
                                isMerged: true,
                                isMergedFirstColumn: (j === 0)
                            }, 
                            ...(styleMap.get(cell["-table:style-name"]) ?? {})
                        };
                    }
                }
            }
        });
    });
    formattedTablePostroutine(formattedTable);
    return formattedTable;
}

// Из XLS в XLSX при помощи библиотеки xlsx
async function getFormattedTableFromXLS(file) {
    /*
    Здесь должна быть функция, которая
    превращает XLS в форматированную таблицу
    со стилями. Пока что здесь будет вызов уже
    существующей функции для XLSX после конвертации
    в этот формат, с потерей стилей.
    */
    let xlsxWorkbook = XLSX.read(file);
    let xlsxRawTable = XLSX.write(xlsxWorkbook, {type: 'binary', bookType: 'xlsx'});
    return await getFormattedTableFromXLSX(xlsxRawTable);
}

// Работает при помощи библиотеки ExcelJS
async function getFormattedTableFromXLSX(file) {
    // 1. Открыть XLSX-таблицу.
    let workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(file);
    let worksheet = workbook.worksheets[0];
    let formattedTable = [];

    // 2. Обойти все клетки таблицы и выделить необходимые данные для Dokuwiki.
    worksheet.eachRow(function(row, rowNumber) {
        // Дозаполнить таблицу пустыми строки
        while(formattedTable.length < rowNumber-1) 
            formattedTable.push([]);
        let formattedRow = [];
        row._cells.forEach(function(cell, colNumber) {
            // Дозаполнить строку пустыми клетками
            while(formattedRow.length < colNumber) 
                formattedRow.push({isEmpty: true});
            // Рассматриваем клетку
            let formattedCell = {};
            if(!cell.isMerged) {
                // a) Если это обычная клетка с данными
                formattedCell = {
                    value: cell.value ?? "",
                    isEmpty: !cell.value?.length,
                    isMerged: false,
                    isMergedFirstColumn: false,
    
                    isBold: cell.style.font?.bold || false,
                    isItalic: cell.style.font?.italic || false,
                    isUnderline: cell.style.font?.underline || false,
                    isStrike: cell.style.font?.strike || false,
                    alignmentHorizontal: cell.style?.alignment?.horizontal || "left"
                };
            } else if((cell?._mergeCount ?? 0) > 0) {
                // b) Если _mergeCount > 0, то это главная клетка
                formattedCell = {
                    value: cell.value ?? "",
                    isEmpty: false,     // Чтобы отличить главную от присоединённой на главном столбце
                    isMerged: true,
                    isMergedFirstColumn: true,
    
                    isBold: cell.style.font?.bold || false,
                    isItalic: cell.style.font?.italic || false,
                    isUnderline: cell.style.font?.underline || false,
                    isStrike: cell.style.font?.strike || false,
                    alignmentHorizontal: cell.style?.alignment?.horizontal || "left"
                };
            } else if(formattedTable[rowNumber-2]?.[colNumber]?.isMergedFirstColumn) {
                // c) Если клетка находится в главном столбце
                formattedCell = {
                    isEmpty: true,
                    isMerged: true,
                    isMergedFirstColumn: true,
                };
            } else {
                // d) Последний вариант - клетка находится не в главном столбце
                formattedCell = {
                    isEmpty: true,
                    isMerged: true,
                    isMergedFirstColumn: false,
                };
            }
            formattedRow.push(formattedCell);
        });
        formattedTable.push(formattedRow);
    });
    formattedTablePostroutine(formattedTable);
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
