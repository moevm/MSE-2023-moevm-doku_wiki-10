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
                alignmentHorizontal: cell?.alignmentHorizontal ?? "left",

                colorFont: (cell?.colorFont && cell.colorFont.length === 7) ? cell.colorFont : "#000000",
                colorBackground: (cell?.colorBackground && cell.colorBackground.length === 7) ? cell.colorBackground : "#FFFFFF"
            };
        });
    });
}

// From ODS to XLSX using the xlsx library
async function getFormattedTableFromODS(file) {
    // 1. Present the ODS file as a ZIP archive and get the content.xml file from it.    let zip = new JSZip();
    await zip.loadAsync(file);
    let tableStringXML = await zip.files["content.xml"].async('text');

    // 2. Convert XML string to JSON object.
    let tableJSON = JSON.parse(xmlToJson(tableStringXML));

    console.log(tableJSON);

    // 3. Get cell styles with the required fields as a map object.
    let styleMap = new Map(tableJSON["office:document-content"]["office:automatic-styles"]["style:style"].map(style => [style["-style:name"], {
        isBold: (style["style:text-properties"]?.["-fo:font-weight"] === "bold"),
        isItalic: (style["style:text-properties"]?.["-fo:font-style"] === "italic"),
        isUnderline: (style["style:text-properties"]?.["-style:text-underline-style"] === "solid"),
        isStrike: (style["style:text-properties"]?.["-style:text-line-through-style"] === "solid"),
        alignmentHorizontal: (style["style:paragraph-properties"]?.["-fo:text-align"] === "start") ? "left"
            : (style["style:paragraph-properties"]?.["-fo:text-align"] === "end") ? "right"
            : style["style:paragraph-properties"]?.["-fo:text-align"] ?? "left",
        colorFont: style["style:text-properties"]?.["-fo:color"] ?? "#000000",
        colorBackground: style["style:table-cell-properties"]?.["-fo:background-color"] ?? "#FFFFFF",
    }]));

    // 4. Get the cell data from the JSON object as a two-dimensional array.
    // To do this, iterate and remake from a non-uniform array
    // (infinite nesting through the "#item" field is possible) into a two-dimensional one.
    let table = [];
    let recursiveCellExportODS = function(row, exportParsedRow) {
        // a) Check the series for correctness - this is a regular series of non-zero length,
        // and no indicator at the end of the rows, showing the style at "infinity"
        // (usually row count over 10^6).
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
        // b) For each cell in the array.
        row.forEach(cell => {
            if(cell["#item"] !== undefined) {
                // c) "#item" can be any depth, so
                // recursively call the function on this object.
                recursiveCellExportODS(cell["#item"], exportParsedRow);
            } else if(
                // d) If it's a regular cell, not an indicator at the end of the array (see point a).
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
    // 5. Remove extreme empty lines.
    while((table.length > 0) && (table[table.length-1].length === 0))
        table.pop();

    // 6. Allocate the necessary data for Dokuwiki.
    let formattedTable = [];
    table.forEach((row, rowIndex) => {
        // Существует ли уже такой ряд? Если нет, то внести в массив рядов.
        // Поскольку номера рядов не уменьшаются, то считаю допустимым
        // использовать push НОВОГО пустого массива.
        if(formattedTable[rowIndex] === undefined)
            formattedTable.push(new Array(0));
        row.forEach((cell, columnIndex) => {
            if(formattedTable[rowIndex]?.[columnIndex] !== undefined) {
                // a) A cell has already been written at this position as a result of additions.
                // (see implementation below in "c").
                return;
            }
            let mergedColumns = Number(cell["-table:number-columns-spanned"]) || 1;
            let mergedRows = Number(cell["-table:number-rows-spanned"]) || 1;
            if((mergedColumns === 1) && (mergedRows === 1)) {
                // b) Single cell.
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
                // c) The remaining option is a merged cell.
                // It is necessary to write the merged cells according to the corresponding indices
                // to the right and bottom of the main cell.
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

// From XLS to XLSX using the xlsx library
async function getFormattedTableFromXLS(file) {
    /**
      There should be a function here
      turns XLS into a formatted table
      with styles. For now, there will be a challenge already
      existing function for XLSX after conversion
      to this format, with the loss of styles.
     */
    let xlsxWorkbook = XLSX.read(file);
    let xlsxRawTable = XLSX.write(xlsxWorkbook, {type: 'binary', bookType: 'xlsx'});
    return await getFormattedTableFromXLSX(xlsxRawTable);
}

// Works with the ExcelJS library
async function getFormattedTableFromXLSX(file) {
    // 1. Open the XLSX table.
    let workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(file);
    let worksheet = workbook.worksheets[0];
    let formattedTable = [];

    // 1.1. Save from the table today about colors.
    // This is needed for cases where the color was selected
    // from the suggested color themes.
    let themesXML = workbook._themes?.theme1 ?? "";
    let themesJSON = JSON.parse(xmlToJson(themesXML));
    let colorsJSON = Object.values(themesJSON
        ?.["a:theme"]
        ?.["a:themeElements"]
        ?.["a:clrScheme"] ?? {})
        .map((item) => {
            let color = item["a:srgbClr"]?.["-val"]
                ?? item["a:sysClr"]?.["-lastClr"]
                ?? undefined;
            if(!!color)
                color = "#" + color;
            return color;
        }).slice(1);

    // 1.2. A function to define a color.
    // It's easier to declare and describe it here,
    // so as not to write a lot in paragraphs 2a, 2b.
    function getColorXLSX(cell, type) {
        switch(type) {
            case "font":
                let fontStyle = cell.style.font?.color;
                if(fontStyle?.argb)
                    return "#" + (cell.style.font.color.argb.slice(2) || "000000");
                if(fontStyle?.theme !== undefined) {
                    if(fontStyle.theme === 0)
                        return "#000000";
                    return colorsJSON[fontStyle.theme];
                }
                return undefined;
            default:
            case "background":
                let fgStyle = cell.style.fill?.fgColor;
                if(fgStyle?.argb)
                    return "#" + (cell.style.fill?.fgColor?.argb?.slice(2) || "FFFFFF");
                if(fgStyle?.theme !== undefined) {
                    if(fgStyle.theme === 0)
                        return "#FFFFFF";
                    return colorsJSON[fgStyle.theme];
                }
                return undefined;
        }
    }

    // 2. Walk through all the cells in the table and extract the necessary data for Dokuwiki.
    worksheet.eachRow(function(row, rowNumber) {
        // Fill the table with empty rows
        while(formattedTable.length < rowNumber-1)
            formattedTable.push([]);
        let formattedRow = [];
        row._cells.forEach(function(cell, colNumber) {
            // Fill the line with empty cells
            while(formattedRow.length < colNumber)
                formattedRow.push({isEmpty: true});
            // Consider the cell
            let formattedCell = {};
            if(!cell.isMerged) {
                // a) If it's a regular cell with data
                formattedCell = {
                    value: cell.value ?? "",
                    isEmpty: !cell.value?.length,
                    isMerged: false,
                    isMergedFirstColumn: false,

                    isBold: cell.style.font?.bold || false,
                    isItalic: cell.style.font?.italic || false,
                    isUnderline: cell.style.font?.underline || false,
                    isStrike: cell.style.font?.strike || false,
                    alignmentHorizontal: cell.style?.alignment?.horizontal || "left",

                    colorFont: getColorXLSX(cell, "font"),
                    colorBackground: getColorXLSX(cell, "background")
                };
            } else if((cell?._mergeCount ?? 0) > 0) {
                // b) If _mergeCount > 0, then this is the main cell
                formattedCell = {
                    value: cell.value ?? "",
                    isEmpty: false,     // To distinguish main from attached on main column
                    isMerged: true,
                    isMergedFirstColumn: true,

                    isBold: cell.style.font?.bold || false,
                    isItalic: cell.style.font?.italic || false,
                    isUnderline: cell.style.font?.underline || false,
                    isStrike: cell.style.font?.strike || false,
                    alignmentHorizontal: cell.style?.alignment?.horizontal || "left",

                    colorFont: getColorXLSX(cell, "font"),
                    colorBackground: getColorXLSX(cell, "background")
                };
            } else if(formattedTable[rowNumber-2]?.[colNumber]?.isMergedFirstColumn) {
                // c) If the cell is in the main column
                formattedCell = {
                    isEmpty: true,
                    isMerged: true,
                    isMergedFirstColumn: true,
                };
            } else {
                // d) The last option - the cell is not in the main column
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

function setStyle(cell) {
    let styledCell = cell.value;

    if (cell.isBold) styledCell = `**${styledCell}**`;
    if (cell.isItalic) styledCell = `\/\/${styledCell}\/\/`;
    if (cell.isUnderline) styledCell = `__${styledCell}__`;
    if (cell.isStrike) styledCell = `<del>${styledCell}</del>`;

    if (cell.isMerged) {
        if (cell.isMergedFirstColumn) {
            if (!cell.isEmpty) return '  ' + styledCell + '  ';
            else return ':::';
        } else {
            return '';
        }
    }

    if (!cell.value && !cell.isMerged) {
        return ' ';
    }

    switch(cell.alignmentHorizontal) {
        case "left":
            styledCell = styledCell + '  ';
            break;
        case "center":
            styledCell = '  ' + styledCell + '  ';
            break;
        case "right":
            styledCell = '  ' + styledCell;
            break;
    }
    return styledCell;
}

// Output styles.
function getTextFromFormattedTable(formattedTable) {
    return formattedTable
        .map(formattedRow => {
            return "|" +
            formattedRow
                .map((cell) => setStyle(cell))
                .join("|") +
            "|";
        }).join("\n");
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
