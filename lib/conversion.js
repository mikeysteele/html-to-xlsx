var path = require("path"),
        fs = require("fs"),
        uuid = require("uuid").v1,
        tmpDir = require("os").tmpdir(),
        excelbuilder = require('msexcel-builder'),
        numeral = require('numeral');

function componentToHex(c) {
    var hex = parseInt(c).toString(16);
    return hex.length === 1 ? "0" + hex : hex;
}

function rgbToHex(c) {
    return componentToHex(c[0]) + componentToHex(c[1]) + componentToHex(c[2]);
}

function isColorDefined(c) {
    if (!c){
        return false;
    }
    return c[0] !== "0" || c[1] !== "0" || c[2] !== "0" || c[3] !== "0";
}

function getMaxLength(array) {
    var max = 0;
    array.forEach(function (a) {
        if (a.length > max)
            max = a.length;
    });
    return max;
}

function getBorderStyle(border) {
    if (border === "none")
        return undefined;

    if (border === "solid")
        return "thin";

    if (border === "double")
        return "double";

    return undefined;
}


function convert(html, cb) {
    var id = uuid();

    function icb(err, table) {
        if (err)
            return cb(err);

        tableToXlsx(table, id, cb);
    }

    if (options.strategy === "phantom-server")
        return require("./serverStrategy.js")(options, html, id, icb);
    if (options.strategy === "dedicated-process")
        return require("./dedicatedProcessStrategy.js")(options, html, id, icb);

    cb(new Error("Unsupported strategy " + options.strategy));
}

function handleCellValue(theValue, format) {
    var format = format || 'string';
    var val = {
        set: theValue || ' '
    };
    try {
        switch (format) {

            case "date":
                val.set = new Date(theValue+'Z');
                val.numberFormat = 'd-mmm-yy';
                break;
            case 'currency':
            case 'number':
                if (!theValue) {
                    val.set = null;
                } else {
                    val.set = numeral().unformat(theValue);
                }

                val.numberFormat = '#,##0 ;(#,##0)'
                break;

        }
    } catch (e) {
        console.log(e);
    }
    return val;
}


function tableToXlsx(tables, id, cb) {
    console.log(tables);
    var summary = tables.summaryTable;
    var theTables = tables.details instanceof Array ? tables.details : [tables];
    if (summary) {
        summaryRows = summary.rows;
        console.log(summaryRows);
    }
    var workbook = excelbuilder.createWorkbook(options.tmpDir, id + ".xlsx");
    for (var xx = 0; xx < theTables.length; xx++) {
        var table = theTables[xx];
        var name = table.sheetName || 'sheet' + (xx + 1);
        var maxWidths = [];
        var rows = [];
                if (summary){
                    rows = summaryRows;
                }
                rows = rows.concat(table.rows);
  
        var sheet1 = workbook.createSheet(name, getMaxLength(rows), rows.length);
        for (var i = 0; i < rows.length; i++) {
            var maxHeight = 0;
            var currentCell = 1;
            for (var j = 0; j < rows[i].length; j++) {
                var cell = rows[i][j];
                sheet1.set(currentCell, i + 1, handleCellValue(cell.value, cell.format));
                sheet1.align(currentCell, i + 1, cell.horizontalAlign);
                sheet1.valign(currentCell, i + 1, cell.verticalAlign === "middle" ? "center" : cell.verticalAlign);
                var cellWidth;
                var merge = (cell.colspan || 1) - 1;
                if (merge > 0) {
                    sheet1.merge({col: currentCell, row: i + 1}, {col: currentCell + merge, row: i + 1});
                    cellWidth = (cell.width / merge + 1) / 5;
                } else {
                    cellWidth = cell.width / 5;
                }
                for (var x = 0; x <= merge; x++) {
                    if (cell.height > maxHeight) {
                        maxHeight = cell.height;
                    }

                    if (cellWidth > (maxWidths[j] || 0)) {
                        sheet1.width(currentCell, cellWidth > 1 ? cellWidth : 1);
                        maxWidths[j] = cellWidth;
                    }
                    if (isColorDefined(cell.backgroundColor)) {
                        sheet1.fill(currentCell, i + 1, {
                            type: 'solid',
                            fgColor: 'FF' + rgbToHex(cell.backgroundColor)
                        });
                    }

                    sheet1.font(currentCell, i + 1, {
                        family: '3',
                        scheme: 'minor',
                        sz: parseInt(cell.fontSize.replace("px", "")) * 18 / 24,
                        bold: cell.fontWeight === "bold" || parseInt(cell.fontWeight, 10) >= 700,
                        //color: isColorDefined(cell.foregroundColor) ? ('FF' + rgbToHex(cell.foregroundColor)) : 'FF000FF'
                    });
                    sheet1.border(currentCell, i + 1, {
                        left: getBorderStyle(cell.border.left),
                        top: getBorderStyle(cell.border.top),
                        right: getBorderStyle(cell.border.right),
                        bottom: getBorderStyle(cell.border.bottom)
                    });
                    currentCell++;
                }

            }
            sheet1.height(currentCell, maxHeight);
        }
    }

    try {
        workbook.save(function (err) {
            if (err)
                return cb(err);

            cb(null, fs.createReadStream(path.join(options.tmpDir, id + ".xlsx")));
        });
    } catch (e) {
        e.message = JSON.stringify(e.message);
        cb(e);
    }
}



module.exports = function (opt) {
    options = opt || {};
    options.timeout = options.timeout || 10000;
    options.tmpDir = options.tmpDir || tmpDir;
    options.strategy = options.strategy || "phantom-server";

    // always set env var names for phantom-workers (don't let the user override this config)
    options.hostEnvVarName = 'PHANTOM_WORKER_HOST';
    options.portEnvVarName = 'PHANTOM_WORKER_PORT';

    convert.options = options;
    return convert;
};

