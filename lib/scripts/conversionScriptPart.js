page.evaluate(function () {
    var tables = {details: []};

    var tableEls = document.querySelectorAll("table");
    var tableHeaderTemplate = document.getElementById('excel-summary');
    if (tableHeaderTemplate) {
        var summaryTable = document.createElement('table');
        summaryTable.innerHTML = tableHeaderTemplate.innerHTML;
        tables.summaryTable = handleTable(summaryTable);
    }

    if (!tableEls[0])
        return [{rows: []}];
    for (var i = 0; i < tableEls.length; i++) {
        var table = tableEls[i];
        table.sheetName = table.getAttribute('excel-sheet-name');
        tables['details'].push(handleTable(table));

    }
    return tables;
    function handleTable(table) {
        var tableOut = {
            rows: [],
            sheetName: table.getAttribute('excel-sheet-name')
        };

        for (var r = 0, n = table.rows.length; r < n; r++) {
            var row = [];
            tableOut.rows.push(row);

            for (var c = 0, m = table.rows[r].cells.length; c < m; c++) {
                var cell = table.rows[r].cells[c];
                var cs = document.defaultView.getComputedStyle(cell, null);
                row.push({
                    value: cell.innerHTML,
                    backgroundColor: cs.getPropertyValue('background-color').match(/\d+/g),
                    foregroundColor: cs.getPropertyValue('color').match(/\d+/g),
                    fontSize: cs.getPropertyValue('font-size'),
                    fontWeight: cs.getPropertyValue('font-weight'),
                    verticalAlign: cs.getPropertyValue('vertical-align'),
                    colspan: cell.colSpan,
                    horizontalAlign: cs.getPropertyValue('text-align'),
                    width: cell.clientWidth,
                    height: cell.clientHeight,
                    format: cell.getAttribute('excel-cell-format'),
                    border: {
                        top: cs.getPropertyValue('border-top-style'),
                        right: cs.getPropertyValue('border-right-style'),
                        bottom: cs.getPropertyValue('border-bottom-style'),
                        left: cs.getPropertyValue('border-left-style'),
                        width: cs.getPropertyValue('border-width-style')
                    }
                });
            }
        }
        return tableOut;
    }
});