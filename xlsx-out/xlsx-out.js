var xlsx = require('json-as-xlsx');

module.exports = function (RED) {
    function XlsxOut(config) {
        RED.nodes.createNode(this, config);
        var node = this;
        node.on('input', function (msg) {
            var sheetName = config.sheetName || msg.sheetName || 'Sheet1';
            msg.payload = ToXlsx(JsonToTable(msg.payload), sheetName);
            node.send(msg);
        });
    }

    function JsonToTable(payload) {
        var columns = [];
        var content = [];
        if (Array.isArray(payload)) {
            if (payload.length > 0) {
                var c = 0;
                for (const p in payload[0]) {
                    columns.push({ label: p, value: `col${c++}` });
                }
                payload.forEach(row => {
                    var itemRow = {};
                    var c = 0;
                    for (const p in row) {
                        itemRow[`col${c++}`] = row[p];
                    }
                    content.push(itemRow);
                });
            }
        }
        return { columns, content };
    }

    function ToXlsx(table, sheetName) {


        var settings = {
            sheetName: sheetName,
            fileName: 'temp',
            extraLength: 3,
            writeOptions: {} // Style options from https://github.com/SheetJS/sheetjs#writing-options
        };

        var download = false; // If true will download the xlsx file, otherwise will return a buffer

        return xlsx(table.columns, table.content, settings, download); // Will download the excel file        
    }
    RED.nodes.registerType("xlsx-out", XlsxOut);
}