var Excel = require('exceljs');
var mysql = require('mysql');

var workbook = new Excel.Workbook();
var filePath = './honda.xlsx';

var conn = mysql.createConnection({
    database: 'accessary_management',
    host: 'localhost',
    user: 'accessary_management',
    password: 'accessary_management'
});

conn.connect(function(err) {
    console.log("CONNECT SUCCESSFULLY");
    if (err) throw err;
    readExcel();
});

function readExcel() {
    workbook.xlsx.readFile(filePath).then(function() {
        console.log("READ FILE SUCCESSFULLY");
        var worksheet = workbook.getWorksheet('OIL');
        worksheet.eachRow({ includeEmpty: false }, function(row, rowNumber) {
            console.log("ROW: " + rowNumber);
            if (rowNumber > 3) {
                let product = [];
                row.eachCell({ includeEmpty: false }, function(cell, colNumber) {
                    switch (colNumber) {
                        case 3:
                            product['code'] = cell.value;
                            product['upc'] = cell.value;
                            break;
                        case 6:
                            product['name'] = cell.value;
                            break;
                        case 13:
                            product['import_price'] = cell.value;
                            product['price'] = cell.value;
                            break;
                    }
                });
                insearchProduct(product);
            }
        });
    });
}

function insearchProduct(product) {
    let sql = "INSERT INTO products (code, upc, name, import_price, price)\
        VALUES ('" + product.code + "', '" + product.upc + "', '" + product.name + "', '" + product.import_price + "', '" + product.price + "')";

    conn.query(sql, function(err, results) {
        if (err) throw err;
        console.log("INSERT " + product.name + ' successfully');
    });
}