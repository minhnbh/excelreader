var Excel = require('exceljs');
var mysql = require('mysql');

var workbook = new Excel.Workbook();
var filePath = './yamaha.xlsx';
var product_codes = [];

var conn = mysql.createConnection({
    database: 'accessary_management',
    host: 'localhost',
    user: 'accessary_management',
    password: 'accessary_management'
});

conn.connect(function(err) {
    getProducts(function(products) {
        if (products && products.length > 0) {
            for (let i = 0; i < products.length; i++) {
                product_codes.push(products[i].code);
            }
        }
        if (err) throw err;
        readExcel();
    });
});

function readExcel() {
    workbook.xlsx.readFile(filePath).then(function() {
        for (let i = 0; i < 27; i++) {
            let sheet = 'sp' + (i + 1);
            let worksheet = workbook.getWorksheet(sheet);
            worksheet.eachRow({ includeEmpty: false }, function(row, rowNumber) {
                let product = [];
                let is_valid = true;
                row.eachCell({ includeEmpty: false }, function(cell, colNumber) {
                    switch (colNumber) {
                        case 2:
                            if (!product_codes.includes(cell.value) && cell.value != 'undefined') {
                                product['code'] = cell.value;
                                product['upc'] = cell.value;
                                product_codes.push(cell.value);
                            } else {
                                is_valid = false;
                            }
                            break;
                        case 1:
                            product['name'] = cell.value;
                            break;
                        case 3:
                            let price = cell.value.result;
                            if (!isNaN(price)) {
                                let format_price = Math.trunc(cell.value.result);
                                if (format_price % 10 != 0) {
                                    format_price = Math.floor(cell.value.result);
                                    if (format_price % 10 != 0) {
                                        format_price = Math.ceil(cell.value.result);   
                                    }   
                                }
                                product['import_price'] = format_price;
                                product['price'] = format_price;
                            } else {
                                product['import_price'] = 0;
                                product['price'] = 0;
                            }
                            break;
                        case 4:
                            product['model'] = cell.value;
                            break;
                    }
                });
                if (is_valid) {
                    insearchProduct(product);
                }
            });
        }
    });
}

function insearchProduct(product) {
    let sql = "INSERT INTO products (code, upc, name, import_price, price, model)\
        VALUES ('" + product.code + "', '" + product.upc + "', '" + product.name + "', '" + product.import_price + "', '" + product.price + "', '" + product.model + "')";

    conn.query(sql, function(err, results) {
        if (err) throw err;
    });
}

function getProducts(callback) {
    let sql = "SELECT * FROM products";

    conn.query(sql, function(err, results) {
        callback(results);
    });
}