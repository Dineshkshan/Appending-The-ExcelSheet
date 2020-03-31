var Excel = require("exceljs");
//var sheet=require("xlsx");
var workbook = new Excel.Workbook();
workbook.xlsx.readFile("excel1.xlsx")
    .then(async function () {
        var worksheet = workbook.getWorksheet('Sheet1');
        var column = worksheet.getColumn(1);
        var arr=['1','JUST A CHECK', 5064061, 'Home Depot', 9789526360];
        worksheet.spliceRows(2,0,arr);
        column.eachCell(function(cell, colNumber) {
            console.log('Cell ' + colNumber + ' = ' + cell.value);
            if(colNumber>1)
            {
                cell.value=colNumber-1;
                console.log("The value of the id is",cell.value);
            }
          });          
        //worksheet.addRow();
        //console.log("The added value is",row);
        return workbook.xlsx.writeFile('excel1.xlsx');
    });
