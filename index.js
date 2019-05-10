// Require library
const fs = require('fs');
const excel = require('excel4node');

// Create a new instance of a Workbook class
let workbook = new excel.Workbook();

// Add Worksheets to the workbook
var worksheet = workbook.addWorksheet('Jay 10');
worksheet.cell(1, 1).string('Date');

worksheet.cell(1, 2).string('FO-Paper');
worksheet.cell(1, 3).string('Start Cursor');
worksheet.cell(1, 4).string('End Cursor');


worksheet.cell(1, 6).string('FO-ELECTRONIC');
worksheet.cell(1, 7).string('Start Cursor');
worksheet.cell(1, 8).string('End Cursor');


worksheet.cell(1, 10).string('DLIR/ICA');
worksheet.cell(1, 11).string('Start Cursor');
worksheet.cell(1, 12).string('End Cursor');


worksheet.cell(1, 14).string('REMIT');
worksheet.cell(1, 15).string('Start Cursor');
worksheet.cell(1, 16).string('End Cursor');


// Read the text file and save into an array
let arr = fs.readFileSync('edl.txt').toString().split("\n");

let tableArr = [];
let dateObj = {};
arr.forEach( el => {
    let subArr = el.split("|");
    let strSource = subArr[0].trim(); strDate = subArr[1].trim();
    let strStartCursor = subArr[2].trim(); strEndCursor = subArr[3].trim();
    let strRecCount = subArr[4].trim();
    // console.log(`Source: ${strSource}, Date: ${strDate}, StarCursor: ${strStartCursor}, EndCursor: ${strEndCursor}, Count: ${strRecCount}`);
    // Check if the key exists 
    if (dateObj[strDate] == undefined){
        dateObj[strDate] = [];
    } 
    // Store data into array by source by date
    if (strSource=='DLIR/ICA'){
        dateObj[strDate].push({'DLIR/ICA' : [strStartCursor,strEndCursor,strRecCount]})
    } else if (strSource=='FO-ELECTRONIC'){
        dateObj[strDate].push({'FO-ELECTRONIC' : [strStartCursor,strEndCursor,strRecCount]})
    } else if (strSource=='FO-PAPER'){
        dateObj[strDate].push({'FO-PAPER' : [strStartCursor,strEndCursor,strRecCount]})
    } else if (strSource=='Remit'){
        dateObj[strDate].push({'REMIT' : [strStartCursor,strEndCursor,strRecCount]})
    }
})  

// console.log(dateObj)
let count=1;
for (let el in dateObj){
    // console.log(el, dateObj[el])
    count++;
    dateObj[el].forEach( source => {
        worksheet.cell(count, 1).date(el); 
        if (Object.keys(source)[0] == 'DLIR/ICA'){
            worksheet.cell(count, 10).number(+source['DLIR/ICA'][2]); 
            worksheet.cell(count, 11).number(+source['DLIR/ICA'][0]); 
            worksheet.cell(count, 12).number(+source['DLIR/ICA'][1]); 
        } else if (Object.keys(source)[0] == 'FO-ELECTRONIC'){
            worksheet.cell(count, 6).number(+source['FO-ELECTRONIC'][2]); 
            worksheet.cell(count, 7).number(+source['FO-ELECTRONIC'][0]); 
            worksheet.cell(count, 8).number(+source['FO-ELECTRONIC'][1]); 
        } else if (Object.keys(source)[0] == 'FO-PAPER'){
            worksheet.cell(count, 2).number(+source['FO-PAPER'][2]); 
            worksheet.cell(count, 3).number(+source['FO-PAPER'][0]); 
            worksheet.cell(count, 4).number(+source['FO-PAPER'][1]); 
        } else if (Object.keys(source)[0] == 'REMIT'){
            worksheet.cell(count, 14).number(+source['REMIT'][2]);
            worksheet.cell(count, 15).number(+source['REMIT'][0]);
            worksheet.cell(count, 16).number(+source['REMIT'][1]); 
        }
    })
}

// console.log(arr.length)

workbook.write('Excel.xlsx');