const Excel = require("exceljs");
const child_process = require("child_process");
const fs = require("fs");

(function delExcel() {
    let files = fs.readdirSync("./")
    files.forEach(filename => {
        if (filename.match(/^(\d+-?)+\.xlsx$/g)) {
            fs.unlinkSync(filename);
        }
    })
})();

const workbook = new Excel.Workbook();

let worksheet = workbook.addWorksheet('sheet1');

const style = {
    // font: { name: 'Arial Black', color: { argb: 'FFC0000' }, bold: true },
    alignment: { vertical: 'middle', horizontal: 'center' }
};

worksheet.columns = [
    { header: '###', key: 'tag', width: 6, outlineLevel: 1, style },
    { header: 'Id', key: 'id', width: 10, style },
    { header: 'Name', key: 'name', width: 32, style },
    { header: 'D.O.B.', key: 'DOB', width: 10, style }
];
worksheet.getColumn('DOB').key = 'dob';
// worksheet.getRow(1).hidden = true;
worksheet.getRow(1).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFFF0000' }
};

for (let i = 0; i < 10; i++) {
    if (i % 3 == 0) {
        let kr = worksheet.addRow(["###"]);
    }

    let r = worksheet.addRow({ id: 1, name: 'John Doe', dob: new Date(1970, 1, 1) });

    r.outlineLevel = 1;
}

// worksheet.mergeCells('A4:B5')
// write to a file
let date = new Date();
let filename = `${date.getHours()}-${date.getMinutes()}-${date.getSeconds()}.xlsx`;

workbook.xlsx.writeFile(filename)
    .then(function () {
        console.log("done")
    });
