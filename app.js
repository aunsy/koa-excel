const Excel = require('exceljs/modern.nodejs');
const Koa = require('koa');
const http = require("http");
const zlib = require('zlib');
// const koaBody = require('koa-body');
const app = new Koa();
// app.use(koaBody());

function delay(ms = 3000) {
    return new Promise((resolve, reject) => {
        setTimeout(() => {
            resolve()
        }, ms);
    })
}

// const server = http.createServer();
// server.on("request", async (req, res) => {
//     res.setHeader("Content-disposition", "attachment; filename=wahaha.xlsx");
//     res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
//     var options = {
//         stream: res,
//         useStyles: true,
//         useSharedStrings: true,
//         zip: {
//             zlib: { level: 9 } // Sets the compression level.
//         }
//     };

//     workbook = new Excel.stream.xlsx.WorkbookWriter(options);
//     // ctx.body = workbook;
//     let worksheet = workbook.addWorksheet('My Sheet');

//     // 之后调整pageSetup设置
//     worksheet.pageSetup.margins = {
//         left: 0.7, right: 0.7,
//         top: 0.75, bottom: 0.75,
//         header: 0.3, footer: 0.3
//     };

//     worksheet.columns = [
//         { header: 'Id', key: 'id', width: 10 },
//         { header: 'Name', key: 'name', width: 32 },
//         { header: 'D.O.B.', key: 'DOB', width: 10, outlineLevel: 1 }
//     ];

//     worksheet.addRow({ id: 1, name: 'John Doe', DOB: new Date(1970, 1, 1) });
//     worksheet.addRow({ id: 2, name: 'Jane Doe', DOB: new Date(1965, 1, 7) });
//     worksheet.addRow([3, 'Sam', new Date()]);
//     workbook.commit()

// })

// server.listen(3000, () => {
//     console.log("service bound at port: 3000")
// })

/********************************************************************************* */
/********************************************************************************* */
/********************************************************************************* */

// app.use(async ctx => {
//     console.log(ctx.query)
//     console.log(ctx.query.name)
//     console.log(`${ctx.query.name || "Report"}.xlsx`)
//     ctx.set('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
//     ctx.attachment(`${ctx.query.name || "Report"}.xlsx`)
//     let workbook = new Excel.Workbook();

//     workbook.creator = 'Me';
//     workbook.lastModifiedBy = 'Her';
//     workbook.created = new Date(1985, 8, 30);
//     workbook.modified = new Date();
//     workbook.lastPrinted = new Date(2016, 9, 27);

//     workbook.views = [
//         {
//             x: 0, y: 0, width: 5, height: 5,
//             firstSheet: 0, activeTab: 1, visibility: 'visible'
//         }
//     ]

//     var worksheet = workbook.addWorksheet('My Sheet');

//     // 之后调整pageSetup设置
//     worksheet.pageSetup.margins = {
//         left: 0.7, right: 0.7,
//         top: 0.75, bottom: 0.75,
//         header: 0.3, footer: 0.3
//     };

//     worksheet.columns = [
//         { header: 'Id', key: 'id', width: 10 },
//         { header: 'Name', key: 'name', width: 32 },
//         { header: 'D.O.B.', key: 'DOB', width: 10, outlineLevel: 1 }
//     ];

//     worksheet.addRow({ id: 1, name: 'John Doe', DOB: new Date(1970, 1, 1) });
//     worksheet.addRow({ id: 2, name: 'Jane Doe', DOB: new Date(1965, 1, 7) });
//     worksheet.addRow([3, 'Sam', new Date()]);

//     worksheet.getCell('B1').note = {
//         texts: [
//             { 'font': { 'size': 12, 'color': { 'theme': 0 }, 'name': 'Calibri', 'family': 2, 'scheme': 'minor' }, 'text': 'This is ' },
//             { 'font': { 'italic': true, 'size': 12, 'color': { 'theme': 0 }, 'name': 'Calibri', 'scheme': 'minor' }, 'text': 'a' },
//             { 'font': { 'size': 12, 'color': { 'theme': 1 }, 'name': 'Calibri', 'family': 2, 'scheme': 'minor' }, 'text': ' ' },
//             { 'font': { 'size': 12, 'color': { 'argb': 'FFFF6600' }, 'name': 'Calibri', 'scheme': 'minor' }, 'text': 'colorful' },
//             { 'font': { 'size': 12, 'color': { 'theme': 1 }, 'name': 'Calibri', 'family': 2, 'scheme': 'minor' }, 'text': ' text ' },
//             { 'font': { 'size': 12, 'color': { 'argb': 'FFCCFFCC' }, 'name': 'Calibri', 'scheme': 'minor' }, 'text': 'with' },
//             { 'font': { 'size': 12, 'color': { 'theme': 1 }, 'name': 'Calibri', 'family': 2, 'scheme': 'minor' }, 'text': ' in-cell ' },
//             { 'font': { 'bold': true, 'size': 12, 'color': { 'theme': 1 }, 'name': 'Calibri', 'family': 2, 'scheme': 'minor' }, 'text': 'format' },
//         ],
//     };

//     ctx.body = await workbook.xlsx.writeBuffer()
// });

// app.listen(3000);


/********************************************************************************* */
/********************************************************************************* */
/********************************************************************************* */

// app.use(async ctx => {
//     console.log(ctx.res instanceof http.ServerResponse);
//     ctx.set('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
//     // ctx.set("Content-Encoding", "gzip");
//     ctx.attachment("wahaha.xlsx")

//     // const gzip = zlib.createGzip();
//     // ctx.body = gzip;
//     var options = {
//         // stream: gzip,
//         useStyles: true,
//         useSharedStrings: true,
//         // zip: {
//         //     zlib: { level: 9 } // Sets the compression level.
//         // }
//     };

//     workbook = new Excel.stream.xlsx.WorkbookWriter(options);
//     ctx.body = workbook.stream;
//     let worksheet = workbook.addWorksheet('My Sheet');

//     // 之后调整pageSetup设置
//     worksheet.pageSetup.margins = {
//         left: 0.7, right: 0.7,
//         top: 0.75, bottom: 0.75,
//         header: 0.3, footer: 0.3
//     };

//     worksheet.columns = [
//         { header: 'Id', key: 'id', width: 10 },
//         { header: 'Name', key: 'name', width: 32 },
//         { header: 'D.O.B.', key: 'DOB', width: 10, outlineLevel: 1 }
//     ];
//     // await delay();
//     worksheet.addRow({ id: 1, name: 'John Doe', DOB: new Date(1970, 1, 1) }).commit();
//     await delay();
//     worksheet.addRow({ id: 2, name: 'Jane Doe', DOB: new Date(1965, 1, 7) }).commit();
//     await delay();
//     worksheet.addRow([3, 'Sam', new Date()]).commit();
//     await workbook.commit()

// })

// app.listen(3000);

