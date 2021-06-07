const qr = require('qr-image');
const Excel = require('exceljs');

async function generateQrImage(input) {
  return streamToBuffer(qr.image(input, {
    type: "png",
  }));
}

async function streamToBuffer(stream) {
  return new Promise((resolve, reject) => {
    const data = [];

    stream.on('data', (chunk) => {
      data.push(chunk);
    });

    stream.on('end', () => {
      resolve(Buffer.concat(data))
    })

    stream.on('error', (err) => {
      reject(err)
    })

  })
}



(async () => {
  try {
    const workbook = new Excel.Workbook();
    workbook.xlsx.readFile("./kod.xlsm").then(async () => {
      const worksheet = await workbook.getWorksheet(1);
      await Promise.all(worksheet.getRows(2, worksheet.rowCount - 2).map(async (row, rowIndex) => {
        const rowId = rowIndex + 1;

        const valueToQr = row.getCell("F").value?.result;

        const imageBuffer = await generateQrImage(valueToQr);
        const imageId = workbook.addImage({
          buffer: imageBuffer,
          extension: 'png',
        });

        worksheet.addImage(imageId, {
          tl: { col: 6, row:  rowId},
          ext: { width: 100, height: 100 },
        });

        row.height = 100;

      }));

      const column = worksheet.getColumn(7);

      column.width = 15;

      await workbook.xlsx.writeFile("./out.xlsx");
    });
  } catch (er) {
    console.log(er);
  }
})()
