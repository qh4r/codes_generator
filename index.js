const qr = require('qr-image');
const Excel = require('exceljs');
const fs = require('fs');
const path = require('path');

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
    const dirContent = fs.readdirSync('in');
    dirContent.map(file => {
      const parsedReadPath = path.join("in", file);
      const parsedWritePath = path.join("out", file);
      const workbook = new Excel.Workbook();
      workbook.xlsx.readFile(parsedReadPath).then(async () => {
        const worksheet = await workbook.getWorksheet(1);
        await Promise.all(worksheet.getRows(19, worksheet.rowCount - 18).map(async (row, rowIndex) => {
          const rowId = rowIndex + 18;

          const valueToQr = row.getCell("G").value?.result;

          const imageBuffer = await generateQrImage(valueToQr);
          const imageId = workbook.addImage({
            buffer: imageBuffer,
            extension: 'png',
          });

          worksheet.addImage(imageId, {
            tl: { col: 7, row:  rowId},
            ext: { width: 100, height: 100 },
          });

          row.height = 90;

        }));

        const column = worksheet.getColumn(7);

        column.width = 15;

        await workbook.xlsx.writeFile(parsedWritePath);
      });
    });
  } catch (er) {
    console.log(er);
  }
})()
