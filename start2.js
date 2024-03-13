const ExcelJS = require('exceljs');

const readline = require('readline');


async function main() {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('assets/DATABASEE.xlsx');
    const sheet = workbook.getWorksheet("Raw Data");
    const columnCount = sheet.columnCount;

    const portfolioBook = new ExcelJS.Workbook();
    await portfolioBook.xlsx.readFile('assets/Portfolio.xlsx');
    const codesSheet = portfolioBook.getWorksheet("Codes");

    const codesMap = new Map();
    codesSheet.eachRow((row, rowNumber) => {
        codesMap.set(row.getCell("A").value, row.getCell("D").value);
    });

    sheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) {
            row.getCell("AW").value = "XIN";
            return;
        }
        const code = codesMap.get(row.getCell("M").value.toString());
        const xinValue = row.getCell("AW");
        
        if (code === undefined) {
            xinValue.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFFF0000' }
            };
            xinValue.value = "NO CODE";
            return;
        }
        xinValue.value = code;
    });


    await workbook.xlsx.writeFile('DATABASEE.xlsx');


    console.log('Done!');
}

main().catch(console.error);