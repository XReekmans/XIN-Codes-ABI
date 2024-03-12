const ExcelJS = require('exceljs');

const { load, loadCountries } = require('./loaders');
const { calculateCSize, calculatePackaging } = require('./calculators');

async function getWorkbook(name) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(name);
    return workbook;
}

function codeGenerator(items, brandsMap, subBrandsMap, cTypeMap) {
    const [brand, subBrand, cType, cSize, packaging] = items;
    const brandId = brandsMap.get(brand);
    const subBrandId = subBrandsMap.get(subBrand);
    const cTypeId = cTypeMap.get(cType);
    const formattedCSize = calculateCSize(cSize);
    const formattedPackaging = calculatePackaging(packaging);
    //console.log(`${brand}:${brandId}, ${subBrand}:${subBrandId}, ${cType}:${cTypeId}, ${cSize}:${formattedCSize}, ${packaging}:${formattedPackaging}`);
    //console.log(`${brandId}${subBrandId}${cTypeId}${formattedCSize}${formattedPackaging}`)
    return `${brandId}${subBrandId}${cTypeId}${formattedCSize}${formattedPackaging}`;
}

function iterateRows(sheet, brandsMap, subBrandsMap, cTypeMap) {
    const columnCount = sheet.columnCount;
    sheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) {
            row.getCell(columnCount).value = "XIN"
            return;
        } 
        const [brand, subBrand, cType, cSize, packaging] = [row.getCell("M").value, row.getCell("N").value, row.getCell("O").value, row.getCell("P").value, row.getCell("R").value];
        const items = [brand, subBrand, cType, cSize, packaging];
        const code = codeGenerator(items, brandsMap, subBrandsMap, cTypeMap);

        const newCell = row.getCell(columnCount);
        newCell.value = code;
    });
}

async function main() {
    const workbook = await getWorkbook('assets/Portfolio.xlsx');
    const mainSheet = workbook.getWorksheet('Export');

    const brandsSheet = workbook.getWorksheet('Brands');
    const brandsMap = load(brandsSheet);

    const subBrandsSheet = workbook.getWorksheet('Sub Brands');
    const subBrandsMap = load(subBrandsSheet);

    const CTypeSheet = workbook.getWorksheet('C Type');
    const cTypeMap = load(CTypeSheet);

    iterateRows(mainSheet, brandsMap, subBrandsMap, cTypeMap);

    const MDworkbook = await getWorkbook('assets/MD.xlsx');
    const MDSheet = MDworkbook.getWorksheet('COUNTRY MD');
    const countryMap = loadCountries(MDSheet);
    console.log(countryMap);

    await workbook.xlsx.writeFile('assets/Portfolio.xlsx');
}

main().catch(console.error);