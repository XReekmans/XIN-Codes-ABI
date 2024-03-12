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

function getCountryCode(countryMap, country) {
    return countryMap.get(country);
}

function iterateRows(sheet, brandsMap, subBrandsMap, cTypeMap, countryMap) {
    const columnCount = sheet.columnCount;
    sheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) {
            row.getCell(columnCount + 1).value = "XIN"
            row.getCell(columnCount + 2).value = "L-XIN"
            return;
        } 
        const [brand, subBrand, cType, cSize, packaging] = [row.getCell("M").value, row.getCell("N").value, row.getCell("O").value, row.getCell("P").value, row.getCell("R").value];
        const items = [brand, subBrand, cType, cSize, packaging];
        const code = codeGenerator(items, brandsMap, subBrandsMap, cTypeMap);

        const countryID = getCountryCode(countryMap, row.getCell("L").value);

        const lCode = `${countryID || "!COUNTRY NOT FOUND!"}${code}`;

        const xinValue = row.getCell(columnCount + 1);
        xinValue.value = code;

        const lxinValue = row.getCell(columnCount + 2);
        lxinValue.value = lCode;
    });
}

async function main() {
    const workbook = await getWorkbook('assets/Portfolio.xlsx');
    const mainSheet = workbook.getWorksheet('Export');

    const MDworkbook = await getWorkbook('assets/MD.xlsx');
    const MDSheet = MDworkbook.getWorksheet('COUNTRY MD');
    const countryMap = loadCountries(MDSheet);

    const brandsSheet = workbook.getWorksheet('Brands');
    const brandsMap = load(brandsSheet);

    const subBrandsSheet = workbook.getWorksheet('Sub Brands');
    const subBrandsMap = load(subBrandsSheet);

    const CTypeSheet = workbook.getWorksheet('C Type');
    const cTypeMap = load(CTypeSheet);

    iterateRows(mainSheet, brandsMap, subBrandsMap, cTypeMap, countryMap);

    await workbook.xlsx.writeFile('assets/Portfolio.xlsx');
}

main().catch(console.error);