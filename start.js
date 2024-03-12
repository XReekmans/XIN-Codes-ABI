const ExcelJS = require('exceljs');

async function getWorkbook(name) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(name);
    return workbook;
}

function loadBrands(sheet) {
    const brandsMap = new Map();
    sheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;
        const [,id, name] = row.values;
        brandsMap.set(name, id);
    });
    return brandsMap;
}

function loadSubBrands(sheet) {
    const subBrandsMap = new Map();
    sheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;
        const [,id, name] = row.values;
        subBrandsMap.set(name, id);
    });
    return subBrandsMap;
}

function loadCType(sheet) {
    const cTypeMap = new Map();
    sheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;
        const [,id, name] = row.values;
        cTypeMap.set(name, id);
    });
    return cTypeMap;
}

function calculateCSize(number) {
    return Math.round(number * 1000).toString().padStart(5, "0");
}

function calculatePackaging(value) {
    if (value === null) return 'null';
    const [firstPart, secondPart] = value.split('x');
    const formattedValue = `${firstPart.padStart(2, '0')}${secondPart.padStart(2, '0')}`;
    return formattedValue;
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

function loadCountries(sheet) {
    const countryMap = new Map();
    sheet.eachRow((row, rowNumber) => {
        if (rowNumber === 2) return;
        const [letterCode, name] = [row.getCell("K").value, row.getCell("I").value];
        countryMap.set(name, letterCode);
    });
    return countryMap;
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
    const workbook = await getWorkbook('Portfolio.xlsx');
    const mainSheet = workbook.getWorksheet('Export');

    const brandsSheet = workbook.getWorksheet('Brands');
    const brandsMap = loadBrands(brandsSheet);

    const subBrandsSheet = workbook.getWorksheet('Sub Brands');
    const subBrandsMap = loadSubBrands(subBrandsSheet);

    const CTypeSheet = workbook.getWorksheet('C Type');
    const cTypeMap = loadCType(CTypeSheet);

    iterateRows(mainSheet, brandsMap, subBrandsMap, cTypeMap);

    const MDworkbook = await getWorkbook('MD.xlsx');
    const MDSheet = MDworkbook.getWorksheet('COUNTRY MD');
    const countryMap = loadCountries(MDSheet);
    console.log(countryMap);

    await workbook.xlsx.writeFile('Portfolio.xlsx');
}

main().catch(console.error);