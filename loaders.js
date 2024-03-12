function loadSubBrands(sheet) {
    const subBrandsMap = new Map();
    sheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;
        const [, id, name] = row.values;
        subBrandsMap.set(name, id);
    });
    return subBrandsMap;
}
exports.loadSubBrands = loadSubBrands;

function loadCType(sheet) {
    const cTypeMap = new Map();
    sheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;
        const [, id, name] = row.values;
        cTypeMap.set(name, id);
    });
    return cTypeMap;
}
exports.loadCType = loadCType;

function loadBrands(sheet) {
    const brandsMap = new Map();
    sheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;
        const [, id, name] = row.values;
        brandsMap.set(name, id);
    });
    return brandsMap;
}
exports.loadBrands = loadBrands;

function loadCountries(sheet) {
    const countryMap = new Map();
    sheet.eachRow((row, rowNumber) => {
        if (rowNumber === 2) return;
        const [letterCode, name] = [row.getCell("K").value, row.getCell("I").value];
        countryMap.set(name, letterCode);
    });
    return countryMap;
}
exports.loadCountries = loadCountries;

