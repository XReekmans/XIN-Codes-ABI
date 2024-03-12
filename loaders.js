function load(sheet) {
    const brandsMap = new Map();
    sheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;
        const [, id, name] = row.values;
        brandsMap.set(name, id);
    });
    return brandsMap;
}
exports.load = load;

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

