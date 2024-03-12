
function calculateCSize(number) {
    return Math.round(number * 1000).toString().padStart(5, "0");
}
exports.calculateCSize = calculateCSize;

function calculatePackaging(value) {
    if (value === null) return 'null';
    const [firstPart, secondPart] = value.split('x');
    const formattedValue = `${firstPart.padStart(2, '0')}${secondPart.padStart(2, '0')}`;
    return formattedValue;
}
exports.calculatePackaging = calculatePackaging;
