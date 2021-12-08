function lettersToNumber(letters){
    var chrs = ' ABCDEFGHIJKLMNOPQRSTUVWXYZ', mode = chrs.length - 1, number = 0;
    for(var p = 0; p < letters.length; p++){
        number = number * mode + chrs.indexOf(letters[p]);
    }
    return number;
}
function numberToLetters(num) {
    let letters = ''
    while (num >= 0) {
        letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'[num % 26] + letters
        num = Math.floor(num / 26) - 1
    }
    return letters
}
var genRow = (worksheet, rowIndex) => {
    if(!worksheet['!ref'])
        return {'A':worksheet['A'+rowIndex]?.v, 'B':worksheet['B'+rowIndex]?.v, 'C':worksheet['C'+rowIndex]?.v};
    else{
        const result = worksheet['!ref'].match(/(\w\D*)(\d*):(\w\D*)(\d*)/);
        console.log('from', result[1], result[3])
        const from = lettersToNumber(result[1])
        const to = lettersToNumber(result[3])
        let obj = {};
        for(let i = from; i <= to; i++){
            obj[numberToLetters(i)] = worksheet[numberToLetters(i)+rowIndex]?.v
        }
        return obj;
    }
}