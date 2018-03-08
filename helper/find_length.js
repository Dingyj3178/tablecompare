// 在定义书中查找相关科目的定义
const xlsx = require('xlsx');
const utils = xlsx.utils; // XLSX.utilsのalias
// const xlsx_style = require('xlsx-style');

module.exports = (obj, query) =>{
    for (let key in obj) {
        const value_t = obj[key];
        if (value_t.v === '長さ') {
            // console.log(key);
            // console.log(utils.decode_cell(key));
            const type_col = utils.decode_cell(key).c;
            for (let key in obj) {
                const value_k = obj[key];
                if (value_k.v === query) {
                    const key_row = utils.decode_cell(key).r;
                    const addr = {c:type_col,r:key_row};
                    const addr_encode = utils.encode_cell(addr);
                    // console.log(utils.encode_cell(addr));
                    // console.log(key);
                    return obj[addr_encode].v;
                }
        
            }
        }
    }

};