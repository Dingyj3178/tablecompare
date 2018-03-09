// const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const xlsx_style = require('xlsx-style');
const utils = xlsx.utils; // XLSX.utilsのalias
const conf = require('../config/config');

module.exports = (filename_tb_s4,filename_tb_hana) =>{
    const filepath_tb_s4 = path.join(conf.root,'/tables_s4_prefix',filename_tb_s4);
    const filepath_tb_hana = path.join(conf.root,'/tables_hana',filename_tb_hana);
    const book_s4 = xlsx.readFile(filepath_tb_s4);
    const book_hana = xlsx.readFile(filepath_tb_hana);
    const book_s4_s = xlsx_style.readFile(filepath_tb_s4,{cellStyles: true,cellDates:true});
    const sheetNames_s4 = book_s4.SheetNames;
    const sheetNames_hana = book_hana.SheetNames;
    const sheet_hana = book_hana.Sheets[sheetNames_hana[0]];
    const sheet_s4 = book_s4.Sheets[sheetNames_s4[0]];
    const decodeRange_s4 = utils.decode_range(sheet_s4['!ref']);
    const decodeRange_hana = utils.decode_range(sheet_hana['!ref']);
    let diff_count = 0;
    let same_count = 0;
    const column_counter = decodeRange_hana.e.c - decodeRange_s4.e.c;
    if(column_counter <3){
        throw filename_tb_s4 + '項目数が異なるため、比較ファイルを確認してください';
    }


    for(let c_s4 = decodeRange_s4.s.c; c_s4 <= decodeRange_s4.e.c; c_s4++){
        if(sheet_s4[utils.encode_cell({c:c_s4, r:0})].v === sheet_hana[utils.encode_cell({c:c_s4+3, r:0})].v){
            for(let r_s4 = decodeRange_s4.s.r; r_s4 <= decodeRange_s4.e.r; r_s4++){
                if (sheet_s4[utils.encode_cell({c:c_s4, r:r_s4})].v.toString() === sheet_hana[utils.encode_cell({c:c_s4+3, r:r_s4})].v.toString()){
                    sheet_s4[utils.encode_cell({c:c_s4, r:r_s4})].s = { fill:{
                        patternType: 'solid',
                        fgColor: { rgb: '00ff72' },
                        bgColor: { indexed: 64 } }};
                    book_s4_s.Sheets['Sheet1'][utils.encode_cell({c:c_s4, r:r_s4})] = sheet_s4[utils.encode_cell({c:c_s4, r:r_s4})];
            
                }else{
                    sheet_s4[utils.encode_cell({c:c_s4, r:r_s4})].s = { fill:{
                        patternType: 'solid',
                        fgColor: { rgb: 'FF0000' },
                        bgColor: { indexed: 64 } }};
                    book_s4_s.Sheets['Sheet1'][utils.encode_cell({c:c_s4, r:r_s4})] = sheet_s4[utils.encode_cell({c:c_s4, r:r_s4})];
                    diff_count = diff_count + 1;
                }    
            }
            same_count = same_count + 1;
        }
    }
    if (decodeRange_s4.e.c - same_count > 0){
        console.log(filename_tb_s4 +'に比較されていない項目がある');
        filename_tb_s4 = filename_tb_s4.replace('.XLSX','')+'_result_NG.XLSX';
    }
    else if (diff_count > 0 ){
        console.log(filename_tb_s4 +'に一致しないデータが存在する');
        filename_tb_s4 = filename_tb_s4.replace('.XLSX','')+'_result_NG.XLSX';
    } else{
        console.log('OK！');
        filename_tb_s4 = filename_tb_s4.replace('.XLSX','')+'_result_OK.XLSX';
    }
    xlsx_style.writeFile(book_s4_s, path.join(conf.root,'/tables_result',filename_tb_s4));
    return filename_tb_s4;
};