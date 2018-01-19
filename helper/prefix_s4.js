// const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const xlsx_style = require('xlsx-style');
const utils = xlsx.utils; // XLSX.utilsのalias
const conf = require('../config/config');
const find_def_key = require('./find_def_key');

module.exports = (filename_tb,filename_def) =>{
    const filePath_tb = path.join(conf.root,'/tables_s4',filename_tb);
    const filePath_def = path.join(conf.root,'/tables_def',filename_def);
    const book_s4 = xlsx.readFile(filePath_tb);
    const book_def = xlsx.readFile(filePath_def);
    const book_s4_s = xlsx_style.readFile(filePath_tb,{cellStyles: true,cellDates:true});
    const sheetNames_s4 = book_s4.SheetNames;
    const sheetNames_def = book_def.SheetNames;
    const sheet_s4_def = book_def.Sheets[sheetNames_def[1]];
    const sheet_s4 = book_s4.Sheets[sheetNames_s4[0]];
    const decodeRange_s4 = utils.decode_range(sheet_s4['!ref']);

    for(let c_s4 = decodeRange_s4.s.c; c_s4 <= decodeRange_s4.e.c; c_s4++){
        delete book_s4_s.Sheets['Sheet1'][utils.encode_cell({c:c_s4, r:0})].s;
        if(find_def_key(sheet_s4_def,sheet_s4[utils.encode_cell({c:c_s4, r:0})].v) === conf.extractor_type.DATS){
            for(let r_s4 = decodeRange_s4.s.r+1; r_s4 <= decodeRange_s4.e.r; r_s4++){
                book_s4_s.Sheets['Sheet1'][utils.encode_cell({c:c_s4, r:r_s4})].z = 'yyyymmdd';
                delete book_s4_s.Sheets['Sheet1'][utils.encode_cell({c:c_s4, r:r_s4})].s;
                // console.log(book_s4_s.Sheets['Sheet1'][utils.encode_cell({c:c_s4, r:r_s4})]);
                
            }
        }
        else if(find_def_key(sheet_s4_def,sheet_s4[utils.encode_cell({c:c_s4, r:0})].v) === conf.extractor_type.TIMS){
            for(let r_s4 = decodeRange_s4.s.r+1; r_s4 <= decodeRange_s4.e.r; r_s4++){
                book_s4_s.Sheets['Sheet1'][utils.encode_cell({c:c_s4, r:r_s4})].z = 'hhmmss';
                delete book_s4_s.Sheets['Sheet1'][utils.encode_cell({c:c_s4, r:r_s4})].s;
            }
        }
        else{
            for(let r_s4 = decodeRange_s4.s.r+1; r_s4 <= decodeRange_s4.e.r; r_s4++){
                if(sheet_s4[utils.encode_cell({c:c_s4, r:r_s4})].v === ''){
                    book_s4_s.Sheets['Sheet1'][utils.encode_cell({c:c_s4, r:r_s4})].v = ' ';
                    delete book_s4_s.Sheets['Sheet1'][utils.encode_cell({c:c_s4, r:r_s4})].s;
                }
            }
        }
    }
    filename_tb = filename_tb.replace('.XLSX','')+'_prefixed.XLSX';
    xlsx_style.writeFile(book_s4_s, path.join(conf.root,'/tables_s4',filename_tb));
    function textify (filename_tb){
        const filePath_tb = path.join(conf.root,'/tables_s4',filename_tb);
        const filePath_def = path.join(conf.root,'/tables_def',filename_def);
        const book_s4 = xlsx.readFile(filePath_tb);
        const book_def = xlsx.readFile(filePath_def);
        const book_s4_s = xlsx_style.readFile(filePath_tb,{cellStyles: true,cellDates:true});
        const sheetNames_s4 = book_s4.SheetNames;
        const sheetNames_def = book_def.SheetNames;
        const sheet_s4_def = book_def.Sheets[sheetNames_def[1]];
        const sheet_s4 = book_s4.Sheets[sheetNames_s4[0]];
        const decodeRange_s4 = utils.decode_range(sheet_s4['!ref']);
    
        for(let c_s4 = decodeRange_s4.s.c; c_s4 <= decodeRange_s4.e.c; c_s4++){
            delete book_s4_s.Sheets['Sheet1'][utils.encode_cell({c:c_s4, r:0})].s;
            if(find_def_key(sheet_s4_def,sheet_s4[utils.encode_cell({c:c_s4, r:0})].v) === conf.extractor_type.DATS){
                for(let r_s4 = decodeRange_s4.s.r+1; r_s4 <= decodeRange_s4.e.r; r_s4++){
                    book_s4_s.Sheets['Sheet1'][utils.encode_cell({c:c_s4, r:r_s4})].v = book_s4_s.Sheets['Sheet1'][utils.encode_cell({c:c_s4, r:r_s4})].w;
                    book_s4_s.Sheets['Sheet1'][utils.encode_cell({c:c_s4, r:r_s4})].t = 's';
                    delete book_s4_s.Sheets['Sheet1'][utils.encode_cell({c:c_s4, r:r_s4})].s;
                    // console.log(book_s4_s.Sheets['Sheet1'][utils.encode_cell({c:c_s4, r:r_s4})]);
                    
                }
            }
            if(find_def_key(sheet_s4_def,sheet_s4[utils.encode_cell({c:c_s4, r:0})].v) === conf.extractor_type.TIMS){
                for(let r_s4 = decodeRange_s4.s.r+1; r_s4 <= decodeRange_s4.e.r; r_s4++){
                    book_s4_s.Sheets['Sheet1'][utils.encode_cell({c:c_s4, r:r_s4})].v = book_s4_s.Sheets['Sheet1'][utils.encode_cell({c:c_s4, r:r_s4})].w;
                    book_s4_s.Sheets['Sheet1'][utils.encode_cell({c:c_s4, r:r_s4})].t = 's';
                    delete book_s4_s.Sheets['Sheet1'][utils.encode_cell({c:c_s4, r:r_s4})].s;
                    // console.log(book_s4_s.Sheets['Sheet1'][utils.encode_cell({c:c_s4, r:r_s4})]);
                }
            }
        }
        xlsx_style.writeFile(book_s4_s, path.join(conf.root,'/tables_s4',filename_tb));
    }
    
    return textify(filename_tb);
};