// const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const xlsx_style = require('xlsx-style');
const utils = xlsx.utils; // XLSX.utilsのalias
const conf = require('../config/config');
const find_def_key = require('./find_def_key');
const async = require('async');

module.exports = (filename_tb,filename_def) =>{
    async.series(
        [         
            function(callback){
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
                            // book_s4_s.Sheets['Sheet1'][utils.encode_cell({c:c_s4, r:r_s4})].t = 's';
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
                // 
                xlsx_style.writeFile(book_s4_s, path.join(conf.root,'/tables_s4_prefix',filename_tb));

                callback(null, null); 
            },
            // 第一次执行就算改变了值得形式保存后还是无法正常显示，所以这里要重新保存后再读一遍
            function (callback){
                const filePath_tb = path.join(conf.root,'/tables_s4_prefix',filename_tb);
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
                xlsx_style.writeFile(book_s4_s, path.join(conf.root,'/tables_s4_prefix',filename_tb));
                callback(null, null); 
            },
            //从SAP导出的数据最后一列是group by的结果，需要删除
            function (callback){
                const filePath_tb = path.join(conf.root,'/tables_s4_prefix',filename_tb);
                // const book_s4 = xlsx.readFile(filePath_tb);
                let book_s4_s = xlsx_style.readFile(filePath_tb,{cellStyles: true,cellDates:true});
                let sheetNames_s4_s = book_s4_s.SheetNames;
                let sheet_s4_s = book_s4_s.Sheets[sheetNames_s4_s[0]];
                let decodeRange_s4_s = utils.decode_range(sheet_s4_s['!ref']);
                decodeRange_s4_s.e.c = decodeRange_s4_s.e.c -1;
                let obname = Object.keys(book_s4_s.Sheets['Sheet1']);
                obname.forEach((obname_test)=>{
                    if(obname_test.match(/(L)/)) {
                        delete book_s4_s.Sheets['Sheet1'][obname_test];
                    }
                });
                book_s4_s.Sheets['Sheet1']['!ref'] = utils.encode_range(decodeRange_s4_s);
                // book_s4_s.Sheets['Sheet1'][utils.decode_range(book_s4_s.Sheets['Sheet1']['!ref'])] = decodeRange_s4;
                // book_s4_s.Sheets.sheet1.forea = sheet_s4_s;
                xlsx_style.writeFile(book_s4_s, path.join(conf.root,'/tables_s4_prefix',filename_tb));
                callback(null, null); 
            }
        ]
        // function (err, results) {
        //     // results是返回值的数组
        //     console.log('event ' + results[0] + results[1] + ' occurs');
        // }
        
    );
};