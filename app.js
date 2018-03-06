/* eslint-disable no-console */

const fs = require('fs');
const path = require('path');
const config = require('./config/config');
const async = require('async');
const prefix = require('./helper/prefix_s4');
const compare = require('./helper/compare');


console.log(path.resolve(config.root));
const filePath_s4 = path.join(path.resolve(config.root),'./tables_s4');
const filePath_hana = path.join(path.resolve(config.root),'./tables_hana');
const filePath_def = path.join(path.resolve(config.root),'./tables_def');
const filePath_result = path.join(path.resolve(config.root),'./tables_result');

// fs.readdirSync(filePath_s4);

const files_s4 = fs.readdirSync(filePath_s4);
const files_hana = fs.readdirSync(filePath_hana);
const files_def = fs.readdirSync(filePath_def);
const files_result = fs.readdirSync(filePath_result);

function check_result(files_result_p,filename_result_sliced_p){
    let ok_counter = 0;
    files_result_p.forEach((filename_result) =>{
        if (filename_result.match(filename_result_sliced_p) !== null){
            ok_counter = ok_counter+1;
        }
    });
    return ok_counter;
}

async.series(
    [
        function (callback) {
            files_s4.forEach((filename_s4)=>{
                if (filename_s4.match(/\.(XLSX|xlsx)/)){
                    // console.log(filename_s4);
                    console.log(filename_s4.slice(0,-4));
                    let filename_s4_sliced = filename_s4.slice(0,-4) + 'xlsx';
                    let filename_result_sliced = filename_s4.slice(0,-5) + '_prefixed_result_OK.XLSX';
                    let check_result_count = check_result(files_result,filename_result_sliced);
                    // console.log(check_result_count);
                    if(check_result_count === 0){
                        files_def.forEach((filename_def)=>{
                            if (filename_def.match(filename_s4_sliced) !== null)
                            {    
                                console.log(filename_s4_sliced+'の定義書'+filename_def+'を見つかりました');
                                prefix(filename_s4,filename_def);
                            }
                        });
                    }else{
                        console.log(filename_s4 + 'はすで比較成功しているため対象外');
                    }
                }
            });
            callback(null, null); 
        },
        function (callback) {
            const filePath_s4_pre = path.join(path.resolve(config.root),'./tables_s4_prefix');
            const files_s4_pre = fs.readdirSync(filePath_s4_pre);
            
            files_s4_pre.forEach((filename_s4_pre)=>{
                if (filename_s4_pre.match(/\.(XLSX|xlsx)/)){
                    // console.log(filename_s4);
                    console.log(filename_s4_pre.slice(0,-14));
                    let filename_s4_pre_sliced = filename_s4_pre.slice(0,-14) + '.xlsx';
                    let filename_result_sliced = filename_s4_pre.slice(0,-14) + '_prefixed_result_OK.XLSX';
                    let check_result_count = check_result(files_result,filename_result_sliced);
                    if(check_result_count === 0){
                        files_hana.forEach((filename_hana)=>{
                            if (filename_hana.match(filename_s4_pre_sliced) !== null)
                            {
                                compare(filename_s4_pre,filename_hana);
                            }
                        });
                    }else{
                        console.log(filename_s4_pre + 'はすで比較成功しているため対象外');
                    }

                }
            });
            callback(null, null); 
        }
    ]
    // function (err, results) {
    //     // results是返回值的数组
    //     console.log('event ' + results[0] + results[1] + ' occurs');
    // }
);

// async.series(
//     [
//         ()=>{const files_s4 = fs.readdirSync(filePath_s4);
//             const files_hana = fs.readdirSync(filePath_hana);
//             console.log(files_s4);
//         },
//         // ()=>{
//         //     console.log(files_s4);
//         //     console.log(files_hana);
//         // }

//     ]
// );

