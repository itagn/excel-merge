const excel = require('node-xlsx'), fs = require('fs');
const destFile = 'dest.xlsx';

//  判断文件/文件夹是否存在
function fsExistsSync(path) {
    try{
        fs.accessSync(path, fs.F_OK);
    }catch(e){
        return false;
    }
    return true;
}

//  同步创建文件
function mkFile(path){
    return new Promise((resolve, reject)=>{
        fs.createWriteStream(path);
        resolve(true);
    })
}

//  构造对象
function Demo(head, body){
    let obj = {};
    for(let i=0;i<head.length;i++){
        obj[head[i]] = body[i];
    }
    return obj;
}

//  合并对象
function mergeObj(resource, dest){
    let arr = [];
    for(let i in dest){
        arr.push(i);
    }
    for(let i in resource){
        if(arr.indexOf(i) == -1 || dest[i] === undefined){
            dest[i] = resource[i];
        }
    }
    return dest;
}

//  获取文件夹下所有xlsx文件
function readFileList(path, filesList) {
    let files = fs.readdirSync(path);
    files.forEach(function (item, index) {
        let stat = fs.statSync(path + item);
        if (stat.isDirectory()) {
        //递归读取文件
            readFileList(path + item + "/", filesList);
        } else {
            let fileArr = item.split('.'), fileType = fileArr[fileArr.length-1];
            if(fileType === 'xlsx'){
                filesList.push(path + item);
            }
        }
    });
}
//  获取文件夹下所有的xlsx
function getXlsx(path){
    var filesList = [];
    readFileList(path, filesList);
    return filesList;
}

//  合并操作
function merge(destObj, resourceObj, tag){
    try{
        const destData = destObj.data, resourceName = resourceObj.name, resourceData = resourceObj.data;
        let data = [];
        if(resourceData.length>0){
            if(destData.length>0){
                let title = [...new Set([...destData[0], ...resourceData[0]])];
                let resourceArr = [], destArr = [], destTag = [], values = [];
                for(let i=1;i<resourceData.length;i++){
                    let demo = Demo(resourceData[0], resourceData[i]);
                    resourceArr.push(demo);
                }
                for(let i=1;i<destData.length;i++){
                    let demo = Demo(destData[0], destData[i]);
                    destArr.push(demo);
                    destTag.push(demo[tag]);
                }
                //  通过对象的方式合并同名函数
                for(let i=0;i<resourceArr.length;i++){
                    if(resourceArr[i][tag] !== undefined){
                        if(destTag.indexOf(resourceArr[i][tag]) == -1){
                            destArr.push(resourceArr[i]);
                        }else{
                            for(let j=0;j<destArr.length;j++){
                                if(resourceArr[i][tag] === destArr[j][tag]){
                                    let obj = mergeObj(resourceArr[i], destArr[j]);
                                    destArr.splice(j, 1, obj);
                                }
                            }
                        }
                    }else{
                        destArr.push(resourceArr[i]);
                    }
                }
                for(let i=0;i<destArr.length;i++){
                    let arr = [];
                    for(let j=0;j<title.length;j++){
                        arr.push(destArr[i][title[j]]);
                    }
                    values.push(arr);
                }
                data = [title, ...values];
            }else {
                data = resourceData;
            }
        }else{
            data = destData;
        }
        const result =  {
            name: resourceName,
            data: data
        };
        return result;
    }catch(err){
        console.log(err);
    }
}

//  合并递归单页操作
function mergeSimple(resource, tag){
    let destObj = excel.parse(destFile), resourceObj = excel.parse(resource);
    let buildArr = [];
    let result = merge(destObj[0], resourceObj[0], tag[0]);
    buildArr.push(result);
    let buffer = excel.build(buildArr);
    fs.writeFileSync(destFile, buffer, { 'flag': 'w' });
}

//  合并递归多页操作
function mergeMutiple(resource, tag){
    let destObj = excel.parse(destFile), resourceObj = excel.parse(resource);
    let buildArr = [];
    for(let i=0;i<resourceObj.length;i++){
        if(!destObj[i]){
            destObj[i] = {name: 'test', data: []};
        }
        let result = merge(destObj[i], resourceObj[i], tag[i] || tag[0]);
        buildArr.push(result);
    }
    let buffer = excel.build(buildArr);
    fs.writeFileSync(destFile, buffer, { 'flag': 'w' });
}

//  合并递归操作
async function excelInit(arr=[], tag, type='simple'){
    if(arr.length === 0 ){
        console.log('请输入文件');
    }else {
        const allow = arr.every(function(val){
            const fileArr = val.split('.'), fileType = fileArr[fileArr.length-1];
            if(fileArr.length>1 && fileType === 'xlsx'){
                if(fsExistsSync(val)){
                    return true;
                }else{
                    console.log('源文件不存在');
                    return false;
                }
            }else{
                console.log('只支持xlsx的文件合并');
                return false;
            }
        });
        if(allow){
            if(fsExistsSync(destFile)){
                fs.unlinkSync(destFile);
            }
            await mkFile(destFile);
            if(type === 'simple'){
                for(let i=0;i<arr.length;i++){
                    mergeSimple(arr[i], tag);
                }
                console.log('合并成功');
            }else if(type === 'mutiple'){
                for(let i=0;i<arr.length;i++){
                    mergeMutiple(arr[i], tag);
                }
                console.log('合并成功');
            }else{
                console.log('类型参数错误');
            }
        }
    }
}

let excelTool = {
    dir(path, tag, type){
        let arr = [];
        for(let i=0;i<path.length;i++){
            arr = getXlsx(path[i]);
        }
        excelInit(arr, tag, type);
    },
    file(arr, tag, type){
        excelInit(arr, tag, type);
    }
}

module.exports = excelTool;