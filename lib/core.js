const excel = require('node-xlsx'), fs = require('fs');
const destFile = `dest.xlsx`;

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
        if(arr.indexOf(i) == -1){
            dest[i] = resource[i];
        }
    }
    return dest;
}
//  合并具体操作
async function mergeExcel(dest, resource, tag){
    try{
        //  删除又生成文件，是为了确保excel合成后，修改源文件后继续合并失效。
        if(fsExistsSync(dest)){
            fs.unlinkSync(dest);
        }
        if(!fsExistsSync(dest)){
            await mkFile(dest);
        }
        const destObj = excel.parse(dest), resourceObj = excel.parse(resource);
        const destName = destObj[0].name, destData = destObj[0].data;
        const resourceName = resourceObj[0].name, resourceData = resourceObj[0].data;
        let data = [], values = [];

        if(destData.length>0 && resourceData.length>0){
            let title = [...new Set([...destData[0], ...resourceData[0]])];
            let resourceArr = [], destArr = [], destTag = [];
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
        let buffer = excel.build([
            {
                name: resourceName || destName,
                data: data
            }
        ]);
        fs.writeFileSync(dest, buffer, { 'flag': 'w' });
    }catch(err){
        console.log(err);
    }
}
//  递归合并
function excelInit(arr=[], tag, type='simple'){
    if(arr.length === 0 ){
        console.log('请输入文件');
    }else {
        const allow = arr.every(function(val){
            const fileArr = val.split('.'), fileType = fileArr[fileArr.length-1];
            if(fileArr.length>1 && fileType === 'xlsx'){
                if(fsExistsSync(val)){
                    return true
                }else{
                    console.log('源文件不存在');
                    return false
                }
            }else{
                console.log('只支持xlsx的文件合并');
                return false
            }
        });
        if(allow){
            if(type === 'simple'){
                mergeExcel(destFile, arr[0], tag);
                arr.shift();
                if(arr.length > 0){
                    excelInit(arr, tag, type);
                }else{
                    console.log(`结果文件保存在${destFile}`);
                }
            }else{
                console.log('类型参数错误');
            }
        }
    }
}

module.exports = excelInit;