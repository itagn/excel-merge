## excel-merge
:rocket: 原生nodejs合并多个excel表

    author: 蔡东
    desc: excel文档合并
    createdOn: 2017/12/18

## 操作说明 
### 配置环境，安装node，官网下载 [node](https://nodejs.org/en/)
    
	npm install excel-merge

### 操作

simple操作

    创建simple.js，内容如下，然后执行node simple.js

```javascript
/* 合并多个excel表 */
const merge = require('excel-merge');
const arr = ['one.xlsx', 'two.xlsx'], tag = '姓名';

merge(arr, tag);
```

## 功能以及注意事项说明
### 功能

    simple操作
    1.合并多个excel表到一个新的excel表
    2.不影响源文件

### 目录结构

    下载后你的文件夹会出现
    dest.xlsx: 合并多个excel文件之后的新excel文件

### 注意事项

    目前只支持单页的excel表
    必须有一个标识符号，选择excel表第一行标题中的一个
    excel表内容尽量不要冲突，标识符是唯一的存在

## 函数说明

    这里提供了一个函数，参数为(arr, tag, type)
    arr: 本地源多个excel文件的相对路径或者绝对路径组成的数组，必填
    tag: 标识符号，作为多个excel表合并的依据，必填
    type: 'simple'，默认为'simple'，可不填

    最后一个参数可以不用设置

作者：微博 [@itagn][1] - Github [@itagn][2] 

[1]: https://weibo.com/p/1005053782707172
[2]: https://github.com/itagn