# excel-merge
<p>
  <a href="https://www.npmjs.com/package/excel-merge"><img src="https://img.shields.io/npm/dm/excel-merge.svg" alt="Downloads"></a>
  <a href="https://www.npmjs.com/package/excel-merge"><img src="https://img.shields.io/npm/v/excel-merge.svg" alt="Version"></a>
  <a href="https://www.npmjs.com/package/excel-merge"><img src="https://img.shields.io/npm/l/excel-merge.svg" alt="License"></a>
</p>

## 介绍  
:rocket: nodejs合并多个excel表  

## 操作说明 
配置环境，安装node，官网下载 [node](https://nodejs.org/en/)
    
	npm install excel-merge

## 目录结构

    simple.js
    mutiple.js
    demo1文件夹/
        one.xlsx
        two.xlsx
    demo2文件夹
        one.xlsx
        two.xlsx

## 操作

simple操作

    创建simple.js，内容如下，然后执行node simple.js

```javascript
/* 合并多个excel表，只合并全部excel的第一页 (Sheet1) */
const merge = require('excel-merge');
//  文件 相对完整路径
const path = ['demo1/one.xlsx','demo1/two.xlsx','demo2/one.xlsx','demo2/two.xlsx'];
//  第一页（Sheet 1）的标识
const tag = ['姓名'];
merge.file(path, tag, 'simple');
```

```javascript
/* 合并多个excel表，只合并全部excel的第一页 （Sheet 1）*/
const merge = require('excel-merge');
//  文件夹 相对路径
const path = ['demo1/','demo2/'];
//  第一页（Sheet 1）的标识
const tag = ['姓名'];
merge.dir(path, tag, 'simple');
```

mutiple操作

    创建mutiple.js，内容如下，然后执行node mutiple.js

```javascript
/* 合并多个excel表，只合并全部excel的所有页（Sheet n），建议所有excel页数相同 */
const merge = require('excel-merge');
//  文件 相对完整路径
const path = ['demo1/one.xlsx','demo1/two.xlsx','demo2/one.xlsx','demo2/two.xlsx'];
//  如果所有页（Sheet n）的标识都相同，只需要设置一个值，否则全部都要写进来
const tag = ['姓名'];
//  const tag = [tag1, tag2, tag3, tag4];
merge.file(path, tag, 'mutiple');
```

```javascript
/* 合并多个excel表，只合并全部excel的所有页（Sheet n），建议所有excel页数相同 */
const merge = require('excel-merge');
//  文件夹 相对路径
const path = ['demo1/','demo2/'];
//  如果所有页（Sheet n）的标识都相同，只需要设置一个值，否则全部都要写进来
const tag = ['姓名'];
//  const tag = [tag1, tag2, tag3, tag4];
merge.dir(path, tag, 'mutiple');
```

## 标识说明

以下为标识，每个excel表相同页（Sheet n）都有相同的标识：'姓名'且值唯一
（不一定是'姓名'，也可以选择其他的，每页的标识都可以不同）

表一 第一页（Sheet 1）

![excel1.xlsx](https://github.com/itagn/excel-merge/raw/master/img/excel1.png)

表二 第一页（Sheet 1）

![excel2.xlsx](https://github.com/itagn/excel-merge/raw/master/img/excel2.png)

输出dest.xlsx （Sheet 1）

![dest.xlsx](https://github.com/itagn/excel-merge/raw/master/img/dest.png)

## 函数说明

    这里提供了两个函数

    .file(path, tag, type)
    path: <Array> 本地源多个excel文件的相对路径或者绝对路径组成的数组，必填
    tag: <Array> 每个excel第一页（Sheet 1）都有的标识，作为多个excel表多页合并的依据，必填
    type: <String> 'simple'和'mutiple'，默认为'simple'

    .dir(path, tag, type)
    path: <Array> 本地源多个excel文件夹组成的数组，获取文件夹下所有excel表，必填
    tag: <Array> 每个excel每页（Sheet n）都有的标识，作为多个excel表多页合并的依据，必填
        （如果每页的tag相同，则数组只设置一个值即可）
    type: <String> 'simple'和'mutiple'，默认为'simple'

## 功能说明

    simple操作
    1.合并多个excel表到一个新的excel表
    2.只合并所有excel表的第一页
    3.不影响源文件

    mutiple操作
    1.合并多个excel表到一个新的excel表
    2.只合并所有excel表的所有页
    3.不影响源文件

## 输出文档

    下载后你的文件夹会出现
    dest.xlsx: 合并多个excel文件之后的新excel文件

## 注意事项

    目前只支持xlsx文件格式
    所有excel表页数（Sheet n）相同，否则mutiple只能合并所有表格共有的前n页（Sheet n）
    每一页（Sheet n）必须有一个相同标识，标识最好具有唯一性
    标识选择excel表第一行标题中的随意一个
    excel表内容尽量不要冲突，如果发生冲突，以靠前的表为标准

作者：微博 [@itagn][1] - Github [@itagn][2] 

[1]: https://weibo.com/p/1005053782707172
[2]: https://github.com/itagn


