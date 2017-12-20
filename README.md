## excel-merge
:rocket: nodejs合并多个excel表

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
/* 合并多个excel表，只合并全部excel的某一页 */

//  只合并所有excel表的第二页
const merge = require('excel-merge');
const arr = ['one.xlsx', 'two.xlsx']; 
const tag = ['姓名', '工号'];
const num = 2;

merge(arr, tag, 'simple', 2);
```

mutiple操作

    创建mutiple.js，内容如下，然后执行node mutiple.js

```javascript
/* 合并多个excel表，合并所有页 */

//  合并所有excel表的所有页
const merge = require('excel-merge');
const arr = ['one.xlsx', 'two.xlsx']; 
const tag = ['姓名', '工号'];

merge(arr, tag, 'mutiple');
```

以下为标识，不一定是'姓名'，也可以选择其他的

表一，标识设置为'姓名'

![标识1](https://github.com/itagn/excel-merge/blob/master/title1.png)

表二，标识设置为'学号'

![标识2](https://github.com/itagn/excel-merge/blob/master/title2.png)

## 函数说明

    这里提供了一个函数，参数为(arr, tag, type, num)
    arr: <Array> 本地源多个excel文件的相对路径或者绝对路径组成的数组，必填
    tag: <Array> 每页excel标识符号组成的数组，作为多个excel表多页合并的依据，必填
    type: <String> 'simple'和'mutiple'，默认为'simple'，为'simple'时需要带最后一个参数
    num: <Number> 作为'simple'操作的协助参数，只合并某一页由它决定

## 功能以及注意事项说明
### 功能

    simple操作
    1.合并多个excel表到一个新的excel表
    2.只合并所有excel表的具体一页
    2.不影响源文件

    mutiple操作
    1.合并多个excel表到一个新的excel表
    2.合并所有excel表的所有页
    2.不影响源文件

### 目录结构

    下载后你的文件夹会出现
    dest.xlsx: 合并多个excel文件之后的新excel文件

### 注意事项

    目前只支持xlsx表
    每一页必须有一个标识，标识最好具有唯一性
    标识选择excel表第一行标题中的随意一个
    excel表内容尽量不要冲突，如果发生冲突，以靠前的表为标准

作者：微博 [@itagn][1] - Github [@itagn][2] 

[1]: https://weibo.com/p/1005053782707172
[2]: https://github.com/itagn