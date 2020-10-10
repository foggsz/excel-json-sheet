<!--
 * @Author: fog
 * @Description: excel解析组件
 -->
# excel-json-sheet
excel-json-sheet组件基于sheet.js, 在其基础功能上做了二次封装, 目前支持功能如下：
- excel解析成json数据或二维数组
- excel文件解析过程的规则校验, 如限制行,列数, 限制模版头部, 校验规则可扩展
- excel解析支持异步配置
- json数组数据, 转excel文件下载
- 

## 使用示例

## Excel类实例

### 构造函数参数说明  options = { callback,  rules,  sheetToJsonOpt,  parseOpt } 解析及生成excel，可以通过直接导入Excel类方式的实现
| 参数          | 类型                | 说明 |是否必传|默认值|
| -------------- | ------------------ | --- |--|--|
| callback     | function  | 解析结果回调函数, 第一个参数为error(失败时有值),  第二参数为解析的数据(error不存在时有值)|是|无|
| rules        | object  |解析校验规则,支持的校验 maxRows:number, maxCols:number,  headerEqualkeys:Array |否|{} |
| sheetToJsonOpt | object  |  建议只对header属性配置,  https://github.com/SheetJS/sheetjs#sheet_to_json  |
| parseOpt     | object  |  建议不做配置, https://github.com/SheetJS/sheetjs#parsing-options |否 | { sheetRows: 0,  type: 'binary', raw: true,  onlyOneSheet: true } |
|ruleInstance | Excel.ExcelRules实例对象 | 可继承Excel.ExcelRules，自定义改写校验规则|否|只实例默认规则|

### excelParse方法
| 参数          | 类型                | 说明 |
| -------------- | ------------------ | --- |
| file     | 上传的excel单个文件|  上传的excel单个文件|
#### 使用示例
```javascript
    import Excel from 'excel-json-sheet'
    let config = {
        callback:(error,  list)=>{
            if(error){
              return alert(error)
            }
            let  res = list[0]
            let { fileName, data, sheetName,  sheetHeader} = res
            fileName // 上传的excel文件名
            data // json数组对象  excel第二行及以后数据
            sheetName //表格名
            sheetHeader // excel第一行数据
        }
    }
    let excel = new Excel(config)
    excel.sheetToJson(file)
```

### jsonToSheet方法
| 参数          | 类型                | 说明 |
| -------------- | ------------------ | --- |
| jsonData     | {sheetName:string,  config:object,  list:array}[] | sheetName:表格名, config.reflectHeader： 决定表格第一行填充数据,  list: 需要生成的json数组|
|fileName| string| 生成下载excel带后缀格式的文件名|

#### 使用示例
```javascript
    import Excel from 'excel-json-sheet'
    let jsonData = [{
        sheetName:'妖怪名单', 
        config:{
            reflectHeader:{name:'名字',age:'年纪' }, 
        }, 
        list:[
            {name:'犬夜叉',  age:'100000' }, 
            {name:'杀生丸',  age:'100001' }, 
        ]
    }]
    let config = {
        callback:(error)=>{
            if(error){
                return alert(error)
            }
            console.log('json转excel文件下载成功')
        }
    } 
    let excel = new Excel(config) 
    excel.jsonToSheet(jsonData,  'abc.xlsx' ) // 直接下载
    
```