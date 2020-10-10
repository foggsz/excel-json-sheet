/**
 * excel解析 created by fog
 */

import { utils, read, writeFile } from 'xlsx'
import { helper, ParseError } from './help.js'

// excel解析校验规则
class ExcelRules {
  constructor(
    rules,
    parseOpt = Excel.DefaultParseOpt,
    sheet = Excel.DefaultSheet
  ) {
    const defaultRules = {
      maxRows: null // 最大限制条数
    }
    this.rules = Object.assign({}, defaultRules, rules)
    this.rulesKey = this.getValidkeys()
    this.sheet = sheet
    this.parseOpt = parseOpt
  }

  changeSheet(sheet) {
    this.sheet = Object.assign({}, this.sheet, sheet)
  }

  getValidkeys() {
    let rules = this.rules
    const keys = Object.keys(rules)
    const validkeys = keys.filter(key => {
      return rules[key] != null
    })
    return validkeys || []
  }

  maxRows(maxRows) {
    // 最大行数
    // 去除表名  头部
    if (this.sheet.rows > maxRows) {
      throw new ParseError(`限制最大行${maxRows},实际最大行${this.sheet.rows}`)
    }
  }

  maxCols(maxCols) {
    // 最大列数
    if (this.sheet.cols > maxCols) {
      throw new ParseError(`限制最大列${maxCols},实际最大列${this.sheet.cols}`)
    }
  }

  headerEqualkeys(headerEqualkeys) {
    // excel第一行表头校验
    let sheetHeader = this.sheet.sheetHeader

    for (let i = 0; i < sheetHeader.length; i++) {
      if (!headerEqualkeys.includes(sheetHeader[i])) {
        throw new ParseError('excel模版不正确')
      }
    }
  }

  validate() {
    if (!this.parseOpt.onlyOneSheet) {
      // parse多个表格，规则不起效, 不做验证
      return true
    }
    for (let i = 0; i < this.rulesKey.length; i++) {
      let ruleKey = this.rulesKey[i]
      let rule = this.rules[ruleKey]
      if (ruleKey === 'validate') {
        continue
      }
      if (
        ExcelRules.prototype.hasOwnProperty(ruleKey) ||
        this.hasOwnProperty(ruleKey)
      ) {
        if (this[ruleKey] instanceof Function) {
          this[ruleKey](rule)
        }
      }
    }
    return true
  }
}

class Excel {
  /**
   * @param {object}    options  - 来源 配置数据
   * @param {function(error, data){}}    options.callback - 接收结果的处理函数
   * @param {object}    options.parseOpt -  excel文件解析配置 详见sheet.js parseOpt配置选项
   * @param {object}    options.sheetToJsonOpt  - excel数据转json配置  详见sheet.js sheet_to_json配置选项
   * @param {instance ExcelRules}  options.ruleInstance -  excel解析自定义校验规则实例
   */

  constructor(options) {
    let { callback, rules, sheetToJsonOpt, parseOpt, ruleInstance } =
      options || {}
    if (!(callback instanceof Function)) {
      throw new Error('callback must is a function')
    }
    this.callback = callback.bind(this)
    this.sheetToJsonOpt = Object.assign(
      {},
      Excel.DefaultSheetToJsonOpt,
      sheetToJsonOpt
    )
    this.parseOpt = Object.assign({}, Excel.DefaultParseOpt, parseOpt)
    this.ruleInstance =
      ruleInstance instanceof ExcelRules
        ? ruleInstance
        : new ExcelRules(rules, parseOpt) // 规则实例，可自定义扩张
    this.data = null
  }

  res(error, callback) {
    let err = error || null
    this.callback.call(
      null,
      err,
      callback instanceof Function ? callback(this.data) : this.data
    )
  }

  isAllowFileType(fileType) {
    const allowFileTypes = [
      'xlsx',
      'xls',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    ]
    return allowFileTypes.some(item => {
      return fileType === item
    })
  }

  getValidRange(sheet) {
    let keys = Object.keys(sheet)
    let start = 'A2'
    let end = false
    let sentry = '0'
    let headerStart = 'A1'
    let twoRowSentry = '0'
    let res = {
      headerRange: false,
      range: false
    }
    // console.log('keys', keys)
    for (let key in keys) {
      let item = keys[key]
      let res = helper.extractNumEnChar(item)

      if (res) {
        let enChar = res[1]
        let sz = helper.extractNum(res[0])
        if (sz >= 1) {
          if (sz === 1) {
            twoRowSentry = enChar > twoRowSentry ? enChar : twoRowSentry
          }
          if (enChar > sentry) {
            sentry = res[1]
          }

          end = res[0]
        }
      }
    }
    if (twoRowSentry === '0') {
      throw new ParseError('excel格式不正确')
    }
    res.headerRange = `${headerStart}:${twoRowSentry}1`

    if (end) {
      end = end.replace(/[a-zA-Z]+/g, sentry)
    }
    res.range = start + ':' + (end || start)
    return res
  }

  getJsonSheetRange(sheet, config) {
    let res = []
    let temp = utils.sheet_to_json(sheet, config)
    res = res.concat(temp)
    return res
  }

  /**
   * @param {file}    file  - excel文件对象
   * @returns { error||null,   list:object[] || null }
   * @returns @param {number}  list[].rows  - 行数
   * @returns @param {number}  list[].cols  - 列数
   * @returns @param {string}  list[].sheetName  - 表格名字
   * @returns @param {string}  list[].fileName  - 上传的文件名字
   * @returns @param {string[]}  list[].sheetHeader  - 表格第一行
   * @returns @param {object[]}  list[].data  - json数组数据
   */
  sheetToJson(file) {
    const self = this
    // const { type } = file
    try {
      // if (!this.isAllowFileType(type)) {
      //   throw new ParseError('请上传正确的，以xlsx/xls后缀的excel文件')
      // }

      let reader = new FileReader()

      reader.onload = function(e) {
        let result = e.target.result
        var binary = ''
        var bytes = new Uint8Array(result)
        var length = bytes.byteLength
        for (var i = 0; i < length; i++) {
          binary += String.fromCharCode(bytes[i])
        }
        let ws = read(binary, self.parseOpt)
        const sheetNames = ws.SheetNames
        let res = []
        for (let name of sheetNames) {
          try {
            let item = ws.Sheets[name]
            let validRange = self.getValidRange(item)
            let { headerRange, range } = validRange
            let orginSheet = ws.Sheets[name]

            let headerRes = self.getJsonSheetRange(orginSheet, {
              range: headerRange,
              header: 1
            })

            let sheetRes = self.getJsonSheetRange(
              orginSheet,
              Object.assign({}, { range: range }, self.sheetToJsonOpt)
            )
            let sheetData = [].concat(sheetRes)

            let rangeArr = range.split(':')
            let headerRangeArr = headerRange.split(':')
            const currentSheet = Object.assign({}, Excel.DefaultSheet)
            currentSheet.fileName = file.name
            currentSheet.data = sheetData

            currentSheet.sheetName = name
            currentSheet.sheetHeader = headerRes[0]

            if (rangeArr.length === 2 && headerRangeArr.length === 2) {
              let tempCols1 = helper
                .extractEnChar(headerRangeArr[1])
                .charCodeAt()
              let tempCols2 = helper.extractNum(rangeArr[0])
              let temColsMax = tempCols1 > tempCols2 ? tempCols1 : tempCols2
              currentSheet.cols = temColsMax - 'A'.charCodeAt() + 1
              currentSheet.rows = helper.extractNum(rangeArr[1])
            }

            self.ruleInstance.changeSheet(currentSheet)
            self.ruleInstance.validate()

            res.push(currentSheet)
            if (self.parseOpt.onlyParseOne) {
              break
            }
          } catch (error) {
            let { message } = error
            message = message || error
            return self.res(message)
          }
        }
        self.data = res
        return self.res(null)
      }
      reader.onerror = function(e) {
        return self.res('reader解析失败')
      }
      reader.readAsArrayBuffer(file)
    } catch (error) {
      self.data = null
      let { type } = error || {}
      if (type === ParseError.type) {
        return self.res(error.message)
      }
      throw error
    }
  }

  /**
   * @param {object[]}  jsonData  - 来源 配置数据
   * @param {string}    jsonData[].sheetName - 表格名
   * @param {object[]}  jsonData[].list -  待解析的json数组[]
   * @param {object}    jsonData[].config -  解析配置项
   * @param {Object}  jsonData[].config.reflectHeader -  表头映射关系 key=>名字说明
   * @param {string}    fileName  -  带后缀格式的，解析成功下载的文件名
   * @returns { error||null }
   */
  jsonToSheet(jsonData, fileName) {
    const self = this
    try {
      let fileNameArr = fileName.split('.')
      let type = fileNameArr[fileNameArr.length - 1]
      if (!this.isAllowFileType(type)) {
        throw new ParseError('文件名必须以xlsx/xls后缀')
      }
      if (!(jsonData instanceof Array) || jsonData.length === 0) {
        throw new ParseError('传入的数据格式不正确,必须为json数组')
      }

      let wb = utils.book_new()

      jsonData.map(item => {
        let { sheetName, config, list } = item
        sheetName = sheetName || ''
        config = config || {}
        let reflectHeader = config.reflectHeader || {}
        let enHeader = Object.keys(reflectHeader) // 英文头
        let cnHeader = Object.values(reflectHeader) // 中文头
        if (enHeader.length === 0) {
          throw new ParseError('映射头对象不能为空')
        }
        const headerOpts = { skipHeader: false, header: cnHeader }
        // 添加头部
        const ws = utils.json_to_sheet([], headerOpts)
        // 添加行数据
        let rowConfig = {
          skipHeader: true,
          origin: -1
        }
        for (let row of list) {
          let filterRow = {}
          for (let key of enHeader) {
            filterRow[key] = row[key] || ''
          }
          filterRow = [].concat(filterRow)
          utils.sheet_add_json(ws, filterRow, rowConfig)
        }
        utils.book_append_sheet(wb, ws, sheetName)
      })
      writeFile(wb, fileName)
      return self.res(null)
    } catch (error) {
      self.data = null
      let { type } = error || {}
      if (type === ParseError.type) {
        return self.res(error.message)
      }
      throw error
    }
  }
}

Excel.DefaultSheet = {
  cols: 0,
  rows: 0,
  data: null,
  sheetName: null,
  fileName: '',
  sheetHeader: []
}
Excel.DefaultSheetToJsonOpt = {
  header: 1
}
Excel.DefaultParseOpt = {
  sheetRows: 0,
  type: 'binary',
  raw: true,
  onlyOneSheet: true // 只解析一个表格
}
Excel.ParseError = ParseError
Excel.ExcelRules = ExcelRules
export default Excel
