class ParseError extends Error {
    constructor(...rest) {
      super(rest)
      this.type = ParseError.type
      this.data = null
    }
  }
  ParseError.type = 'ParseError'
  
  const helper = {
    extractNum: function (str) {
      let reg = /[0-9]+/
      let res = reg.exec(str)
      if (res) {
        return parseInt(res[0])
      }
      return false
    },
  
    extractEnChar: function (str) {
      let reg = /[a-zA-Z]+/
      let res = reg.exec(str)
      if (res) {
        return res[0]
      }
      return false
    },
    extractNumEnChar: function (str) {
      let reg = /([a-zA-Z])[0-9]+/
      let res = reg.exec(str)
      if (res) {
        return res
      }
      return false
    },
  }
  
  export {
    helper,
    ParseError,
  }
  