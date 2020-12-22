<template>
  <div id="app">
    <impxspSheet
      :jsonParam="jsonParam" :blobParam="blobParam" :isFile="isFile" :isBlob="isBlob" :isJson="isJson" :isLocalFile="isLocalFile"
      @jsonResult="jsonResult" @blobResult="blobResult"
    ></impxspSheet>
  </div>
</template>

<script>
import impxspSheet from '../src/plugins/impxspsheet/impxsp-sheet'
export default {
  name: 'exspsheetDemo',
  data () {
    return {
      msg: 'exspsheet Demo',
      isLocalFile: true, //是否支持本地导入
      isFile: true, //是否导出本地文件
      isBlob: true, //是否导出blob格式
      isJson: true, //是否导出json格式
      jsonParam: '',
      blobParam: new Blob,
    }
  },
  components:{
    impxspSheet,
  },
  created(){
    this.jsonParam='[{"name":"商品销售明细表","freeze":"A1","styles":[],"merges":[],"rows":{"0":{"cells":{"0":{"text":"商品销售明细表"}}},"1":{"cells":{"0":{"text":"门店编号："},"1":{"text":"CS102003624"},"2":{"text":"门店名称："},"3":{"text":"惠多多零食店"},"5":{"text":" "},"6":{"text":"时间："},"7":{"text":"2020/01/01 - 2020/12/31"},"14":{"text":" "}}},"2":{"cells":{"0":{"text":"汇总一栏"},"1":{"text":"净销量"},"2":{"text":"净销售额(元)"},"3":{"text":"折前金额(元)"},"4":{"text":"销售成本（元）"},"5":{"text":"销售毛利（元）"},"6":{"text":"毛利率（%）"},"7":{"text":"退货数量"},"8":{"text":"退货金额（元）"},"9":{"text":"退货率（%）"}}},"3":{"cells":{"1":{"text":"45868"},"2":{"text":"226729.90 "},"3":{"text":"235092.20 "},"4":{"text":"151196.98 "},"5":{"text":"7"},"6":{"text":"33.32 "},"7":{"text":"262"},"8":{"text":"1565.90 "},"9":{"text":"0.57 "}}},"4":{"cells":{}},"5":{"cells":{"0":{"text":"条形码"},"1":{"text":"商品名称"},"3":{"text":"分类"},"4":{"text":"品牌"},"5":{"text":"净销量"},"6":{"text":"净销售额(元)"},"7":{"text":"折前金额(元)"},"8":{"text":"销售成本（元）"},"9":{"text":"销售均价（元）"},"10":{"text":"销售毛利（元）"},"11":{"text":"毛利率（%）"},"12":{"text":"毛利比重（%）"},"13":{"text":"退货数量"},"14":{"text":"退货金额（元）"},"15":{"text":"退货率（%）"}}},"6":{"cells":{"0":{"text":"2000000004754"},"1":{"text":"巧妈妈水蜜桃/橘子味果冻"},"3":{"text":"巧克力糖果茶饮类"},"4":{"text":"其他"},"5":{"text":"2286"},"6":{"text":"3552.97 "},"7":{"text":"4114.80 "},"8":{"text":"2286.00 "},"9":{"text":"1.55 "},"10":{"text":"1266.97 "},"11":{"text":"35.66 "},"12":{"text":"1.68 "},"13":{"text":"16"},"14":{"text":"25.20 "},"15":{"text":"0.70 "}}},"len":100},"cols":{"len":26},"validations":[],"autofilter":{}}]'
  },
  methods:{
    //导出json格式字符串 用于接口传递
    jsonResult(value){
      alert(value)
    },
    blobResult(value){
      console.log(value)
      this.blobParam = value;
    }
  }
}
</script>

<style lang="scss">
#app {
  font-family: 'Avenir', Helvetica, Arial, sans-serif;
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
  text-align: center;
  color: #2c3e50;
  margin-top: 60px;
}

h1, h2 {
  font-weight: normal;
}

ul {
  list-style-type: none;
  padding: 0;
}

li {
  display: inline-block;
  margin: 0 10px;
}

a {
  color: #42b983;
}
</style>
