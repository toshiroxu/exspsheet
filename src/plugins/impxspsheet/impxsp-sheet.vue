<template>
    <div>
        <div class="mb-md">
            <input type="file" @change="getWorkbook" v-if="isLocalFile">
            <!-- <button @click="getJSON">获取JSON</button> -->
            <button @click="exportJSON" v-if="isJson">导出JSON</button>
            <button @click="exportExcel" v-if="isFile">导出xlsx</button>
            <button @click="exportBlob">导出blob</button>
        </div>
        <!--web spreadsheet组件-->
        <div id="x-spreadsheet-demo"></div>
    </div>

</template>

<script>
    // isLocalFile: true, //是否支持本地导入
    // isFile: true, //是否导出本地文件
    // isBlob: true, //是否导出blob格式
    // isJson: true, //是否导出json格式
    //引入依赖包
    import Spreadsheet from 'x-data-spreadsheet';
    import zhCN from 'x-data-spreadsheet/dist/locale/zh-cn';
    import XLSX from 'xlsx'
    //设置中文
    Spreadsheet.locale('zh-cn', zhCN);
    export default {
        name: "impxspsheet",
        props:{
          isLocalFile: { //是否支持本地导入
            required: false,
            type: Boolean,
            default: false,
          },
          isFile: { //是否导出本地文件
            required: false,
            type: Boolean,
            default: false,
          },
          isBlob: { //是否导出blob格式
            required: false,
            type: Boolean,
            default: false,
          },
          isJson: { //是否导出json格式
            required: false,
            type: Boolean,
            default: false,
          },
          jsonParam: {
            required: false,
            type: String,
          },
          blobParam: {
            type: Blob,
            required: false,
          },
        },
        data() {
            return {
                xs: null,
                jsondata: {
                    type: '',
                    label: ''
                },
                shuju:'',
            };
        },
        mounted() {
            this.init()
            if(this.jsonParam!==''){
              this.xs.loadData(JSON.parse(this.jsonParam));
            }
        },
        watch:{
          blobParam(newVal){
            console.log(newVal)
          }
        },
        methods: {
            init() {
                this.xs = new Spreadsheet('#x-spreadsheet-demo', {showToolbar: true, showGrid: true})
                    .loadData([{
                        styles: [
                            {
                                bgcolor: '#999999',
                                textwrap: true,
                                color: '#900b09',
                                border: {
                                    top: ['thin', '#0366d6'],
                                    bottom: ['thin', '#0366d6'],
                                    right: ['thin', '#0366d6'],
                                    left: ['thin', '#0366d6'],
                                },
                            },
                        ],
                    }]).change((cdata) => {
                        // console.log(cdata);
                        console.log('>>>', this.xs.getData());
                    });

                this.xs.on('cell-selected', (cell, ri, ci) => {
                    console.log('cell:', cell, ', ri:', ri, ', ci:', ci);
                }).on('cell-edited', (text, ri, ci) => {
                    console.log('text:', text, ', ri: ', ri, ', ci:', ci);
                });

                setTimeout(() => {
                    // xs.loadData([{ rows }]);
                    // xs.cellText(14, 3, 'cell-text').reRender();
                    // console.log('cell(8, 8):', this.xs.cell(8, 8));
                    // console.log('cellStyle(8, 8):', this.xs.cellStyle(8, 8));
                }, 5000);
            },
            loadExcelFile(fileSelected) {
                var workbook_object = this.getWorkbook(fileSelected)
                this.xs.loadData(this.stox(workbook_object));
            },
            /**
             *导出excel
             */
            exportExcel(){
                var new_wb = this.xtos(this.xs.getData());
                /* generate download */
                XLSX.writeFile(new_wb, "SheetJS.xlsx");
            },
            exportJSON(){
                //console.log(JSON.stringify(this.xs.getData()))
                this.$emit('jsonResult',JSON.stringify(this.xs.getData()));
                this.shuju = JSON.stringify(this.xs.getData())
            },
            getJSON(){
              console.log(this.shuju)
              // var blob = new Blob(["Hello World!"],{type:"text/plain;charset=utf-8"});  // type:可以设置别的文件类型
              this.xs.loadData(JSON.parse(this.shuju));
            },
            exportBlob(){
              var new_wb = this.xtos(this.xs.getData());
              var blob = new Blob([JSON.stringify(this.xs.getData())], {
                  type: 'text/plain'
              });
              console.info(blob);
              this.$emit('blobResult',blob);
              // console.info(blob.slice(1, 3, 'text/plain'));
            },
            xtos(sdata) {
                var out = XLSX.utils.book_new();
                sdata.forEach(function(xws) {
                    var aoa = [[]];
                    var rowobj = xws.rows;
                    for(var ri = 0; ri < rowobj.len; ++ri) {
                        var row = rowobj[ri];
                        if(!row) continue;
                        aoa[ri] = [];
                        Object.keys(row.cells).forEach(function(k) {
                            var idx = +k;
                            if(isNaN(idx)) return;
                            aoa[ri][idx] = row.cells[k].text;
                        });
                    }
                    var ws = XLSX.utils.aoa_to_sheet(aoa);
                    XLSX.utils.book_append_sheet(out, ws, xws.name);
                });
                return out;
            },
            stox(wb) {
                var out = [];
                wb.SheetNames.forEach(function (name) {
                    var o = {name: name, rows: {}};
                    var ws = wb.Sheets[name];
                    var aoa = XLSX.utils.sheet_to_json(ws, {raw: false, header: 1});
                    aoa.forEach(function (r, i) {
                        var cells = {};
                        r.forEach(function (c, j) {
                            cells[j] = ({text: c});
                        });
                        o.rows[i] = {cells: cells};
                    })
                    out.push(o);
                });
                return out;
            },
            /**
             * 获取文件
             * @param fileSelected
             */
            getWorkbook(fileSelected) {
                let file = fileSelected.target.files[0]
                let reader = new FileReader()
                reader.onload = e => {
                    let data = e.target.result,
                        fixedData = this.fixData(data),
                        workbook = XLSX.read(btoa(fixedData), {type: 'base64'})
                        //this.shuju = JSON.stringify(this.stox(workbook));
                    this.xs.loadData(this.stox(workbook));
                }
                reader.readAsArrayBuffer(file)
                // return workbook
            },
            fixData(data) {
                var o = "", l = 0, w = 10240
                for (; l < data.byteLength / w; ++l) o += String.fromCharCode.apply(null, new Uint8Array(data.slice(l * w, l * w + w)))
                o += String.fromCharCode.apply(null, new Uint8Array(data.slice(l * w)))
                return o
            },
        }
    }
</script>
<style scoped>
</style>
