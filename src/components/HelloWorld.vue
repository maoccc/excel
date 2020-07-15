<template>
  <div class="hello">
    <el-row class="upload">
      <el-upload
        class="upload-main" ref="upload" accept=".xls,.xlsx"
        action="https://jsonplaceholder.typicode.com/posts/"
        :on-change="upload"
        :show-file-list="false"
        :auto-upload="false"
        >
        <!-- 导入的excel表格和展示的表头名el-table的prop对应 -->
        <div slot="trigger">
          上传主表：
          <el-button  size="large">导入excel</el-button>
        </div>
      </el-upload>
      <el-upload
        class="upload-sub" ref="subUpload" accept=".xls,.xlsx"
        action="https://jsonplaceholder.typicode.com/posts/"
        :on-change="subUpload"
        :show-file-list="false"
        :auto-upload="false"
        >
        <!-- 导入的excel表格和展示的表头名el-table的prop对应 -->
        <div slot="trigger">
          上传副表：
          <el-button size="large">导入excel</el-button>
        </div>
      </el-upload>
    </el-row>
    <el-row  class="form-box">

      <el-form ref="form" :model="form" label-width="80px" size='mini'>
        <el-form-item label="主表字段"  style="width:500px">
          <el-input type="textarea" v-model="form.main" placeholder="例如: 姓名,身份证号"></el-input>
        </el-form-item>
        <el-form-item label="副表字段" style="width:500px">
          <el-input type="textarea" v-model="form.sub"  placeholder="例如: 学生姓名,身份证件号"></el-input>
        </el-form-item>
      </el-form>
      <div class="start" @click="start">
        开始比对
      </div>
    </el-row>
    <el-row  :gutter="20" class="table-box">
      <el-col  :span="12">
        <h3>主表：</h3>
        <el-table :data="tableData" border height="500px" size="mini">
          <el-table-column prop="姓名" label="姓名"></el-table-column>
          <el-table-column prop="班别" label="班别"></el-table-column>
          <el-table-column prop="身份证号" label="身份证号"></el-table-column>
          <el-table-column prop="学籍号" label="学籍号"></el-table-column>
        </el-table>
      </el-col>
      <el-col  :span="12">
        <h3>副表：</h3>
        <el-table :data="subTableData" border height="500px" size="mini">
          <el-table-column prop="学生姓名" label="学生姓名"></el-table-column>
          <el-table-column prop="班级" label="班级"></el-table-column>
          <el-table-column prop="身份证件号" label="身份证件号"></el-table-column>
          <el-table-column prop="学籍号" label="学籍号"></el-table-column>
        </el-table>
      </el-col>
    </el-row>
    <el-dialog title="比对结果" :visible.sync="dialogTableVisible" width="90%">
      <el-row style="text-align:right;padding-bottom:10px">
        <!-- <el-button type="success" plain>成功按钮</el-button> -->
          <el-button type="success" @click="exportBtn" size="small" >导出数据</el-button>
      </el-row>
      <el-table :data="resultData" border height="500px" size="mini">
          <el-table-column prop="姓名" label="姓名"></el-table-column>
          <el-table-column prop="班别" label="班别"></el-table-column>
          <el-table-column prop="身份证号" label="身份证号"></el-table-column>
          <el-table-column prop="学籍号" label="学籍号"></el-table-column>
          <el-table-column prop="比对结果" label="比对结果"></el-table-column>
          <el-table-column prop="错误原因" label="错误原因"></el-table-column>
        </el-table>
        <span slot="footer" class="dialog-footer">
          <el-button @click="dialogTableVisible = false">取 消</el-button>
          <el-button type="primary" @click="dialogTableVisible = false">确 定</el-button>
        </span>
    </el-dialog>
  </div>
</template>

<script>
import XLSX from 'xlsx'
export default {
  name: 'HelloWorld',
  props: {
    msg: String
  },
  data(){
    return {
      dialogTableVisible:false,
      tableData:[],
      resultData:[],
      subTableData:[],
      form:{
        main:'',
        sub:''
      }
    }
  },
  methods:{
    upload (file, fileList) {
      console.log(file, 'file')
      console.log(fileList, 'fileList')
      let files = { 0: file.raw }// 取到File
      this.readExcel(files)
    },
    subUpload (file, fileList) {
      console.log(file, 'file')
      console.log(fileList, 'fileList')
      let files = { 0: file.raw }// 取到File
      this.readSubExcel(files)
    },
    readExcel (files) { // 表格导入
      // var that = this
      console.log(files)
      if (files.length <= 0) { // 如果没有文件名
        return false
      } else if (!/\.(xls|xlsx)$/.test(files[0].name.toLowerCase())) {
        this.$Message.error('上传格式不正确，请上传xls或者xlsx格式')
        return false
      }
      const fileReader = new FileReader()
      fileReader.onload = (ev) => {
        try {
          const data = ev.target.result
          const workbook = XLSX.read(data, {
            type: 'binary'
          })
          const wsname = workbook.SheetNames[0]// 取第一张表
          const ws = XLSX.utils.sheet_to_json(workbook.Sheets[wsname])// 生成json表格内容
          console.log(ws, 'ws是表格里的数据，且是json格式')
          this.tableData = ws
          // 重写数据
          this.$refs.upload.value = ''
        } catch (e) {
          return false
        }
      }
      fileReader.readAsBinaryString(files[0])
    },
    readSubExcel (files) { // 表格导入
      // var that = this
      console.log(files)
      if (files.length <= 0) { // 如果没有文件名
        return false
      } else if (!/\.(xls|xlsx)$/.test(files[0].name.toLowerCase())) {
        this.$Message.error('上传格式不正确，请上传xls或者xlsx格式')
        return false
      }
      const fileReader = new FileReader()
      fileReader.onload = (ev) => {
        try {
          const data = ev.target.result
          const workbook = XLSX.read(data, {
            type: 'binary'
          })
          const wsname = workbook.SheetNames[0]// 取第一张表
          const ws = XLSX.utils.sheet_to_json(workbook.Sheets[wsname])// 生成json表格内容
          console.log(ws, 'ws是表格里的数据，且是json格式')
          this.subTableData = ws
          
          // 重写数据
          this.$refs.subUpload.value = ''
        } catch (e) {
          return false
        }
      }
      fileReader.readAsBinaryString(files[0])
    },
    start(){
      this.dialogTableVisible  = true
      let mainOption = this.form.main.split(',')
      let subOption = this.form.sub.split(',')
      if(mainOption.length===0 || subOption.length===0 || mainOption.length !== subOption.length){
            this.$Message('比对条件为空或者比对条件不一致，多个比对条件请用英文逗号分开，并保持主表副表条件一样多')
            return
      }
      for(var i=0 ; i < this.tableData.length;i++ ){
        let item = this.tableData[i]
        let subItem = this.subTableData.find( s =>{
          return s['学籍号'] === item['学籍号']
        })
        try{
          // console.log(subItem)
          if(!subItem){
            this.resultData.push({
              ...item,
              '比对结果':'错',
              '错误原因':'没有找到同样的学籍号'
            })

          }else {
            let result = '对'
            let cause = []
            mainOption.forEach((m,index)=>{
              let mStr = item[m] && item[m].trim()
              let sStr = subItem[subOption[index]] && subItem[subOption[index]].trim()

              if( !mStr || !sStr || mStr !== sStr ){
                result = '错'
                cause.push(m + '(' + (sStr || "") +')')
              }
            })
            this.resultData.push({
                ...item,
                '比对结果':result,
                '错误原因':cause.length > 0 ? '错误字段:' + cause.join(',') : ''
            })
          }
        }catch(err){
          console.log(err)
          console.log(item)
          console.log(subItem)
          this.resultData.push({
                ...item,
                '比对结果':'错',
                '错误原因':'未知原因'
          })
        }
      }
    },


    workbook2blob (workbook) {
    // 生成excel的配置项
      var wopts = {
        // 要生成的文件类型
        bookType: 'xlsx',
        // // 是否生成Shared String Table，官方解释是，如果开启生成速度会下降，但在低版本IOS设备上有更好的兼容性
        bookSST: false,
        type: 'binary'
      }
      var wbout = XLSX.write(workbook, wopts)
      // 将字符串转ArrayBuffer
      function s2ab (s) {
        var buf = new ArrayBuffer(s.length)
        var view = new Uint8Array(buf)
        for (var i = 0; i !== s.length; ++i) view[i] = s.charCodeAt(i) & 0xff
        return buf
      }
      let buf = s2ab(wbout)
      var blob = new Blob([buf], {
        type: 'application/octet-stream'
      })
      return blob
    },

    // 将blob对象 创建bloburl,然后用a标签实现弹出下载框
    openDownloadDialog (blob, fileName) {
      if (typeof blob === 'object' && blob instanceof Blob) {
        blob = URL.createObjectURL(blob) // 创建blob地址
      }
      var aLink = document.createElement('a')
      aLink.href = blob
      // HTML5新增的属性，指定保存文件名，可以不要后缀，注意，有时候 file:///模式下不会生效
      aLink.download = fileName || ''
      var event
      if (window.MouseEvent) event = new MouseEvent('click')
      //   移动端
      else {
        event = document.createEvent('MouseEvents')
        event.initMouseEvent('click', true, false, window, 0, 0, 0, 0, 0, false, false, false, false, 0, null)
      }
      aLink.dispatchEvent(event)
    },
    exportBtn () {
      this.exportExcel()
    },
    exportExcel () {
      let sheet1data = this.resultData
      // let sheet2data = [{ name: '张三', do: '整理文件' }, { name: '李四', do: '打印' }]
      // let sheet3data = [{ name: '王五', do: 'Vue' }, { name: '二楞', do: 'react' }]
      var sheet1 = XLSX.utils.json_to_sheet(sheet1data)
      // var sheet2 = XLSX.utils.json_to_sheet(sheet2data)
      // var sheet3 = XLSX.utils.json_to_sheet(sheet3data)
      // console.log(sheet1, sheet2, sheet3, 'sheet123')
      // 创建一个新的空的workbook
      var wb = XLSX.utils.book_new()
      XLSX.utils.book_append_sheet(wb, sheet1, '比对结果')
      // XLSX.utils.book_append_sheet(wb, sheet2, '行政部')
      // XLSX.utils.book_append_sheet(wb, sheet3, '前端部')
      const workbookBlob = this.workbook2blob(wb)
      this.openDownloadDialog(workbookBlob, '比对结果.xlsx')
    }
  }
}
</script>
<style>
.upload{
  display: flex;
  /* justify-items: ; */

}
.upload-main{
  margin-left: 100px;
}
.upload-sub{
  margin-left: 100px;
}
.table-box, .form-box{
  /* display: flex; */
  padding:20px
}
.form-box{
  display: flex;
  align-items: flex-start;
  padding-bottom: 0;

}
.start{
  width: 100px;
  height:100px;
  cursor: pointer;
  margin-left: 50px;
  text-align: center;
  display: flex;
  align-items: center;
  background-color: #ffcd0a;
  box-shadow:0px 2px 29px 1px rgba(0,0,0,0.12);;
  justify-content: center;
}
</style>