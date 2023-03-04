<script setup >
import { ref } from 'vue'
//import {downloadFile} from '../assets/tempdownload/fileTest.js'
import ExcelJS from 'exceljs'

const props=defineProps(['msg'])

const file = ref(null)
function analyseFile() {
  //console.log(file.value)


  //to server
  // let formData = new FormData()
  // formData.append('file', file.value[0]);
  // formData.append('fileName', file.value[0].name);
  // fetch('http://localhost:9991/batchInvoice/upload', {
  //   method: 'post',
  //   body: formData
  // });

  //to workbook


  // const reader = new FileReader()
  // reader.readAsArrayBuffer(file.value[0])

  // reader.onload = (e) => {
  //   const workbook=new ExcelJS.Workbook();
  //   workbook.xlsx.load(reader.result).then(
  //     excelData=>{
  //       console.log(excelData)
  //       const worksheet=excelData.getWorksheet(1)
  //       console.log(worksheet)
  //       const data=worksheet.getSheetValues(1,1,worksheet.rowCount,worksheet.columnCount)
  //       console.log(data)
  //       const workbook2=new ExcelJS.Workbook();
  //       const worksheet2=workbook2.addWorksheet('sheet1')
  //       worksheet2.addRows(data)
  //       workbook2.xlsx.writeBuffer().then(
  //         buffer=>{
  //           const blob=new Blob([buffer],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'})
  //           //FileSaver.saveAs(blob,'test.xlsx')
  //         }
  //       )
  //     }
  //   )
  // } 

  // downloadAsXml(file.value[0])
   dowanloadAsXml(file.value[0])
}
function download() {
  downloadFile()
  }
</script>

<template>
  <div class="greetings">
    <h1 class="green">{{ msg }}</h1>
    <h3>
      You’ve successfully created a project with
      <a href="https://vitejs.dev/" target="_blank" rel="noopener">Vite</a> +
      <a href="https://vuejs.org/" target="_blank" rel="noopener">Vue 3</a>.
    </h3>
    <v-file-input label="File input" placeholder="Select a file" prepend-icon="mdi-paperclip" v-model="file">
    </v-file-input>
    <v-btn @click="analyseFile">Analyse</v-btn>
    <v-btn @click="download">download</v-btn>
  </div>
</template>

<style scoped>
h1 {
  font-weight: 500;
  font-size: 2.6rem;
  top: -10px;
}

h3 {
  font-size: 1.2rem;
}

.greetings h1,
.greetings h3 {
  text-align: center;
}

@media (min-width: 1024px) {

  .greetings h1,
  .greetings h3 {
    text-align: left;
  }
}
</style>
