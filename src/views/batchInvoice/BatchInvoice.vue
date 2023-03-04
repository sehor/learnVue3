<script setup>
import { ref, computed,watch} from 'vue'
import { readFile, parseInvoices, splitInvoices, dowanloadAsXml } from './batchInvoiceScript.js'
const file = ref(null)
const invoiceList = ref([])

//computeds
const totalAmount = computed(() => {
    let total= invoiceList.value.reduce((total, invoice) => {
        return total + invoice.amount * 1
    }, 0)
    if(total){
        return total.toFixed(2)
    }
    return null
})

//watchs
watch(file, (newVal, oldVal) => {
    invoiceList.value = []
})


//functions
async function analyseFile() {
    console.log(file.value[0])
    let workBook = await readFile(file.value[0])
    let invoices = parseInvoices(workBook)
    invoiceList.value = splitInvoices(invoices, 112900)
    //console.log(invoiceList)
    //dowanloadAsXml(file.value[0]).then(res=>console.log('success'))

}
function itemContent(item) {
    return `${item.itemName}  ${item.modelAndType}   ${item.quantity}  ${item.price}  ${item.amount}`
}
function fileChange(){
    invoiceList.value=[]
}
</script>
 
 
<template>
    <div class="div1">
        <h3>打开文件</h3>
        <v-file-input label="File input" placeholder="Select a file" prepend-icon="mdi-paperclip"
            v-model="file" @append="fileChange"></v-file-input>
        <h3>分析</h3>
        <h4>总金额：{{ totalAmount }}</h4>
        <v-btn @click="analyseFile">分析</v-btn>
        <div v-for="(invoice, i) in invoiceList" :key="i" class="detail">

            <div>{{ invoice.buyer.buyerName }}{{ invoice.amount }}</div>
            <!-- <div v-for="item in invoice.items">
                {{ itemContent(item) }}
            </div> -->
            <v-table theme="dark">
                <thead>
                    <tr>
                        <th>商品名称</th>
                        <th>规格型号</th>
                        <th>数量</th>
                        <th>单价</th>
                        <th>金额</th>
                    </tr>
                </thead>
                <tbody>
                    <tr v-for="item in invoice.items">
                        <td>{{ item.itemName }}</td>
                        <td>{{ item.modelAndType }}</td>
                        <td>{{ item.quantity }}</td>
                        <td>{{ item.price }}</td>
                        <td>{{ item.amount }}</td>
                    </tr>
                </tbody>
            </v-table>
        </div>
    </div>

</template>
 
 
<style lang="scss" scoped>

 /* table css */
    thead tr{
        background-color: #3c8dbc;
        color: #fff;
    }
    tbody tr{
        font: 1em sans-serif;
    }

    .detail{
        margin-bottom: 2rem;
    }

</style>