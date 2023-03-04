import ExcelJS from 'exceljs'
import {saveAs} from 'file-saver'

const headerToKeyMap = new Map([
    ["购方名称", "buyerName"],
    ["购方税号", "buyerTaxNum"],
    ["购方银行账号", "buyerBankInfo"],
    ["购方地址电话", "buyerAddrAndPhone"],
    ["备注", "notes"],
    ["复核人", "reviewer"],
    ["收款人", "payee"],
    ["商品编码版本号", "codeVersion"],
    ["含税标志", "taxFlag"],
    ["商品名称", "itemName"],
    ["规格型号", "modelAndType"],
    ["计量单位", "unitType"],
    ["商品编码", "itemTaxNum"],
    ["企业商品编码", "Qyspbm"],
    ["优惠政策标识", "Syyhzcbz"],
    ["零税率标识", "Lslbz"],
    ["优惠政策说明", "Yhzcsm"],
    ["单价", "price"],
    ["数量", "quantity"],
    ["金额", "amount"],
    ["税率", "taxRate"],
    ["扣除额", "Kce"]
])

function keyToColumnIndex(headerToKeyMap, row) {
    let map = new Map();
    row.eachCell((cell, colNumber) => {
        let key = headerToKeyMap.get(cell.value)
        if (key) {
            map.set(key, colNumber)
        }
    })
    //console.log({coloums});
    return map
}

//reutrn a promise workBook
export function readFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader()
        reader.readAsArrayBuffer(file)

        reader.onload = () => {
            new ExcelJS.Workbook().xlsx.load(reader.result).then((workbook) => {
                //console.log(`before resolve${workbook}`)
                resolve(workbook)
            }).catch((err) => {
                reject(err)
            })
        }

    })

}

// return a invoice list
export function parseInvoices(workbook) {
   // console.log('begin parese...')
    let sheet = workbook.getWorksheet(1)
    let keyToColMap = keyToColumnIndex(headerToKeyMap, sheet.getRow(1))
    //console.log(keyToColMap)
    //set coloums
    let data = [];
    sheet.eachRow({
        includeEmpty: false
    }, function (row, rowNumber) {
        if (rowNumber == 1) {
            return;
        }
        let item = {};
        for (let [k, v] of keyToColMap) {
            let cell = row.getCell(v);
            item[k] = cell.result ? row.getCell(v).result : row.getCell(v).value;
        }
        data.push(item);
    });
    //invioceItems to   invoices 
    let invoices = [];
    let invoice = undefined;
    for (let i = 0; i < data.length; i++) {
        let row = data[i];
        let invoiceItem = {
            itemName: row.itemName,
            modelAndType: row.modelAndType,
            unitType: row.unitType,
            itemTaxNum: row.itemTaxNum,
            Qyspbm: row.Qyspbm,
            Syyhzcbz: row.Syyhzcbz,
            Lslbz: row.Lslbz,
            Yhzcsm: row.Yhzcsm,
            price: row.price,
            quantity: row.quantity,
            amount: row.amount,
            taxRate: row.taxRate,
            Kce: row.Kce,
        }
        if (row.buyerName) {
            //save pre invoice
           if(invoice) {invoices.push(Object.assign({}, invoice));}
           //start a new invoice
            invoice = {items:[]};
            let buyer = {
                buyerName: row.buyerName,
                buyerTaxNum: row.buyerTaxNum,
                buyerBankInfo: row.buyerBankInfo,
                buyerAddrAndPhone: row.buyerAddrAndPhone,
            }
            invoice.buyer = buyer;
            invoice.notes = row.notes;
            invoice.reviewer = row.reviewer;
            invoice.payee = row.payee;
            invoice.codeVersion = row.codeVersion;
            invoice.taxFlag = row.taxFlag;
        }
        invoice.items.push(invoiceItem);
 
    }
    //save last invoice
    if(invoice) {invoices.push(Object.assign({}, invoice));}

    
    invoices.forEach((invoice) => {
        invoice.amount = invoice.items.reduce((sum, item) => {
            return sum + item.amount
        }, 0)
    });
    //('end parese...')
    console.log(invoices)
    return invoices;
}

 function splitInvoice(invoice, maxAmount) {
    let invoices = [];
    if (invoice.amount <= maxAmount) {
        return [Object.assign({}, invoice)];
    } else {
        let newInvoice = Object.assign({}, invoice, {items: []})
        let remain = maxAmount;
        for (let i = 0; i < invoice.items.length;i++) {
            let item = Object.assign({}, invoice.items[i])
            while (item.quantity > 0) { 
                if (remain >= item.quantity * item.price) {
                   // console.log(item.quantity)
                    newInvoice.items.push(Object.assign({}, item));
                    remain-=item.amount
                    break;
                } else {
                    let newItem = Object.assign({}, item)
                    let quantity = prettyQantity(remain, item.price)
                    newItem.quantity = quantity
                    newItem.amount = quantity * item.price;
                    newInvoice.items.push(newItem)

                    item.quantity -= quantity
                    item.amount -= quantity * item.price

                    //one invoice done and start a new one
                    newInvoice.amount=newInvoice.items.reduce((sum,item)=>sum+item.amount,0)
                    invoices.push(Object.assign({}, newInvoice))
                    //console.log("newInvoice")
                    //console.log(newInvoice)
                    newInvoice = Object.assign({}, invoice, {items: []})
                    remain = maxAmount
                }
            }
        }
        //save last invoice
        newInvoice.amount=newInvoice.items.reduce((sum,item)=>sum+item.amount,0)
        invoices.push(newInvoice)
    }
    invoices.map(invoice=>prettyInvoice(invoice))
    return invoices;
}

export function splitInvoices(invoices, maxAmount) {
   // console.log('begin split...')
    let newInvoices = [];
    invoices.forEach((invoice) => {
        newInvoices.push(...splitInvoice(invoice, maxAmount))
    })
   // console.log('end split...')
    //console.log(newInvoices)
    return newInvoices;
}

function prettyQantity(amount, price) {
    let quantity = amount / price;
    if(quantity<1){
        return Math.floor(quantity*100)/100;
    }
    quantity=parseInt(quantity);
    if (quantity > 1000) {
        return Math.floor(quantity / 100) * 100
    } else if (quantity > 100) {
        return Math.floor(quantity / 10) * 10
    } 
    return quantity;
}
function invioceXmlString(invoice){
    let buyer=invoice.buyer;
    let fpHeader=`
             <Gfmc>${buyer.buyerName}</Gfmc>            
             <Gfsh>${buyer.buyerTaxNum}</Gfsh>  
             <Gfyhzh>${buyer.buyerBankInfo}</Gfyhzh>   
             <Gfdzdh>${buyer.buyerAddrAndPhone}</Gfdzdh>
             <Bz>${invoice.notes}</Bz>
             <Fhr>${invoice.reviewer}</Fhr>
             <Skr>${invoice.payee}</Skr>
             <Spbmbbh>${invoice.Spbmbbh}</Spbmbbh>
             <Hsbz>${invoice.Hsbz}</Hsbz>
             <Spxx>
   
  `
     let fpItemsStrs=""
        invoice.items.forEach((item,index)=>{
            let fpItemStr=`
               <Sph>    
                 <Xh>${index+1}</Xh>
                 <Xmmc>${item.itemName}</Xmmc>
                 <Ggxh>${item.modelAndType}</Ggxh>
                 <Jldw>${item.unitType}</Jldw>
                 <Sl>${item.quantity}</Sl>
                 <Dj>${item.price}</Dj>
                 <Je>${item.amount}</Je>
                 <Slv>${item.taxRate}</Slv>
                 <Se>${item.amount * item.taxRate}</Se>
                 <Spbm>${item.Qyspbm}</Spbm>
                 <Zxbm>${item.Syyhzcbz}</Zxbm>
                 <Yhzcbs>${item.Lslbz}</Yhzcbs>
                 <Lslbs>${item.Yhzcsm}</Lslbs>
                 <Kce>${item.Kce}</Kce>
               </Sph>
        `
            fpItemsStrs+=fpItemStr
        })
    fpItemsStrs+=`
           </Spxx>
`
    return fpHeader+fpItemsStrs 
}

function invioceListXmlString(invoiceList){
    let str="";
    str=`
<?xml version="1.0" encoding="GBK" ?>
  <Kp>
    <Version>2.0</Version> 
    <Fpxx>
    <Zsl>${invoiceList.length}</Zsl>                         
    <Fpsj>
    `
    invoiceList.forEach((invoice,index)=>{
        str+=`
        <Fp>
          <Djh>${index+1}</Djh >
          `
        str+=invioceXmlString(invoice)
        str+=`
        </FP>
        `
    })

    str+=`    
     </Fpsj>
    </Fpxx>
  </Kp>`
    return str

}
function prettyInvoice(invoice,quantityDigits=2,priceDigits=4,amountDigits=2){
    invoice.items.forEach((item)=>{
        item.quantity=item.quantity.toFixed(quantityDigits)
        item.price=item.price.toFixed(priceDigits)
        item.amount=item.amount.toFixed(amountDigits)
    })
    invoice.amount=invoice.amount.toFixed(priceDigits)
    return invoice
}

export async function dowanloadAsXml(file){
   let workBook=await readFile(file)
   let invoices=parseInvoices(workBook)
   let invoiceList=splitInvoices(invoices,100000)
    let xmlStr=invioceListXmlString(invoiceList).replaceAll("undefined",'').replaceAll("null",'')
    let blob = new Blob([xmlStr], {type: "text/plain;charset=utf-8"});
    saveAs(blob, "invoice.xml");

}