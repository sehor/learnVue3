'use-strict'
import ExcelJS from 'exceljs'
class Person {
    height = 10;
    constructor(name, age) {
        this.name = name;
        this.age = age;
    }
    getAge() {
        return this.age;
    }
}
//iterate Person properties


function fieldToColumnMap(titleMap, fieldToTitleMap) {
    if (!fieldToTitleMap) {
        return titleMap;
    }
    let map = new Map();
    for ([key, value] of Object.entries(fieldToTitleMap)) {
        let v1 = titleMap.get(value);
        if (v1) {
            map.set(key, v1);
        } else {
            console.warn("titleMap中没有" + value + "对应的列号")
        }

    }
    return map;
}

let workbook=ExcelJS.Workbook()
let worksheet=workbook.addWorksheet('sheet1')
 //add row to worksheet
let row=worksheet.addRow(['name','age','height'])

