<template>
    <div>
        <button @click="onClickPrint">출력</button>
        <table :id="tableId">
            <thead>
            <tr>
                <th v-for="[key, value] in Object.entries(header)">{{ value }}</th>
            </tr>
            </thead>
            <tbody>
            <tr v-for="product in list">
                <td>{{ product.index }}</td>
                <td>{{ product.name }}</td>
                <td>{{ product.price }}</td>
            </tr>
            </tbody>
        </table>
    </div>
</template>

<script>
const FILE_NAME = "test"
const SHEET_NAME = "테스트 시트"
const FILE_EXTENSION = ".xlsx"

import { saveAs } from 'file-saver'
const XLSX = require('xlsx')

export default {
    name: "ExportTable",
    data() {
        return {
            header: {
                name: "이름",
                index: "번호",
                price: "가격"
            },
            list: [],
            tableId: "tableData"
        }
    },
    created() {
        console.log("page created")

        let i = 0
        do {
            this.list.push({
                name: `test ${i}`,
                index: i,
                price: Math.floor(Math.random() * 1000)
            })

            i++
        }
        while (i < 100)

    },
    mounted() {
        console.log("page mounted")
    },
    methods: {
        onClickPrint() {
            this.exportExcel()
        },
        exportExcel() {
            // step 1. workbook 생성
            let wb = XLSX.utils.book_new();

            // step 2. worksheet 생성
            let ws = this.getWorksheet()

            // step 3. workbook 에 새로만든 워크시트에 이름을 주고 붙인다.
            XLSX.utils.book_append_sheet(wb, ws, SHEET_NAME)

            // step 4. 엑셀 파일 만들기
            let wbt = XLSX.write(wb, {
                bookType: 'xlsx',
                type: 'binary',
            })

            // step 5. 엑셀 파일 내보내기
            saveAs(new Blob([this.s2ab(wbt)], {type: "application/octet-stream"}), this.getExcelFileName())
        },
        getExcelData: function () {
            return document.getElementById(this.tableId)
        },
        getWorksheet: function () {
            return XLSX.utils.table_to_sheet(this.getExcelData(), {
                raw: true,
                cellStyles: true
            })
        },
        getExcelFileName() {
            let year = new Date().getFullYear()
            let month = new Date().getMonth()
            let date = new Date().getDate()

            return FILE_NAME + '_' + `${year}-${month}-${date}` + FILE_EXTENSION
        },
        s2ab(s) {

            //convert s to arrayBuffer
            let buf = new ArrayBuffer(s.length)

            //create uint8array as viewer
            let view = new Uint8Array(buf)

            //convert to octet
            for (let i = 0; i < s.length; i++) {
                view[i] = s.charCodeAt(i) & 0xFF
            }

            return buf
        },
    }
}
</script>

<style scoped>

</style>
