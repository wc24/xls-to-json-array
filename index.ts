import path from "path";
import fs, { statSync } from "fs";
import { WorkBook, WorkSheet, readFile, utils, Range } from "xlsx";

let workPath = process.argv[2]
let confPath = process.argv[3]

let ls: number = 0
class ItemConf {
    xlsUrl: string | null
    book: WorkBook
    key: string
    out: any[] = []
    table: { checkCell: string, startCell: string, top?: string, down?: string, left?: string, right?: string }[]
    constructor(srcObject: any) {
        this.xlsUrl = srcObject["xlsUrl"]
        this.table = srcObject["table"]
        this.key = srcObject["key"]
        this.book = readFile(path.normalize(workPath + "/" + this.xlsUrl))
        for (let index = 0; index < this.table.length; index++) {
            const element = this.table[index];
            if (element != null) {
                let sheetOut: any[] = []
                this.out.push(sheetOut)
                let sheet = this.book.Sheets[this.book.SheetNames[index]]
                let range = utils.decode_range(sheet["!ref"]!)
                let top: number
                let left: number
                let down: number
                let right: number
                if (element.startCell == null) {
                    top = element.top ? utils.decode_row(element.top) : range.s.r
                    left = element.left ? utils.decode_col(element.left) : range.s.c
                } else {
                    let cell = utils.decode_cell(element.startCell)
                    top = cell.r
                    left = cell.c
                }
                down = element.down ? utils.decode_row(element.down) : range.e.r
                right = element.right ? utils.decode_col(element.right) : range.e.c
                let checkCell = utils.decode_cell(element.checkCell)
                let CY = checkCell.r
                let CX = checkCell.c
                for (let y = top; y < down+1; y++) {
                    let checkCell = sheet[utils.encode_cell({ c: CX, r: y })]
                    if (checkCell !== null && checkCell.v !== "" && checkCell.v !== " ") {
                        let item: any[] = []
                        sheetOut.push(item)
                        for (let x = left; x < right + 1; x++) {
                            let checkCell = sheet[utils.encode_cell({ c: x, r: CY })]
                            if (checkCell != null && checkCell.v != "" && checkCell.v != " ") {
                                let val: any = sheet[utils.encode_cell({ c: x, r: y })]
                                let out: any = ""
                                if (val != null) {
                                    out = val.v
                                    if (typeof val.v === "string") {
                                        if (val.v.slice(0, 1) == "[" || val.v.slice(0, 1) == "{") {
                                            try {
                                                out = JSON.parse(val.v)
                                            } catch (error) {
                                            }
                                        }
                                    }
                                }
                                item.push(out)
                            }

                        }
                    }else{
                    }
                }
            }
        }
    } 
}
function toJson() {
    let txt = fs.readFileSync(path.normalize(workPath + "/" + confPath), "utf-8")
    let task = JSON.parse(txt)
    for (const dom of task) {
        let domOut: any = {}
        for (const item of dom.src) {
            let itemConf = new ItemConf(item)
            domOut[itemConf.key] = itemConf.out
        }
        let wpath = path.normalize(workPath + "/" + dom.out)
        let dir = path.dirname(wpath)
        if (!fs.existsSync(dir)) {
            fs.mkdirSync(dir)
        }
        fs.writeFileSync(wpath, JSON.stringify(domOut))
    }
}
toJson()
