import * as XLSX from "xlsx";

export function imprimir() {
    /* create new workbook */
    let workbook = XLSX.utils.book_new();
    let i = 5
    let ws_data = [['hello' , 'world']];
    let ws = XLSX.utils.aoa_to_sheet(ws_data);

    workbook.SheetNames.push("Test Sheet");
    workbook.Sheets["Test Sheet"] = ws;

    XLSX.writeFile(workbook, 'out.xls');
};
