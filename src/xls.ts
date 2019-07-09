import * as XLSX from "xlsx";

export function imprimir() {
    /* create new workbook */
    let workbook = XLSX.utils.book_new();
    
    let i;
    let ws_data = [];
    //Inicializar las celdas a usar
    for (i = 0; i < 20; i++) { 
        ws_data[i] = ['','hello' , 'world','','','','','','','','','',''];
    }
    let ws = XLSX.utils.aoa_to_sheet(ws_data);

    let array = ['a','b','c','d'];
	let address = 'B3';

    //escribir en una celda en especifico
    ws['C2'] = { t:'s', v: address};
    ws['C6'] = { t:'s', v: address};
    ws['C10'] = { t:'s', v: address};
    ws['C3'] = { t:'s', v: array[3]};
    ws['C1'] = { t:'n', v: 500 };
    
    //combinar celdas
    let rango = [{s: { c: 0, r: 0 }, e: { c: 0, r: 4 }}]
    ws['!merges'] = rango;

    workbook.SheetNames.push("Test Sheet");
    workbook.Sheets["Test Sheet"] = ws;

    XLSX.writeFile(workbook, 'prueba.xls');
};
