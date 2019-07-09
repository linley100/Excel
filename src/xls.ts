import * as XLSX from "xlsx";

export function imprimir() {
    /* create new workbook */
    let workbook = XLSX.utils.book_new();
    let i = 5
    let ws_data = [['','hello' , 'world'],['','hello' , 'world']];
    let ws = XLSX.utils.aoa_to_sheet(ws_data);

	let address = 'B3';
	ws[address] = 'loco';

	ws['C5'] = { t:'s', v: "cosas"};
	ws['C1'] = { t:'n', v: 5 };
	ws['A3'] = { t:'l', Target:"http://sheetjs.com", Tooltip:"Find us @ SheetJS.com!" };
	ws['A5'] = ({ t:'n', W: 'muchas cosas'});
	ws['A6'] = ({ t:'z', v: 'Ncosas'});

    workbook.SheetNames.push("Test Sheet");
    workbook.Sheets["Test Sheet"] = ws;

    XLSX.writeFile(workbook, 'prueba.xls');
};
