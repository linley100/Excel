import * as XLSX from "xlsx";

export function imprimir(mes, dias) {
    // este el libro de trabajo
    let workbook = XLSX.utils.book_new();
    
    //Auxiliares y contadores
    let aux=0, i=0, j=0, cont=0, contD=0, totalD=0;

    //variable tipo array de string, contiene los meses
    let meses = mes;
    //Variable que guarda la cantidad de meses
    let x = mes.length;
    //Variable tipo arrays of arrrays of arrays de int, contiene los dias de un turno de un mes
    let cantDias =  dias;
    /*let cantDias = [
        [
            [100,200,300,400,500,
            100,200,300,400,500,
            100,200,300,400,500,
            100,200,300,400,500,
            100,200,300,400,500,
            100,200,300,400,500],
            [10,20,30,40,50,
            10,20,30,40,50,
            10,20,30,40,50,
            10,20,30,40,50,
            10,20,30,40,50,
            10,20,30,40,50],
            [1,2,3,4,5,
            1,2,3,4,5,
            1,2,3,4,5,
            1,2,3,4,5,
            1,2,3,4,5,
            1,2,3,4,5]
        ],
        [
            [100,200,300,400,500,
            100,200,300,400,500,
            100,200,300,400,500,
            100,200,300,400,500,
            100,200,300,400,500,
            100,200,300,400,500],
            [10,20,30,40,50,
            10,20,30,40,50,
            10,20,30,40,50,
            10,20,30,40,50,
            10,20,30,40,50,
            10,20,30,40,50],
            [1,2,3,4,5,
            1,2,3,4,5,
            1,2,3,4,5,
            1,2,3,4,5,
            1,2,3,4,5,
            1,2,3,4,5]
        ]
    ];*/

    //Variables de prueba, dias por mes
    let a = cantDias[0][0];
    let b = cantDias[0][0];
    let c = cantDias[0][0];

    //Celdas de la hoja de calculo
    let ws_data = [];

    //Inicializar las celdas a usar
    for (i = 0; i < (34 * x); i++) { 
        ws_data[i] = ['','','','','','','','','','','','','','',''];
    }

    //estas son las hojas de calculo
    let piso1 = XLSX.utils.aoa_to_sheet(ws_data);
    let piso3 = XLSX.utils.aoa_to_sheet(ws_data);
    let hemeroteca = XLSX.utils.aoa_to_sheet(ws_data);

    //Variable para la para combinar las celdas
    let rango = [{s: { c: 1, r: 5 }, e: { c: 12, r: 5 }}]; 
    let rangoAux;

    //Array con palabras usadas
    let text = [
        'CANTIDAD DE USUARIOS SALAS DE ESTUDIOS',
        'TURNO','1er turno','2do turno','3er turno','RESPONSABLE',
        'FECHA','LUN','MART','MIER','JUEV','VIER','SAB','CANT','TOTAL'
    ];


    //Llenar la hoja de excel del piso 1
    piso1['B' + 6] = { t:'s', v: text[0]};
    for (i = 0; i < x; i++) { 
        totalD = 0;
        aux = 7

        //Escribir texto puntual
        piso1['B' + (8 + 27*i)] = { t:'s', v: text[1]};
        piso1['B' + (10 + 27*i)] = { t:'s', v: text[2]};
        piso1['B' + (17 + 27*i)] = { t:'s', v: text[3]};
        piso1['B' + (24 + 27*i)] = { t:'s', v: text[4]};
        piso1['B' + (31 + 27*i)] = { t:'s', v: text[14]};
        piso1['C' + (8 + 27*i)] = { t:'s', v: text[5]};
        piso1['D' + (8 + 27*i)] = { t:'s', v: text[6]};
        piso1['E' + (8 + 27*i)] = { t:'s', v: meses[i]};
        piso1['E' + (9 + 27*i)] = { t:'s', v: text[13]};
        piso1['F' + (9 + 27*i)] = { t:'s', v: text[6]};
        piso1['G' + (9 + 27*i)] = { t:'s', v: text[13]};
        piso1['H' + (9 + 27*i)] = { t:'s', v: text[6]};
        piso1['I' + (9 + 27*i)] = { t:'s', v: text[13]};
        piso1['J' + (9 + 27*i)] = { t:'s', v: text[6]};
        piso1['K' + (9 + 27*i)] = { t:'s', v: text[13]};
        piso1['L' + (9 + 27*i)] = { t:'s', v: text[6]};
        piso1['M' + (9 + 27*i)] = { t:'s', v: text[13]};

        //Escribir texto repetitivo
        for (j = 10; j < 30; j++) {
            
            if(j == 16){ j++ }
            if(j == 23){ j++ }
            
            //Dias de la semana
            piso1['D' + (j + 27*i)] = { t:'s', v: text[aux]};
            piso1['F' + (j + 27*i)] = { t:'s', v: text[aux]};
            piso1['H' + (j + 27*i)] = { t:'s', v: text[aux]};
            piso1['J' + (j + 27*i)] = { t:'s', v: text[aux]};
            piso1['L' + (j + 27*i)] = { t:'s', v: text[aux]};

            //cantidad de usuarios en un dia
            piso1['E' + (j + 27*i)] = { t:'n', v: a[(0 + cont)]};
            piso1['G' + (j + 27*i)] = { t:'n', v: a[(6 + cont)]};
            piso1['I' + (j + 27*i)] = { t:'n', v: a[(12 + cont)]};
            piso1['K' + (j + 27*i)] = { t:'n', v: a[(18 + cont)]};
            piso1['M' + (j + 27*i)] = { t:'n', v: a[(24 + cont)]};
            
            //Sumatoria de los usuarios
            totalD = totalD + a[(0 + cont)] + a[(6 + cont)] + a[(12 + cont)] + a[(18 + cont)] + a[(24 + cont)];
            
            if(aux < 12){
                aux++;
            }else{
                aux = 7;
            }

            if(cont < 5){
                cont++;
            }else{
                cont=0;
                if(contD < 2){
                    contD++;
                }else{
                    contD=0;
                }
                a = cantDias[i][contD];
            }
            
        }

        //Total de usuarios ese mes
        piso1['C' + (31 + 27*i)] = { t:'n', v: totalD};

        //combinar celdas, 's' es la celda inicial y 'e' es la celda final 
        //'c' es la columna y 'r' es la fila. (A1 esta en la posicion '0,0', es decir, c:0 r:0)
        rangoAux = [
            {s: { c: 1, r: (6 + 27*i) }, e: { c: 12, r: (6 + 27*i) }},
            {s: { c: 1, r: (15 + 27*i) }, e: { c: 12, r: (15 + 27*i) }},
            {s: { c: 1, r: (22 + 27*i) }, e: { c: 12, r: (22 + 27*i) }},
            {s: { c: 1, r: (29 + 27*i) }, e: { c: 12, r: (29 + 27*i) }},
            {s: { c: 1, r: (7 + 27*i) }, e: { c: 1, r: (8 + 27*i) }},
            {s: { c: 1, r: (9 + 27*i) }, e: { c: 1, r: (14 + 27*i) }},
            {s: { c: 1, r: (16 + 27*i) }, e: { c: 1, r: (21 + 27*i) }},
            {s: { c: 1, r: (23 + 27*i) }, e: { c: 1, r: (28 + 27*i) }},
            {s: { c: 2, r: (7 + 27*i) }, e: { c: 2, r: (8 + 27*i) }},
            {s: { c: 2, r: (9 + 27*i) }, e: { c: 2, r: (14 + 27*i) }},
            {s: { c: 2, r: (16 + 27*i) }, e: { c: 2, r: (21 + 27*i) }},
            {s: { c: 2, r: (23 + 27*i) }, e: { c: 2, r: (28 + 27*i) }},
            {s: { c: 2, r: (30 + 27*i) }, e: { c: 12, r: (30 + 27*i) }},
            {s: { c: 3, r: (7 + 27*i) }, e: { c: 3, r: (8 + 27*i) }},
            {s: { c: 4, r: (7 + 27*i) }, e: { c: 12, r: (7 + 27*i) }}
        ]
        rango = rango.concat(rangoAux);
        piso1['!merges'] = rango;
    }

    //Llenar la hoja de excel del piso 3
    cont = 0;
    contD = 0;
    rango = [{s: { c: 1, r: 5 }, e: { c: 12, r: 5 }}]; 
    piso3['B' + 6] = { t:'s', v: text[0]};
    for (i = 0; i < x; i++) { 
        totalD = 0;
        aux = 7

        piso3['B' + (8 + 27*i)] = { t:'s', v: text[1]};
        piso3['B' + (10 + 27*i)] = { t:'s', v: text[2]};
        piso3['B' + (17 + 27*i)] = { t:'s', v: text[3]};
        piso3['B' + (24 + 27*i)] = { t:'s', v: text[4]};
        piso3['B' + (31 + 27*i)] = { t:'s', v: text[14]};
        piso3['C' + (8 + 27*i)] = { t:'s', v: text[5]};
        piso3['D' + (8 + 27*i)] = { t:'s', v: text[6]};
        piso3['E' + (8 + 27*i)] = { t:'s', v: meses[i]};
        piso3['E' + (9 + 27*i)] = { t:'s', v: text[13]};
        piso3['F' + (9 + 27*i)] = { t:'s', v: text[6]};
        piso3['G' + (9 + 27*i)] = { t:'s', v: text[13]};
        piso3['H' + (9 + 27*i)] = { t:'s', v: text[6]};
        piso3['I' + (9 + 27*i)] = { t:'s', v: text[13]};
        piso3['J' + (9 + 27*i)] = { t:'s', v: text[6]};
        piso3['K' + (9 + 27*i)] = { t:'s', v: text[13]};
        piso3['L' + (9 + 27*i)] = { t:'s', v: text[6]};
        piso3['M' + (9 + 27*i)] = { t:'s', v: text[13]};

        for (j = 10; j < 30; j++) {

            if(j == 16){ j++ }
            if(j == 23){ j++ }
            
            //Dias de la semana
            piso3['D' + (j + 27*i)] = { t:'s', v: text[aux]};
            piso3['F' + (j + 27*i)] = { t:'s', v: text[aux]};
            piso3['H' + (j + 27*i)] = { t:'s', v: text[aux]};
            piso3['J' + (j + 27*i)] = { t:'s', v: text[aux]};
            piso3['L' + (j + 27*i)] = { t:'s', v: text[aux]};

            //cantidad de usuarios en un dia
            piso3['E' + (j + 27*i)] = { t:'n', v: b[(0 + cont)]};
            piso3['G' + (j + 27*i)] = { t:'n', v: b[(6 + cont)]};
            piso3['I' + (j + 27*i)] = { t:'n', v: b[(12 + cont)]};
            piso3['K' + (j + 27*i)] = { t:'n', v: b[(18 + cont)]};
            piso3['M' + (j + 27*i)] = { t:'n', v: b[(24 + cont)]};
           
            //Sumatoria de los usuarios
            totalD = totalD + b[(0 + cont)] + b[(6 + cont)] + b[(12 + cont)] + b[(18 + cont)] + b[(24 + cont)];

            if(aux < 12){
                aux++;
            }else{
                aux = 7;
            }
            
            if(cont < 5){
                cont++;
            }else{
                cont=0;
                if(contD < 2){
                    contD++;
                }else{
                    contD=0;
                }
                b = cantDias[i][contD];
            }

        }

        //Total de usuarios ese mes
        piso3['C' + (31 + 27*i)] = { t:'n', v: totalD};

        //combinar celdas
        rangoAux = [
            {s: { c: 1, r: (6 + 27*i) }, e: { c: 12, r: (6 + 27*i) }},
            {s: { c: 1, r: (15 + 27*i) }, e: { c: 12, r: (15 + 27*i) }},
            {s: { c: 1, r: (22 + 27*i) }, e: { c: 12, r: (22 + 27*i) }},
            {s: { c: 1, r: (29 + 27*i) }, e: { c: 12, r: (29 + 27*i) }},
            {s: { c: 1, r: (7 + 27*i) }, e: { c: 1, r: (8 + 27*i) }},
            {s: { c: 1, r: (9 + 27*i) }, e: { c: 1, r: (14 + 27*i) }},
            {s: { c: 1, r: (16 + 27*i) }, e: { c: 1, r: (21 + 27*i) }},
            {s: { c: 1, r: (23 + 27*i) }, e: { c: 1, r: (28 + 27*i) }},
            {s: { c: 2, r: (7 + 27*i) }, e: { c: 2, r: (8 + 27*i) }},
            {s: { c: 2, r: (9 + 27*i) }, e: { c: 2, r: (14 + 27*i) }},
            {s: { c: 2, r: (16 + 27*i) }, e: { c: 2, r: (21 + 27*i) }},
            {s: { c: 2, r: (23 + 27*i) }, e: { c: 2, r: (28 + 27*i) }},
            {s: { c: 2, r: (30 + 27*i) }, e: { c: 12, r: (30 + 27*i) }},
            {s: { c: 3, r: (7 + 27*i) }, e: { c: 3, r: (8 + 27*i) }},
            {s: { c: 4, r: (7 + 27*i) }, e: { c: 12, r: (7 + 27*i) }}
        ]
        rango = rango.concat(rangoAux);
        piso3['!merges'] = rango;

    }
    
    //Llenar la hoja de excel del hemeroteca
    cont = 0;
    contD = 0;
    rango = [{s: { c: 1, r: 5 }, e: { c: 12, r: 5 }}]; 
    hemeroteca['B' + 6] = { t:'s', v: text[0]};
    for (i = 0; i < x; i++) { 
        totalD = 0;
        aux = 7

        hemeroteca['B' + (8 + 20*i)] = { t:'s', v: text[1]};
        hemeroteca['B' + (10 + 20*i)] = { t:'s', v: text[2]};
        hemeroteca['B' + (17 + 20*i)] = { t:'s', v: text[3]};
        hemeroteca['B' + (24 + 20*i)] = { t:'s', v: text[14]};
        hemeroteca['C' + (8 + 20*i)] = { t:'s', v: text[5]};
        hemeroteca['D' + (8 + 20*i)] = { t:'s', v: text[6]};
        hemeroteca['E' + (8 + 20*i)] = { t:'s', v: meses[i]};
        hemeroteca['E' + (9 + 20*i)] = { t:'s', v: text[13]};
        hemeroteca['F' + (9 + 20*i)] = { t:'s', v: text[6]};
        hemeroteca['G' + (9 + 20*i)] = { t:'s', v: text[13]};
        hemeroteca['H' + (9 + 20*i)] = { t:'s', v: text[6]};
        hemeroteca['I' + (9 + 20*i)] = { t:'s', v: text[13]};
        hemeroteca['J' + (9 + 20*i)] = { t:'s', v: text[6]};
        hemeroteca['K' + (9 + 20*i)] = { t:'s', v: text[13]};
        hemeroteca['L' + (9 + 20*i)] = { t:'s', v: text[6]};
        hemeroteca['M' + (9 + 20*i)] = { t:'s', v: text[13]};

        for (j = 10; j < 23; j++) {

            if(j == 16){ j++ }
            
            //Dias de la semana
            hemeroteca['D' + (j + 20*i)] = { t:'s', v: text[aux]};
            hemeroteca['F' + (j + 20*i)] = { t:'s', v: text[aux]};
            hemeroteca['H' + (j + 20*i)] = { t:'s', v: text[aux]};
            hemeroteca['J' + (j + 20*i)] = { t:'s', v: text[aux]};
            hemeroteca['L' + (j + 20*i)] = { t:'s', v: text[aux]};

            //cantidad de usuarios en un dia
            hemeroteca['E' + (j + 20*i)] = { t:'n', v: c[(0 + cont)]};
            hemeroteca['G' + (j + 20*i)] = { t:'n', v: c[(6 + cont)]};
            hemeroteca['I' + (j + 20*i)] = { t:'n', v: c[(12 + cont)]};
            hemeroteca['K' + (j + 20*i)] = { t:'n', v: c[(18 + cont)]};
            hemeroteca['M' + (j + 20*i)] = { t:'n', v: c[(24 + cont)]};
            
            //Sumatoria de los dias
            totalD = totalD + c[(0 + cont)] + c[(6 + cont)] + c[(12 + cont)] + c[(18 + cont)] + c[(24 + cont)];

            if(aux < 12){
                aux++;
            }else{
                aux = 7;
            }

            if(cont < 5){
                cont++;
            }else{
                cont=0;
                if(contD < 1){
                    contD++;
                }else{
                    contD=0;
                }
                c = cantDias[i][contD];
            }

        }

        //Total de usuarios ese mes
        hemeroteca['C' + (24 + 20*i)] = { t:'n', v: totalD};

        //combinar celdas
        rangoAux = [
            {s: { c: 1, r: (6 + 20*i) }, e: { c: 12, r: (6 + 20*i) }},
            {s: { c: 1, r: (15 + 20*i) }, e: { c: 12, r: (15 + 20*i) }},
            {s: { c: 1, r: (7 + 20*i) }, e: { c: 1, r: (8 + 20*i) }},
            {s: { c: 1, r: (9 + 20*i) }, e: { c: 1, r: (14 + 20*i) }},
            {s: { c: 1, r: (16 + 20*i) }, e: { c: 1, r: (21 + 20*i) }},
            {s: { c: 1, r: (22 + 20*i) }, e: { c: 12, r: (22 + 20*i) }},
            {s: { c: 2, r: (7 + 20*i) }, e: { c: 2, r: (8 + 20*i) }},
            {s: { c: 2, r: (9 + 20*i) }, e: { c: 2, r: (14 + 20*i) }},
            {s: { c: 2, r: (16 + 20*i) }, e: { c: 2, r: (21 + 20*i) }},
            {s: { c: 2, r: (23 + 20*i) }, e: { c: 12, r: (23 + 20*i) }},
            {s: { c: 3, r: (7 + 20*i) }, e: { c: 3, r: (8 + 20*i) }},
            {s: { c: 4, r: (7 + 20*i) }, e: { c: 12, r: (7 + 20*i) }}
        ]
        rango = rango.concat(rangoAux);
        hemeroteca['!merges'] = rango;

    }

    workbook.SheetNames.push("Piso1");
    workbook.SheetNames.push("Piso3");
    workbook.SheetNames.push("Hemeroteca");
    workbook.Sheets["Piso1"] = piso1;
    workbook.Sheets["Piso3"] = piso3;
    workbook.Sheets["Hemeroteca"] = hemeroteca;

    XLSX.writeFile(workbook, 'UsuariosDeLaSalaDeEstudio.xls');
};

export function imprimirSalas(mes, dias) {
    // este el libro de trabajo
    let workbook = XLSX.utils.book_new();
    
    //Auxiliares y contadores
    let aux=0, i=0, j=0, cont=0, contD=0, totalD=0;

    //variable tipo array de string, contiene los meses
    let meses = mes;
    //Variable que guarda la cantidad de meses
    let x = mes.length;
    //Variable tipo arrays of arrrays of arrays de int, contiene los dias de un turno de un mes
    let cantDias =  dias;
    /*let cantDias = [
        [
            [100,200,300,400,500,
            100,200,300,400,500,
            100,200,300,400,500,
            100,200,300,400,500,
            100,200,300,400,500,
            100,200,300,400,500],
            [10,20,30,40,50,
            10,20,30,40,50,
            10,20,30,40,50,
            10,20,30,40,50,
            10,20,30,40,50,
            10,20,30,40,50],
            [1,2,3,4,5,
            1,2,3,4,5,
            1,2,3,4,5,
            1,2,3,4,5,
            1,2,3,4,5,
            1,2,3,4,5]
        ],
        [
            [100,200,300,400,500,
            100,200,300,400,500,
            100,200,300,400,500,
            100,200,300,400,500,
            100,200,300,400,500,
            100,200,300,400,500],
            [10,20,30,40,50,
            10,20,30,40,50,
            10,20,30,40,50,
            10,20,30,40,50,
            10,20,30,40,50,
            10,20,30,40,50],
            [1,2,3,4,5,
            1,2,3,4,5,
            1,2,3,4,5,
            1,2,3,4,5,
            1,2,3,4,5,
            1,2,3,4,5]
        ]
    ];*/

    //Variables de prueba, dias por mes
    let a = cantDias[0][0];
    let b = cantDias[0][0];
    let c = cantDias[0][0];

    //Celdas de la hoja de calculo
    let ws_data = [];

    //Inicializar las celdas a usar
    for (i = 0; i < (34 * x); i++) { 
        ws_data[i] = ['','','','','','','','','','','','','','',''];
    }

    //estas son las hojas de calculo
    let piso1 = XLSX.utils.aoa_to_sheet(ws_data);
    let piso3 = XLSX.utils.aoa_to_sheet(ws_data);
    let hemeroteca = XLSX.utils.aoa_to_sheet(ws_data);

    //Variable para la para combinar las celdas
    let rango = [{s: { c: 1, r: 5 }, e: { c: 12, r: 5 }}]; 
    let rangoAux;

    //Array con palabras usadas
    let text = [
        'SALAS DE ESTUDIOS SOLICITADAS',
        'TURNO','1er turno','2do turno','3er turno','RESPONSABLE',
        'FECHA','LUN','MART','MIER','JUEV','VIER','SAB','CANT','TOTAL'
    ];


    //Llenar la hoja de excel del piso 1
    piso1['B' + 6] = { t:'s', v: text[0]};
    for (i = 0; i < x; i++) { 
        totalD = 0;
        aux = 7

        //Escribir texto puntual
        piso1['B' + (8 + 27*i)] = { t:'s', v: text[1]};
        piso1['B' + (10 + 27*i)] = { t:'s', v: text[2]};
        piso1['B' + (17 + 27*i)] = { t:'s', v: text[3]};
        piso1['B' + (24 + 27*i)] = { t:'s', v: text[4]};
        piso1['B' + (31 + 27*i)] = { t:'s', v: text[14]};
        piso1['C' + (8 + 27*i)] = { t:'s', v: text[5]};
        piso1['D' + (8 + 27*i)] = { t:'s', v: text[6]};
        piso1['E' + (8 + 27*i)] = { t:'s', v: meses[i]};
        piso1['E' + (9 + 27*i)] = { t:'s', v: text[13]};
        piso1['F' + (9 + 27*i)] = { t:'s', v: text[6]};
        piso1['G' + (9 + 27*i)] = { t:'s', v: text[13]};
        piso1['H' + (9 + 27*i)] = { t:'s', v: text[6]};
        piso1['I' + (9 + 27*i)] = { t:'s', v: text[13]};
        piso1['J' + (9 + 27*i)] = { t:'s', v: text[6]};
        piso1['K' + (9 + 27*i)] = { t:'s', v: text[13]};
        piso1['L' + (9 + 27*i)] = { t:'s', v: text[6]};
        piso1['M' + (9 + 27*i)] = { t:'s', v: text[13]};

        //Escribir texto repetitivo
        for (j = 10; j < 30; j++) {
            
            if(j == 16){ j++ }
            if(j == 23){ j++ }
            
            //Dias de la semana
            piso1['D' + (j + 27*i)] = { t:'s', v: text[aux]};
            piso1['F' + (j + 27*i)] = { t:'s', v: text[aux]};
            piso1['H' + (j + 27*i)] = { t:'s', v: text[aux]};
            piso1['J' + (j + 27*i)] = { t:'s', v: text[aux]};
            piso1['L' + (j + 27*i)] = { t:'s', v: text[aux]};

            //cantidad de usuarios en un dia
            piso1['E' + (j + 27*i)] = { t:'n', v: a[(0 + cont)]};
            piso1['G' + (j + 27*i)] = { t:'n', v: a[(6 + cont)]};
            piso1['I' + (j + 27*i)] = { t:'n', v: a[(12 + cont)]};
            piso1['K' + (j + 27*i)] = { t:'n', v: a[(18 + cont)]};
            piso1['M' + (j + 27*i)] = { t:'n', v: a[(24 + cont)]};
            
            //Sumatoria de los usuarios
            totalD = totalD + a[(0 + cont)] + a[(6 + cont)] + a[(12 + cont)] + a[(18 + cont)] + a[(24 + cont)];
            
            if(aux < 12){
                aux++;
            }else{
                aux = 7;
            }

            if(cont < 5){
                cont++;
            }else{
                cont=0;
                if(contD < 2){
                    contD++;
                }else{
                    contD=0;
                }
                a = cantDias[i][contD];
            }
            
        }

        //Total de usuarios ese mes
        piso1['C' + (31 + 27*i)] = { t:'n', v: totalD};

        //combinar celdas, 's' es la celda inicial y 'e' es la celda final 
        //'c' es la columna y 'r' es la fila. (A1 esta en la posicion '0,0', es decir, c:0 r:0)
        rangoAux = [
            {s: { c: 1, r: (6 + 27*i) }, e: { c: 12, r: (6 + 27*i) }},
            {s: { c: 1, r: (15 + 27*i) }, e: { c: 12, r: (15 + 27*i) }},
            {s: { c: 1, r: (22 + 27*i) }, e: { c: 12, r: (22 + 27*i) }},
            {s: { c: 1, r: (29 + 27*i) }, e: { c: 12, r: (29 + 27*i) }},
            {s: { c: 1, r: (7 + 27*i) }, e: { c: 1, r: (8 + 27*i) }},
            {s: { c: 1, r: (9 + 27*i) }, e: { c: 1, r: (14 + 27*i) }},
            {s: { c: 1, r: (16 + 27*i) }, e: { c: 1, r: (21 + 27*i) }},
            {s: { c: 1, r: (23 + 27*i) }, e: { c: 1, r: (28 + 27*i) }},
            {s: { c: 2, r: (7 + 27*i) }, e: { c: 2, r: (8 + 27*i) }},
            {s: { c: 2, r: (9 + 27*i) }, e: { c: 2, r: (14 + 27*i) }},
            {s: { c: 2, r: (16 + 27*i) }, e: { c: 2, r: (21 + 27*i) }},
            {s: { c: 2, r: (23 + 27*i) }, e: { c: 2, r: (28 + 27*i) }},
            {s: { c: 2, r: (30 + 27*i) }, e: { c: 12, r: (30 + 27*i) }},
            {s: { c: 3, r: (7 + 27*i) }, e: { c: 3, r: (8 + 27*i) }},
            {s: { c: 4, r: (7 + 27*i) }, e: { c: 12, r: (7 + 27*i) }}
        ]
        rango = rango.concat(rangoAux);
        piso1['!merges'] = rango;
    }

    //Llenar la hoja de excel del piso 3
    cont = 0;
    contD = 0;
    rango = [{s: { c: 1, r: 5 }, e: { c: 12, r: 5 }}]; 
    piso3['B' + 6] = { t:'s', v: text[0]};
    for (i = 0; i < x; i++) { 
        totalD = 0;
        aux = 7

        piso3['B' + (8 + 27*i)] = { t:'s', v: text[1]};
        piso3['B' + (10 + 27*i)] = { t:'s', v: text[2]};
        piso3['B' + (17 + 27*i)] = { t:'s', v: text[3]};
        piso3['B' + (24 + 27*i)] = { t:'s', v: text[4]};
        piso3['B' + (31 + 27*i)] = { t:'s', v: text[14]};
        piso3['C' + (8 + 27*i)] = { t:'s', v: text[5]};
        piso3['D' + (8 + 27*i)] = { t:'s', v: text[6]};
        piso3['E' + (8 + 27*i)] = { t:'s', v: meses[i]};
        piso3['E' + (9 + 27*i)] = { t:'s', v: text[13]};
        piso3['F' + (9 + 27*i)] = { t:'s', v: text[6]};
        piso3['G' + (9 + 27*i)] = { t:'s', v: text[13]};
        piso3['H' + (9 + 27*i)] = { t:'s', v: text[6]};
        piso3['I' + (9 + 27*i)] = { t:'s', v: text[13]};
        piso3['J' + (9 + 27*i)] = { t:'s', v: text[6]};
        piso3['K' + (9 + 27*i)] = { t:'s', v: text[13]};
        piso3['L' + (9 + 27*i)] = { t:'s', v: text[6]};
        piso3['M' + (9 + 27*i)] = { t:'s', v: text[13]};

        for (j = 10; j < 30; j++) {

            if(j == 16){ j++ }
            if(j == 23){ j++ }
            
            //Dias de la semana
            piso3['D' + (j + 27*i)] = { t:'s', v: text[aux]};
            piso3['F' + (j + 27*i)] = { t:'s', v: text[aux]};
            piso3['H' + (j + 27*i)] = { t:'s', v: text[aux]};
            piso3['J' + (j + 27*i)] = { t:'s', v: text[aux]};
            piso3['L' + (j + 27*i)] = { t:'s', v: text[aux]};

            //cantidad de usuarios en un dia
            piso3['E' + (j + 27*i)] = { t:'n', v: b[(0 + cont)]};
            piso3['G' + (j + 27*i)] = { t:'n', v: b[(6 + cont)]};
            piso3['I' + (j + 27*i)] = { t:'n', v: b[(12 + cont)]};
            piso3['K' + (j + 27*i)] = { t:'n', v: b[(18 + cont)]};
            piso3['M' + (j + 27*i)] = { t:'n', v: b[(24 + cont)]};
           
            //Sumatoria de los usuarios
            totalD = totalD + b[(0 + cont)] + b[(6 + cont)] + b[(12 + cont)] + b[(18 + cont)] + b[(24 + cont)];

            if(aux < 12){
                aux++;
            }else{
                aux = 7;
            }
            
            if(cont < 5){
                cont++;
            }else{
                cont=0;
                if(contD < 2){
                    contD++;
                }else{
                    contD=0;
                }
                b = cantDias[i][contD];
            }

        }

        //Total de usuarios ese mes
        piso3['C' + (31 + 27*i)] = { t:'n', v: totalD};

        //combinar celdas
        rangoAux = [
            {s: { c: 1, r: (6 + 27*i) }, e: { c: 12, r: (6 + 27*i) }},
            {s: { c: 1, r: (15 + 27*i) }, e: { c: 12, r: (15 + 27*i) }},
            {s: { c: 1, r: (22 + 27*i) }, e: { c: 12, r: (22 + 27*i) }},
            {s: { c: 1, r: (29 + 27*i) }, e: { c: 12, r: (29 + 27*i) }},
            {s: { c: 1, r: (7 + 27*i) }, e: { c: 1, r: (8 + 27*i) }},
            {s: { c: 1, r: (9 + 27*i) }, e: { c: 1, r: (14 + 27*i) }},
            {s: { c: 1, r: (16 + 27*i) }, e: { c: 1, r: (21 + 27*i) }},
            {s: { c: 1, r: (23 + 27*i) }, e: { c: 1, r: (28 + 27*i) }},
            {s: { c: 2, r: (7 + 27*i) }, e: { c: 2, r: (8 + 27*i) }},
            {s: { c: 2, r: (9 + 27*i) }, e: { c: 2, r: (14 + 27*i) }},
            {s: { c: 2, r: (16 + 27*i) }, e: { c: 2, r: (21 + 27*i) }},
            {s: { c: 2, r: (23 + 27*i) }, e: { c: 2, r: (28 + 27*i) }},
            {s: { c: 2, r: (30 + 27*i) }, e: { c: 12, r: (30 + 27*i) }},
            {s: { c: 3, r: (7 + 27*i) }, e: { c: 3, r: (8 + 27*i) }},
            {s: { c: 4, r: (7 + 27*i) }, e: { c: 12, r: (7 + 27*i) }}
        ]
        rango = rango.concat(rangoAux);
        piso3['!merges'] = rango;

    }
    
    //Llenar la hoja de excel del hemeroteca
    cont = 0;
    contD = 0;
    rango = [{s: { c: 1, r: 5 }, e: { c: 12, r: 5 }}]; 
    hemeroteca['B' + 6] = { t:'s', v: text[0]};
    for (i = 0; i < x; i++) { 
        totalD = 0;
        aux = 7

        hemeroteca['B' + (8 + 20*i)] = { t:'s', v: text[1]};
        hemeroteca['B' + (10 + 20*i)] = { t:'s', v: text[2]};
        hemeroteca['B' + (17 + 20*i)] = { t:'s', v: text[3]};
        hemeroteca['B' + (24 + 20*i)] = { t:'s', v: text[14]};
        hemeroteca['C' + (8 + 20*i)] = { t:'s', v: text[5]};
        hemeroteca['D' + (8 + 20*i)] = { t:'s', v: text[6]};
        hemeroteca['E' + (8 + 20*i)] = { t:'s', v: meses[i]};
        hemeroteca['E' + (9 + 20*i)] = { t:'s', v: text[13]};
        hemeroteca['F' + (9 + 20*i)] = { t:'s', v: text[6]};
        hemeroteca['G' + (9 + 20*i)] = { t:'s', v: text[13]};
        hemeroteca['H' + (9 + 20*i)] = { t:'s', v: text[6]};
        hemeroteca['I' + (9 + 20*i)] = { t:'s', v: text[13]};
        hemeroteca['J' + (9 + 20*i)] = { t:'s', v: text[6]};
        hemeroteca['K' + (9 + 20*i)] = { t:'s', v: text[13]};
        hemeroteca['L' + (9 + 20*i)] = { t:'s', v: text[6]};
        hemeroteca['M' + (9 + 20*i)] = { t:'s', v: text[13]};

        for (j = 10; j < 23; j++) {

            if(j == 16){ j++ }
            
            //Dias de la semana
            hemeroteca['D' + (j + 20*i)] = { t:'s', v: text[aux]};
            hemeroteca['F' + (j + 20*i)] = { t:'s', v: text[aux]};
            hemeroteca['H' + (j + 20*i)] = { t:'s', v: text[aux]};
            hemeroteca['J' + (j + 20*i)] = { t:'s', v: text[aux]};
            hemeroteca['L' + (j + 20*i)] = { t:'s', v: text[aux]};

            //cantidad de usuarios en un dia
            hemeroteca['E' + (j + 20*i)] = { t:'n', v: c[(0 + cont)]};
            hemeroteca['G' + (j + 20*i)] = { t:'n', v: c[(6 + cont)]};
            hemeroteca['I' + (j + 20*i)] = { t:'n', v: c[(12 + cont)]};
            hemeroteca['K' + (j + 20*i)] = { t:'n', v: c[(18 + cont)]};
            hemeroteca['M' + (j + 20*i)] = { t:'n', v: c[(24 + cont)]};
            
            //Sumatoria de los dias
            totalD = totalD + c[(0 + cont)] + c[(6 + cont)] + c[(12 + cont)] + c[(18 + cont)] + c[(24 + cont)];

            if(aux < 12){
                aux++;
            }else{
                aux = 7;
            }

            if(cont < 5){
                cont++;
            }else{
                cont=0;
                if(contD < 1){
                    contD++;
                }else{
                    contD=0;
                }
                c = cantDias[i][contD];
            }

        }

        //Total de usuarios ese mes
        hemeroteca['C' + (24 + 20*i)] = { t:'n', v: totalD};

        //combinar celdas
        rangoAux = [
            {s: { c: 1, r: (6 + 20*i) }, e: { c: 12, r: (6 + 20*i) }},
            {s: { c: 1, r: (15 + 20*i) }, e: { c: 12, r: (15 + 20*i) }},
            {s: { c: 1, r: (7 + 20*i) }, e: { c: 1, r: (8 + 20*i) }},
            {s: { c: 1, r: (9 + 20*i) }, e: { c: 1, r: (14 + 20*i) }},
            {s: { c: 1, r: (16 + 20*i) }, e: { c: 1, r: (21 + 20*i) }},
            {s: { c: 1, r: (22 + 20*i) }, e: { c: 12, r: (22 + 20*i) }},
            {s: { c: 2, r: (7 + 20*i) }, e: { c: 2, r: (8 + 20*i) }},
            {s: { c: 2, r: (9 + 20*i) }, e: { c: 2, r: (14 + 20*i) }},
            {s: { c: 2, r: (16 + 20*i) }, e: { c: 2, r: (21 + 20*i) }},
            {s: { c: 2, r: (23 + 20*i) }, e: { c: 12, r: (23 + 20*i) }},
            {s: { c: 3, r: (7 + 20*i) }, e: { c: 3, r: (8 + 20*i) }},
            {s: { c: 4, r: (7 + 20*i) }, e: { c: 12, r: (7 + 20*i) }}
        ]
        rango = rango.concat(rangoAux);
        hemeroteca['!merges'] = rango;

    }

    workbook.SheetNames.push("Piso1");
    workbook.SheetNames.push("Piso3");
    workbook.SheetNames.push("Hemeroteca");
    workbook.Sheets["Piso1"] = piso1;
    workbook.Sheets["Piso3"] = piso3;
    workbook.Sheets["Hemeroteca"] = hemeroteca;

    XLSX.writeFile(workbook, 'SalasDeEstudio.xls');
};