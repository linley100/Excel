import * as XLSX from "xlsx";

export function imprimirUsuarios(mes: any[], diasT1: any[][][], diasT2: any[][][], diasT3: any[][][]) {
    // este el libro de trabajo
    let workbook = XLSX.utils.book_new();
    
    //Auxiliares y contadores
    let aux=0, i=0, j=0, cont=0, contD=0;
    let totalD=0; //Total de usuarios en un mes

    //variable tipo array de entero, contiene los meses
    let meses = mes;

    //Variable que guarda la cantidad de meses
    let x = meses.length;

    //Variable tipo arrays of arrrays of arrays de int, 
    //contiene los usuarios de un dia de un turno de un mes(Mes x Turno X dia) 
    let cantDiasT1 = diasT1; //Piso 1
    let cantDiasT2 = diasT2; //Piso 3
    let cantDiasT3 = diasT3; //Hemeroteca
    /*let cantDias = [
        [
            [100,200,300,400,500,600,
            100,200,300,400,500,600,
            100,200,300,400,500,600,
            100,200,300,400,500,600,
            100,200,300,400,500,600,
            100,200,300,400,500,600],
            [10,20,30,40,50,60,
            10,20,30,40,50,60,
            10,20,30,40,50,60,
            10,20,30,40,50,60,
            10,20,30,40,50,60,
            10,20,30,40,50,60],
            [1,2,3,4,5,6,
            1,2,3,4,5,6,
            1,2,3,4,5,6,
            1,2,3,4,5,6,
            1,2,3,4,5,6,
            1,2,3,4,5,6]
        ],
        [
            [100,200,300,400,500,600,
            100,200,300,400,500,600,
            100,200,300,400,500,600,
            100,200,300,400,500,600,
            100,200,300,400,500,600,
            100,200,300,400,500,600],
            [10,20,30,40,50,60,
            10,20,30,40,50,60,
            10,20,30,40,50,60,
            10,20,30,40,50,60,
            10,20,30,40,50,60,
            10,20,30,40,50,60],
            [1,2,3,4,5,6,
            1,2,3,4,5,6,
            1,2,3,4,5,6,
            1,2,3,4,5,6,
            1,2,3,4,5,6,
            1,2,3,4,5,6]
        ]
    ];*/

    //Celdas de la hoja de calculo
    let ws_data = [];

    //Inicializar las celdas a usar
    for (i = 0; i < (34 * x); i++) { 
        ws_data[i] = ['','','','','','','','','','','','','','','','','',''];
    }

    //Crean las hojas de calculo de los 3 pisos
    let piso1 = XLSX.utils.aoa_to_sheet(ws_data);
    let piso3 = XLSX.utils.aoa_to_sheet(ws_data);
    let hemeroteca = XLSX.utils.aoa_to_sheet(ws_data);

    //Variable para combinar las celdas
    let rango = [{s: { c: 1, r: 5 }, e: { c: 14, r: 5 }}]; 
    let rangoAux; //Auxiliar para meter nuevos rangos en rango

    //Array con palabras usadas
    let text = [
        'CANTIDAD DE USUARIOS SALAS DE ESTUDIOS',
        'TURNO','1er turno','2do turno','3er turno','RESPONSABLE',
        'FECHA','LUN','MART','MIER','JUEV','VIER','SAB','CANT','TOTAL'
    ];

    //Array con los nombres de los meses
    let mesesN = [
        'Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio',
        'Agosto','Septiembre','Otucbre','Noviembre','Diciembre'
    ];

    //Llenar la hoja de excel del piso 1
    piso1['B' + 6] = { t:'s', v: text[0]}; //Escribir titulo
    for (i = 0; i < x; i++) { 
	    //Variables auxiliares, dias por turno ( Turno x Dia )
	    let a = cantDiasT1[i][0]; //Piso 1
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
        piso1['E' + (8 + 27*i)] = { t:'s', v: mesesN[(meses[i] - 1)]};
        piso1['E' + (9 + 27*i)] = { t:'s', v: text[13]};
        piso1['F' + (9 + 27*i)] = { t:'s', v: text[6]};
        piso1['G' + (9 + 27*i)] = { t:'s', v: text[13]};
        piso1['H' + (9 + 27*i)] = { t:'s', v: text[6]};
        piso1['I' + (9 + 27*i)] = { t:'s', v: text[13]};
        piso1['J' + (9 + 27*i)] = { t:'s', v: text[6]};
        piso1['K' + (9 + 27*i)] = { t:'s', v: text[13]};
        piso1['L' + (9 + 27*i)] = { t:'s', v: text[6]};
        piso1['M' + (9 + 27*i)] = { t:'s', v: text[13]};
        piso1['N' + (9 + 27*i)] = { t:'s', v: text[6]};
        piso1['O' + (9 + 27*i)] = { t:'s', v: text[13]};

        //Escribir dias del mes y la cantidad de usuarios
        for (j = 10; j < 30; j++) {
            
            if(j == 16){ j++ }
            if(j == 23){ j++ }
            
            //Dias de la semana
            piso1['D' + (j + 27*i)] = { t:'s', v: text[aux]};
            piso1['F' + (j + 27*i)] = { t:'s', v: text[aux]};
            piso1['H' + (j + 27*i)] = { t:'s', v: text[aux]};
            piso1['J' + (j + 27*i)] = { t:'s', v: text[aux]};
            piso1['L' + (j + 27*i)] = { t:'s', v: text[aux]};
            piso1['N' + (j + 27*i)] = { t:'s', v: text[aux]};

            //cantidad de usuarios en un dia
            piso1['E' + (j + 27*i)] = { t:'n', v: a[(0 + cont)]};
            piso1['G' + (j + 27*i)] = { t:'n', v: a[(6 + cont)]};
            piso1['I' + (j + 27*i)] = { t:'n', v: a[(12 + cont)]};
            piso1['K' + (j + 27*i)] = { t:'n', v: a[(18 + cont)]};
            piso1['M' + (j + 27*i)] = { t:'n', v: a[(24 + cont)]};
            piso1['O' + (j + 27*i)] = { t:'n', v: a[(30 + cont)]};
            
            //Sumatoria de los usuarios
            totalD = totalD + a[(0 + cont)] + a[(6 + cont)] + a[(12 + cont)] + a[(18 + cont)] + a[(24 + cont)] + a[(30 + cont)];
            
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
                a = cantDiasT1[i][contD];
            }
            
        }

        //Total de usuarios ese mes
        piso1['C' + (31 + 27*i)] = { t:'n', v: totalD};

        //combinar celdas, 's' es la celda inicial y 'e' es la celda final 
        //'c' es la columna y 'r' es la fila. (A1 esta en la posicion '0,0', es decir, c:0 r:0)
        rangoAux = [
            {s: { c: 1, r: (6 + 27*i) }, e: { c: 14, r: (6 + 27*i) }},
            {s: { c: 1, r: (15 + 27*i) }, e: { c: 14, r: (15 + 27*i) }},
            {s: { c: 1, r: (22 + 27*i) }, e: { c: 14, r: (22 + 27*i) }},
            {s: { c: 1, r: (29 + 27*i) }, e: { c: 14, r: (29 + 27*i) }},
            {s: { c: 1, r: (7 + 27*i) }, e: { c: 1, r: (8 + 27*i) }},
            {s: { c: 1, r: (9 + 27*i) }, e: { c: 1, r: (14 + 27*i) }},
            {s: { c: 1, r: (16 + 27*i) }, e: { c: 1, r: (21 + 27*i) }},
            {s: { c: 1, r: (23 + 27*i) }, e: { c: 1, r: (28 + 27*i) }},
            {s: { c: 2, r: (7 + 27*i) }, e: { c: 2, r: (8 + 27*i) }},
            {s: { c: 2, r: (9 + 27*i) }, e: { c: 2, r: (14 + 27*i) }},
            {s: { c: 2, r: (16 + 27*i) }, e: { c: 2, r: (21 + 27*i) }},
            {s: { c: 2, r: (23 + 27*i) }, e: { c: 2, r: (28 + 27*i) }},
            {s: { c: 2, r: (30 + 27*i) }, e: { c: 14, r: (30 + 27*i) }},
            {s: { c: 3, r: (7 + 27*i) }, e: { c: 3, r: (8 + 27*i) }},
            {s: { c: 4, r: (7 + 27*i) }, e: { c: 14, r: (7 + 27*i) }}
        ]
        rango = rango.concat(rangoAux);
        piso1['!merges'] = rango; //"!merges" asigna que celdas se combinan es esa hoja
    }

    //Llenar la hoja de excel del piso 3
    cont = 0;
    contD = 0;
    rango = [{s: { c: 1, r: 5 }, e: { c: 14, r: 5 }}]; 
    piso3['B' + 6] = { t:'s', v: text[0]};
    for (i = 0; i < x; i++) { 
    	//Variables auxiliares, dias por turno ( Turno x Dia )
	    let b = cantDiasT2[i][0]; //Piso 3
        totalD = 0;
        aux = 7

        piso3['B' + (8 + 27*i)] = { t:'s', v: text[1]};
        piso3['B' + (10 + 27*i)] = { t:'s', v: text[2]};
        piso3['B' + (17 + 27*i)] = { t:'s', v: text[3]};
        piso3['B' + (24 + 27*i)] = { t:'s', v: text[4]};
        piso3['B' + (31 + 27*i)] = { t:'s', v: text[14]};
        piso3['C' + (8 + 27*i)] = { t:'s', v: text[5]};
        piso3['D' + (8 + 27*i)] = { t:'s', v: text[6]};
        piso3['E' + (8 + 27*i)] = { t:'s', v: mesesN[(meses[i] - 1)]};
        piso3['E' + (9 + 27*i)] = { t:'s', v: text[13]};
        piso3['F' + (9 + 27*i)] = { t:'s', v: text[6]};
        piso3['G' + (9 + 27*i)] = { t:'s', v: text[13]};
        piso3['H' + (9 + 27*i)] = { t:'s', v: text[6]};
        piso3['I' + (9 + 27*i)] = { t:'s', v: text[13]};
        piso3['J' + (9 + 27*i)] = { t:'s', v: text[6]};
        piso3['K' + (9 + 27*i)] = { t:'s', v: text[13]};
        piso3['L' + (9 + 27*i)] = { t:'s', v: text[6]};
        piso3['M' + (9 + 27*i)] = { t:'s', v: text[13]};
        piso3['N' + (9 + 27*i)] = { t:'s', v: text[6]};
        piso3['O' + (9 + 27*i)] = { t:'s', v: text[13]};

        for (j = 10; j < 30; j++) {

            if(j == 16){ j++ }
            if(j == 23){ j++ }
            
            //Dias de la semana
            piso3['D' + (j + 27*i)] = { t:'s', v: text[aux]};
            piso3['F' + (j + 27*i)] = { t:'s', v: text[aux]};
            piso3['H' + (j + 27*i)] = { t:'s', v: text[aux]};
            piso3['J' + (j + 27*i)] = { t:'s', v: text[aux]};
            piso3['L' + (j + 27*i)] = { t:'s', v: text[aux]};
            piso3['N' + (j + 27*i)] = { t:'s', v: text[aux]};

            //cantidad de usuarios en un dia
            piso3['E' + (j + 27*i)] = { t:'n', v: b[(0 + cont)]};
            piso3['G' + (j + 27*i)] = { t:'n', v: b[(6 + cont)]};
            piso3['I' + (j + 27*i)] = { t:'n', v: b[(12 + cont)]};
            piso3['K' + (j + 27*i)] = { t:'n', v: b[(18 + cont)]};
            piso3['M' + (j + 27*i)] = { t:'n', v: b[(24 + cont)]};
            piso3['O' + (j + 27*i)] = { t:'n', v: b[(24 + cont)]};
           
            //Sumatoria de los usuarios
            totalD = totalD + b[(0 + cont)] + b[(6 + cont)] + b[(12 + cont)] + b[(18 + cont)] + b[(24 + cont)] + b[(30 + cont)];

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
                b = cantDiasT2[i][contD];
            }

        }

        //Total de usuarios ese mes
        piso3['C' + (31 + 27*i)] = { t:'n', v: totalD};

        //combinar celdas
        rangoAux = [
            {s: { c: 1, r: (6 + 27*i) }, e: { c: 14, r: (6 + 27*i) }},
            {s: { c: 1, r: (15 + 27*i) }, e: { c: 14, r: (15 + 27*i) }},
            {s: { c: 1, r: (22 + 27*i) }, e: { c: 14, r: (22 + 27*i) }},
            {s: { c: 1, r: (29 + 27*i) }, e: { c: 14, r: (29 + 27*i) }},
            {s: { c: 1, r: (7 + 27*i) }, e: { c: 1, r: (8 + 27*i) }},
            {s: { c: 1, r: (9 + 27*i) }, e: { c: 1, r: (14 + 27*i) }},
            {s: { c: 1, r: (16 + 27*i) }, e: { c: 1, r: (21 + 27*i) }},
            {s: { c: 1, r: (23 + 27*i) }, e: { c: 1, r: (28 + 27*i) }},
            {s: { c: 2, r: (7 + 27*i) }, e: { c: 2, r: (8 + 27*i) }},
            {s: { c: 2, r: (9 + 27*i) }, e: { c: 2, r: (14 + 27*i) }},
            {s: { c: 2, r: (16 + 27*i) }, e: { c: 2, r: (21 + 27*i) }},
            {s: { c: 2, r: (23 + 27*i) }, e: { c: 2, r: (28 + 27*i) }},
            {s: { c: 2, r: (30 + 27*i) }, e: { c: 14, r: (30 + 27*i) }},
            {s: { c: 3, r: (7 + 27*i) }, e: { c: 3, r: (8 + 27*i) }},
            {s: { c: 4, r: (7 + 27*i) }, e: { c: 14, r: (7 + 27*i) }}
        ]
        rango = rango.concat(rangoAux);
        piso3['!merges'] = rango;

    }
    
    //Llenar la hoja de excel del hemeroteca
    cont = 0;
    contD = 0;
    rango = [{s: { c: 1, r: 5 }, e: { c: 14, r: 5 }}]; 
    hemeroteca['B' + 6] = { t:'s', v: text[0]};
    for (i = 0; i < x; i++) { 
	    //Variables auxiliares, dias por turno ( Turno x Dia )
	    let c = cantDiasT3[i][0]; //Hemeroteca
        totalD = 0;
        aux = 7

        hemeroteca['B' + (8 + 20*i)] = { t:'s', v: text[1]};
        hemeroteca['B' + (10 + 20*i)] = { t:'s', v: text[2]};
        hemeroteca['B' + (17 + 20*i)] = { t:'s', v: text[3]};
        hemeroteca['B' + (24 + 20*i)] = { t:'s', v: text[14]};
        hemeroteca['C' + (8 + 20*i)] = { t:'s', v: text[5]};
        hemeroteca['D' + (8 + 20*i)] = { t:'s', v: text[6]};
        hemeroteca['E' + (8 + 20*i)] = { t:'s', v: mesesN[(meses[i] - 1)]};
        hemeroteca['E' + (9 + 20*i)] = { t:'s', v: text[13]};
        hemeroteca['F' + (9 + 20*i)] = { t:'s', v: text[6]};
        hemeroteca['G' + (9 + 20*i)] = { t:'s', v: text[13]};
        hemeroteca['H' + (9 + 20*i)] = { t:'s', v: text[6]};
        hemeroteca['I' + (9 + 20*i)] = { t:'s', v: text[13]};
        hemeroteca['J' + (9 + 20*i)] = { t:'s', v: text[6]};
        hemeroteca['K' + (9 + 20*i)] = { t:'s', v: text[13]};
        hemeroteca['L' + (9 + 20*i)] = { t:'s', v: text[6]};
        hemeroteca['M' + (9 + 20*i)] = { t:'s', v: text[13]};
        hemeroteca['N' + (9 + 20*i)] = { t:'s', v: text[6]};
        hemeroteca['O' + (9 + 20*i)] = { t:'s', v: text[13]};

        for (j = 10; j < 23; j++) {

            if(j == 16){ j++ }
            
            //Dias de la semana
            hemeroteca['D' + (j + 20*i)] = { t:'s', v: text[aux]};
            hemeroteca['F' + (j + 20*i)] = { t:'s', v: text[aux]};
            hemeroteca['H' + (j + 20*i)] = { t:'s', v: text[aux]};
            hemeroteca['J' + (j + 20*i)] = { t:'s', v: text[aux]};
            hemeroteca['L' + (j + 20*i)] = { t:'s', v: text[aux]};
            hemeroteca['N' + (j + 20*i)] = { t:'s', v: text[aux]};

            //cantidad de usuarios en un dia
            hemeroteca['E' + (j + 20*i)] = { t:'n', v: c[(0 + cont)]};
            hemeroteca['G' + (j + 20*i)] = { t:'n', v: c[(6 + cont)]};
            hemeroteca['I' + (j + 20*i)] = { t:'n', v: c[(12 + cont)]};
            hemeroteca['K' + (j + 20*i)] = { t:'n', v: c[(18 + cont)]};
            hemeroteca['M' + (j + 20*i)] = { t:'n', v: c[(24 + cont)]};
            hemeroteca['O' + (j + 20*i)] = { t:'n', v: c[(30 + cont)]};
            
            //Sumatoria de los dias
            totalD = totalD + c[(0 + cont)] + c[(6 + cont)] + c[(12 + cont)] + c[(18 + cont)] + c[(24 + cont)] + c[(30 + cont)];

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
                c = cantDiasT3[i][contD];
            }

        }

        //Total de usuarios ese mes
        hemeroteca['C' + (24 + 20*i)] = { t:'n', v: totalD};

        //combinar celdas
        rangoAux = [
            {s: { c: 1, r: (6 + 20*i) }, e: { c: 14, r: (6 + 20*i) }},
            {s: { c: 1, r: (15 + 20*i) }, e: { c: 14, r: (15 + 20*i) }},
            {s: { c: 1, r: (7 + 20*i) }, e: { c: 1, r: (8 + 20*i) }},
            {s: { c: 1, r: (9 + 20*i) }, e: { c: 1, r: (14 + 20*i) }},
            {s: { c: 1, r: (16 + 20*i) }, e: { c: 1, r: (21 + 20*i) }},
            {s: { c: 1, r: (22 + 20*i) }, e: { c: 14, r: (22 + 20*i) }},
            {s: { c: 2, r: (7 + 20*i) }, e: { c: 2, r: (8 + 20*i) }},
            {s: { c: 2, r: (9 + 20*i) }, e: { c: 2, r: (14 + 20*i) }},
            {s: { c: 2, r: (16 + 20*i) }, e: { c: 2, r: (21 + 20*i) }},
            {s: { c: 2, r: (23 + 20*i) }, e: { c: 14, r: (23 + 20*i) }},
            {s: { c: 3, r: (7 + 20*i) }, e: { c: 3, r: (8 + 20*i) }},
            {s: { c: 4, r: (7 + 20*i) }, e: { c: 14, r: (7 + 20*i) }}
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

export function imprimirSalas(mes: any[], diasT1: any[][][], diasT2: any[][][], diasT3: any[][][]) {
    // este el libro de trabajo
    let workbook = XLSX.utils.book_new();
    
    //Auxiliares y contadores
    let aux=0, i=0, j=0, cont=0, contD=0, totalD=0;

    //variable tipo array de string, contiene los meses
    let meses = mes;
    //Variable que guarda la cantidad de meses
    let x = meses.length;
    //Variable tipo arrays of arrrays of arrays de int, contiene los dias de un turno de un mes
    let cantDiasT1 =  diasT1;
    let cantDiasT2 =  diasT2;
    let cantDiasT3 =  diasT3;
    /*let cantDias = [
        [
            [100,200,300,400,500,600,
            100,200,300,400,500,600,
            100,200,300,400,500,600,
            100,200,300,400,500,600,
            100,200,300,400,500,600,
            100,200,300,400,500,600],
            [10,20,30,40,50,60,
            10,20,30,40,50,60,
            10,20,30,40,50,60,
            10,20,30,40,50,60,
            10,20,30,40,50,60,
            10,20,30,40,50,60],
            [1,2,3,4,5,6,
            1,2,3,4,5,6,
            1,2,3,4,5,6,
            1,2,3,4,5,6,
            1,2,3,4,5,6,
            1,2,3,4,5,6]
        ],
        [
            [100,200,300,400,500,600,
            100,200,300,400,500,600,
            100,200,300,400,500,600,
            100,200,300,400,500,600,
            100,200,300,400,500,600,
            100,200,300,400,500,600],
            [10,20,30,40,50,60,
            10,20,30,40,50,60,
            10,20,30,40,50,60,
            10,20,30,40,50,60,
            10,20,30,40,50,60,
            10,20,30,40,50,60],
            [1,2,3,4,5,6,
            1,2,3,4,5,6,
            1,2,3,4,5,6,
            1,2,3,4,5,6,
            1,2,3,4,5,6,
            1,2,3,4,5,6]
        ]
    ];*/

    //Celdas de la hoja de calculo
    let ws_data = [];

    //Inicializar las celdas a usar
    for (i = 0; i < (34 * x); i++) { 
        ws_data[i] = ['','','','','','','','','','','','','','','','',''];
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

    //Array con los nombres de los meses
    let mesesN = [
        'Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio',
        'Agosto','Septiembre','Otucbre','Noviembre','Diciembre'
    ];

    //Llenar la hoja de excel del piso 1
    piso1['B' + 6] = { t:'s', v: text[0]}; //Escribir titulo
    for (i = 0; i < x; i++) { 
	    //Variables auxiliares, dias por turno ( Turno x Dia )
	    let a = cantDiasT1[i][0]; //Piso 1
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
        piso1['E' + (8 + 27*i)] = { t:'s', v: mesesN[(meses[i] - 1)]};
        piso1['E' + (9 + 27*i)] = { t:'s', v: text[13]};
        piso1['F' + (9 + 27*i)] = { t:'s', v: text[6]};
        piso1['G' + (9 + 27*i)] = { t:'s', v: text[13]};
        piso1['H' + (9 + 27*i)] = { t:'s', v: text[6]};
        piso1['I' + (9 + 27*i)] = { t:'s', v: text[13]};
        piso1['J' + (9 + 27*i)] = { t:'s', v: text[6]};
        piso1['K' + (9 + 27*i)] = { t:'s', v: text[13]};
        piso1['L' + (9 + 27*i)] = { t:'s', v: text[6]};
        piso1['M' + (9 + 27*i)] = { t:'s', v: text[13]};
        piso1['N' + (9 + 27*i)] = { t:'s', v: text[6]};
        piso1['O' + (9 + 27*i)] = { t:'s', v: text[13]};

        //Escribir dias del mes y la cantidad de usuarios
        for (j = 10; j < 30; j++) {
            
            if(j == 16){ j++ }
            if(j == 23){ j++ }
            
            //Dias de la semana
            piso1['D' + (j + 27*i)] = { t:'s', v: text[aux]};
            piso1['F' + (j + 27*i)] = { t:'s', v: text[aux]};
            piso1['H' + (j + 27*i)] = { t:'s', v: text[aux]};
            piso1['J' + (j + 27*i)] = { t:'s', v: text[aux]};
            piso1['L' + (j + 27*i)] = { t:'s', v: text[aux]};
            piso1['N' + (j + 27*i)] = { t:'s', v: text[aux]};

            //cantidad de salas reservadas en un dia
            piso1['E' + (j + 27*i)] = { t:'n', v: a[(0 + cont)]};
            piso1['G' + (j + 27*i)] = { t:'n', v: a[(6 + cont)]};
            piso1['I' + (j + 27*i)] = { t:'n', v: a[(12 + cont)]};
            piso1['K' + (j + 27*i)] = { t:'n', v: a[(18 + cont)]};
            piso1['M' + (j + 27*i)] = { t:'n', v: a[(24 + cont)]};
            piso1['O' + (j + 27*i)] = { t:'n', v: a[(30 + cont)]};
            
            //Sumatoria de las salas reservadas
            totalD = totalD + a[(0 + cont)] + a[(6 + cont)] + a[(12 + cont)] + a[(18 + cont)] + a[(24 + cont)] + a[(30 + cont)];
            
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
                a = cantDiasT1[i][contD];
            }
            
        }

        //Total de salas reservadas ese mes
        piso1['C' + (31 + 27*i)] = { t:'n', v: totalD};

        //combinar celdas, 's' es la celda inicial y 'e' es la celda final 
        //'c' es la columna y 'r' es la fila. (A1 esta en la posicion '0,0', es decir, c:0 r:0)
        rangoAux = [
            {s: { c: 1, r: (6 + 27*i) }, e: { c: 14, r: (6 + 27*i) }},
            {s: { c: 1, r: (15 + 27*i) }, e: { c: 14, r: (15 + 27*i) }},
            {s: { c: 1, r: (22 + 27*i) }, e: { c: 14, r: (22 + 27*i) }},
            {s: { c: 1, r: (29 + 27*i) }, e: { c: 14, r: (29 + 27*i) }},
            {s: { c: 1, r: (7 + 27*i) }, e: { c: 1, r: (8 + 27*i) }},
            {s: { c: 1, r: (9 + 27*i) }, e: { c: 1, r: (14 + 27*i) }},
            {s: { c: 1, r: (16 + 27*i) }, e: { c: 1, r: (21 + 27*i) }},
            {s: { c: 1, r: (23 + 27*i) }, e: { c: 1, r: (28 + 27*i) }},
            {s: { c: 2, r: (7 + 27*i) }, e: { c: 2, r: (8 + 27*i) }},
            {s: { c: 2, r: (9 + 27*i) }, e: { c: 2, r: (14 + 27*i) }},
            {s: { c: 2, r: (16 + 27*i) }, e: { c: 2, r: (21 + 27*i) }},
            {s: { c: 2, r: (23 + 27*i) }, e: { c: 2, r: (28 + 27*i) }},
            {s: { c: 2, r: (30 + 27*i) }, e: { c: 14, r: (30 + 27*i) }},
            {s: { c: 3, r: (7 + 27*i) }, e: { c: 3, r: (8 + 27*i) }},
            {s: { c: 4, r: (7 + 27*i) }, e: { c: 14, r: (7 + 27*i) }}
        ]
        rango = rango.concat(rangoAux);
        piso1['!merges'] = rango; //"!merges" asigna que celdas se combinan es esa hoja
    }

    //Llenar la hoja de excel del piso 3
    cont = 0;
    contD = 0;
    rango = [{s: { c: 1, r: 5 }, e: { c: 14, r: 5 }}]; 
    piso3['B' + 6] = { t:'s', v: text[0]};
    for (i = 0; i < x; i++) { 
	    //Variables auxiliares, dias por turno ( Turno x Dia )
	    let b = cantDiasT2[i][0]; //Piso 3
        totalD = 0;
        aux = 7

        piso3['B' + (8 + 27*i)] = { t:'s', v: text[1]};
        piso3['B' + (10 + 27*i)] = { t:'s', v: text[2]};
        piso3['B' + (17 + 27*i)] = { t:'s', v: text[3]};
        piso3['B' + (24 + 27*i)] = { t:'s', v: text[4]};
        piso3['B' + (31 + 27*i)] = { t:'s', v: text[14]};
        piso3['C' + (8 + 27*i)] = { t:'s', v: text[5]};
        piso3['D' + (8 + 27*i)] = { t:'s', v: text[6]};
        piso3['E' + (8 + 27*i)] = { t:'s', v: mesesN[(meses[i] - 1)]};
        piso3['E' + (9 + 27*i)] = { t:'s', v: text[13]};
        piso3['F' + (9 + 27*i)] = { t:'s', v: text[6]};
        piso3['G' + (9 + 27*i)] = { t:'s', v: text[13]};
        piso3['H' + (9 + 27*i)] = { t:'s', v: text[6]};
        piso3['I' + (9 + 27*i)] = { t:'s', v: text[13]};
        piso3['J' + (9 + 27*i)] = { t:'s', v: text[6]};
        piso3['K' + (9 + 27*i)] = { t:'s', v: text[13]};
        piso3['L' + (9 + 27*i)] = { t:'s', v: text[6]};
        piso3['M' + (9 + 27*i)] = { t:'s', v: text[13]};
        piso3['N' + (9 + 27*i)] = { t:'s', v: text[6]};
        piso3['O' + (9 + 27*i)] = { t:'s', v: text[13]};

        for (j = 10; j < 30; j++) {

            if(j == 16){ j++ }
            if(j == 23){ j++ }
            
            //Dias de la semana
            piso3['D' + (j + 27*i)] = { t:'s', v: text[aux]};
            piso3['F' + (j + 27*i)] = { t:'s', v: text[aux]};
            piso3['H' + (j + 27*i)] = { t:'s', v: text[aux]};
            piso3['J' + (j + 27*i)] = { t:'s', v: text[aux]};
            piso3['L' + (j + 27*i)] = { t:'s', v: text[aux]};
            piso3['N' + (j + 27*i)] = { t:'s', v: text[aux]};

            //cantidad de salas reservadas en un dia
            piso3['E' + (j + 27*i)] = { t:'n', v: b[(0 + cont)]};
            piso3['G' + (j + 27*i)] = { t:'n', v: b[(6 + cont)]};
            piso3['I' + (j + 27*i)] = { t:'n', v: b[(12 + cont)]};
            piso3['K' + (j + 27*i)] = { t:'n', v: b[(18 + cont)]};
            piso3['M' + (j + 27*i)] = { t:'n', v: b[(24 + cont)]};
            piso3['O' + (j + 27*i)] = { t:'n', v: b[(24 + cont)]};
           
            //Sumatoria de las salas reservadas
            totalD = totalD + b[(0 + cont)] + b[(6 + cont)] + b[(12 + cont)] + b[(18 + cont)] + b[(24 + cont)] + b[(30 + cont)];

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
                b = cantDiasT2[i][contD];
            }

        }

        //Total de salas reservadas ese mes
        piso3['C' + (31 + 27*i)] = { t:'n', v: totalD};

        //combinar celdas
        rangoAux = [
            {s: { c: 1, r: (6 + 27*i) }, e: { c: 14, r: (6 + 27*i) }},
            {s: { c: 1, r: (15 + 27*i) }, e: { c: 14, r: (15 + 27*i) }},
            {s: { c: 1, r: (22 + 27*i) }, e: { c: 14, r: (22 + 27*i) }},
            {s: { c: 1, r: (29 + 27*i) }, e: { c: 14, r: (29 + 27*i) }},
            {s: { c: 1, r: (7 + 27*i) }, e: { c: 1, r: (8 + 27*i) }},
            {s: { c: 1, r: (9 + 27*i) }, e: { c: 1, r: (14 + 27*i) }},
            {s: { c: 1, r: (16 + 27*i) }, e: { c: 1, r: (21 + 27*i) }},
            {s: { c: 1, r: (23 + 27*i) }, e: { c: 1, r: (28 + 27*i) }},
            {s: { c: 2, r: (7 + 27*i) }, e: { c: 2, r: (8 + 27*i) }},
            {s: { c: 2, r: (9 + 27*i) }, e: { c: 2, r: (14 + 27*i) }},
            {s: { c: 2, r: (16 + 27*i) }, e: { c: 2, r: (21 + 27*i) }},
            {s: { c: 2, r: (23 + 27*i) }, e: { c: 2, r: (28 + 27*i) }},
            {s: { c: 2, r: (30 + 27*i) }, e: { c: 14, r: (30 + 27*i) }},
            {s: { c: 3, r: (7 + 27*i) }, e: { c: 3, r: (8 + 27*i) }},
            {s: { c: 4, r: (7 + 27*i) }, e: { c: 14, r: (7 + 27*i) }}
        ]
        rango = rango.concat(rangoAux);
        piso3['!merges'] = rango;

    }
    
    //Llenar la hoja de excel del hemeroteca
    cont = 0;
    contD = 0;
    rango = [{s: { c: 1, r: 5 }, e: { c: 14, r: 5 }}]; 
    hemeroteca['B' + 6] = { t:'s', v: text[0]};
    for (i = 0; i < x; i++) { 
	    //Variables auxiliares, dias por turno ( Turno x Dia )
	    let c = cantDiasT3[i][0]; //Hemeroteca
        totalD = 0;
        aux = 7

        hemeroteca['B' + (8 + 20*i)] = { t:'s', v: text[1]};
        hemeroteca['B' + (10 + 20*i)] = { t:'s', v: text[2]};
        hemeroteca['B' + (17 + 20*i)] = { t:'s', v: text[3]};
        hemeroteca['B' + (24 + 20*i)] = { t:'s', v: text[14]};
        hemeroteca['C' + (8 + 20*i)] = { t:'s', v: text[5]};
        hemeroteca['D' + (8 + 20*i)] = { t:'s', v: text[6]};
        hemeroteca['E' + (8 + 20*i)] = { t:'s', v: mesesN[(meses[i] - 1)]};
        hemeroteca['E' + (9 + 20*i)] = { t:'s', v: text[13]};
        hemeroteca['F' + (9 + 20*i)] = { t:'s', v: text[6]};
        hemeroteca['G' + (9 + 20*i)] = { t:'s', v: text[13]};
        hemeroteca['H' + (9 + 20*i)] = { t:'s', v: text[6]};
        hemeroteca['I' + (9 + 20*i)] = { t:'s', v: text[13]};
        hemeroteca['J' + (9 + 20*i)] = { t:'s', v: text[6]};
        hemeroteca['K' + (9 + 20*i)] = { t:'s', v: text[13]};
        hemeroteca['L' + (9 + 20*i)] = { t:'s', v: text[6]};
        hemeroteca['M' + (9 + 20*i)] = { t:'s', v: text[13]};
        hemeroteca['N' + (9 + 20*i)] = { t:'s', v: text[6]};
        hemeroteca['O' + (9 + 20*i)] = { t:'s', v: text[13]};

        for (j = 10; j < 23; j++) {

            if(j == 16){ j++ }
            
            //Dias de la semana
            hemeroteca['D' + (j + 20*i)] = { t:'s', v: text[aux]};
            hemeroteca['F' + (j + 20*i)] = { t:'s', v: text[aux]};
            hemeroteca['H' + (j + 20*i)] = { t:'s', v: text[aux]};
            hemeroteca['J' + (j + 20*i)] = { t:'s', v: text[aux]};
            hemeroteca['L' + (j + 20*i)] = { t:'s', v: text[aux]};
            hemeroteca['N' + (j + 20*i)] = { t:'s', v: text[aux]};

            //cantidad de salas reservadas en un dia
            hemeroteca['E' + (j + 20*i)] = { t:'n', v: c[(0 + cont)]};
            hemeroteca['G' + (j + 20*i)] = { t:'n', v: c[(6 + cont)]};
            hemeroteca['I' + (j + 20*i)] = { t:'n', v: c[(12 + cont)]};
            hemeroteca['K' + (j + 20*i)] = { t:'n', v: c[(18 + cont)]};
            hemeroteca['M' + (j + 20*i)] = { t:'n', v: c[(24 + cont)]};
            hemeroteca['O' + (j + 20*i)] = { t:'n', v: c[(30 + cont)]};
            
            //Sumatoria de las salas reservadas
            totalD = totalD + c[(0 + cont)] + c[(6 + cont)] + c[(12 + cont)] + c[(18 + cont)] + c[(24 + cont)] + c[(30 + cont)];

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
                c = cantDiasT3[i][contD];
            }

        }

        //Total de salas reservadas ese mes
        hemeroteca['C' + (24 + 20*i)] = { t:'n', v: totalD};

        //combinar celdas
        rangoAux = [
            {s: { c: 1, r: (6 + 20*i) }, e: { c: 14, r: (6 + 20*i) }},
            {s: { c: 1, r: (15 + 20*i) }, e: { c: 14, r: (15 + 20*i) }},
            {s: { c: 1, r: (7 + 20*i) }, e: { c: 1, r: (8 + 20*i) }},
            {s: { c: 1, r: (9 + 20*i) }, e: { c: 1, r: (14 + 20*i) }},
            {s: { c: 1, r: (16 + 20*i) }, e: { c: 1, r: (21 + 20*i) }},
            {s: { c: 1, r: (22 + 20*i) }, e: { c: 14, r: (22 + 20*i) }},
            {s: { c: 2, r: (7 + 20*i) }, e: { c: 2, r: (8 + 20*i) }},
            {s: { c: 2, r: (9 + 20*i) }, e: { c: 2, r: (14 + 20*i) }},
            {s: { c: 2, r: (16 + 20*i) }, e: { c: 2, r: (21 + 20*i) }},
            {s: { c: 2, r: (23 + 20*i) }, e: { c: 14, r: (23 + 20*i) }},
            {s: { c: 3, r: (7 + 20*i) }, e: { c: 3, r: (8 + 20*i) }},
            {s: { c: 4, r: (7 + 20*i) }, e: { c: 14, r: (7 + 20*i) }}
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

export function imprimir(sanciones: any[][]) {
    // Este es el libro de trabajo
    let workbook = XLSX.utils.book_new();
    
    //Auxiliares y contadores
    let aux=0, i=0, j=0, cont=0, contD=0, totalD=0;

    //Celdas de la hoja de calculo
    let ws_data = [];

	//Cantidad de sancionados y sancionados
    let x = sanciones.length;
    let sancion = sanciones;
    /*let sancion = [
    	['Juan','Perez','28987654','01/02/19','01/03/19'],
    	['Juan1','Perez1','28987655','02/02/19','02/03/19'],
    	['Juan2','Perez2','28987656','03/02/19','03/03/19'],
    	['Juan3','Perez3','28987657','04/02/19','04/03/19']
    ];*/

    //Inicializar las celdas a usar
    for (i = 0; i < (8 + x); i++) { 
        ws_data[i] = ['','','','','','','','','','','','','','',''];
    }

    //Estas es la hoja de calculo
    let sancionados = XLSX.utils.aoa_to_sheet(ws_data);

    //Array con palabras usadas
    let text = [
        'USUARIOS SANCIONADOS',
        'Nombre','Apellido','C.I.','Inicio','Fin'
    ];


    //Llenar la hoja de excel de los sancionados
    sancionados['B' + 6] = { t:'s', v: text[0]};
    sancionados['B' + 7] = { t:'s', v: text[1]};
    sancionados['C' + 7] = { t:'s', v: text[2]};
    sancionados['D' + 7] = { t:'s', v: text[3]};
    sancionados['E' + 7] = { t:'s', v: text[4]};
    sancionados['F' + 7] = { t:'s', v: text[5]};
    for (i = 0; i < x; i++) { 
        
        //Escribir
        sancionados['B' + (8 + i)] = { t:'s', v: sancion[i][0]};
        sancionados['C' + (8 + i)] = { t:'s', v: sancion[i][1]};
        sancionados['D' + (8 + i)] = { t:'s', v: sancion[i][2]};
        sancionados['E' + (8 + i)] = { t:'s', v: sancion[i][3]};
        sancionados['F' + (8 + i)] = { t:'s', v: sancion[i][4]};

    }

    //combinar celdas
    //Variable para la para combinar las celdas
    let rango = [{s: { c: 1, r: 5 }, e: { c: 5, r: 5 }}];
    sancionados['!merges'] = rango;

    workbook.SheetNames.push("Sancionados");
    workbook.Sheets["Sancionados"] = sancionados;

    XLSX.writeFile(workbook, 'UsuariosDeLaSalaDeEstudio.xls');
};