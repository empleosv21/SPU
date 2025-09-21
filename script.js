// Variables globales
const fileInput = document.getElementById('file');
const btnLoad = document.getElementById('btnLoad');
const btnCalc = document.getElementById('btnCalc');
const btnExport = document.getElementById('btnExport');
const status = document.getElementById('status');
const preview = document.getElementById('preview');
const mesInput = document.getElementById('mesreporte');
const progressBar = document.getElementById('progressBar');
const progressBarInner = document.getElementById('progressBarInner');
const alertBox = document.getElementById('alertBox');
const tabButtons = document.querySelectorAll('.tab-button');

let workbook = null, sheetsData = {}, results = [];

const obsCatalog = {
    "00": "No tiene observacion",
    "01": "Sin cambios con respecto al mes anterior",
    "02": "Pagos adicionales",
    "03": "Aprendices",
    "04": "Pensionado",
    "05": "Licencia",
    "06": "Incapacidad",
    "07": "Retiro",
    "08": "Ingreso",
    "09": "Vacaciones",
    "10": "Vacaciones mas pagos adicionales",
    "12": "Subsidio",
    "13": "IPSFA"
};

// Funciones de utilidad
function showAlert(msg, type) { 
    alertBox.innerHTML = `<div class="alert alert-${type}">${msg}</div>`; 
    setTimeout(() => alertBox.innerHTML = '', 5000); 
}

function updateProgress(p) { 
    progressBar.style.display = 'block'; 
    p = Math.min(Math.max(p, 0), 100); 
    progressBarInner.style.width = p + '%'; 
}

function log(msg) { 
    const timestamp = new Date().toLocaleTimeString();
    status.textContent += `[${timestamp}] ${msg}\n`; 
    status.scrollTop = status.scrollHeight; 
}

function padLeft(s, len) { 
    s = String(s || ''); 
    while(s.length < len) s = '0' + s; 
    return s; 
}

function digitsOnly(s) { 
    return String(s || '').replace(/\D/g, ''); 
}

// Función para eliminar tildes de un texto
function removeAccents(text) {
    if (!text || typeof text !== 'string') return text;
    return text.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

function splitName(full) { 
    if(!full || typeof full !== 'string') return ['', '', '', '']; 
    full = removeAccents(full.trim().replace(/\s+/g, ' '));
    const parts = full.split(' '); 
    const c = ['de', 'del', 'la', 'los', 'las']; 
    let pn = '', sn = '', pa = '', sa = ''; 
    
    if(parts.length === 1) pn = parts[0]; 
    else if(parts.length === 2) { 
        pn = parts[0]; 
        pa = parts[1]; 
    } else if(parts.length === 3) { 
        pn = parts[0]; 
        if(c.includes(parts[1].toLowerCase())) 
            pa = parts[1] + ' ' + parts[2]; 
        else { 
            sn = parts[1]; 
            pa = parts[2]; 
        } 
    } else if(parts.length === 4) { 
        pn = parts[0]; 
        sn = parts[1]; 
        pa = parts[2]; 
        sa = parts[3]; 
    } else { 
        pn = parts[0]; 
        sn = parts[1]; 
        pa = parts[2]; 
        sa = parts.slice(3).join(' '); 
    } 
    return [pn, sn, pa, sa]; 
}

// busca valor en fila usando variantes de nombre de columna
function getField(row, variants) {
    variants = variants.map(v => v.toString().toUpperCase());
    for(const k of Object.keys(row)) {
        const up = k.toString().toUpperCase();
        for(const v of variants) {
            if(up === v || up.replace(/Ñ/g, 'N') === v.replace(/Ñ/g, 'N') || up.includes(v) || v.includes(up)) {
                return row[k];
            }
        }
    }
    // fallback: try direct keys
    for(const v of variants) if(row[v] !== undefined) return row[v];
    return undefined;
}

function toNumber(val) { 
    const n = Number(String(val).replace(/[^0-9\.\-]/g, '')); 
    return isNaN(n) ? 0 : n; 
}

// Manejar pestañas
tabButtons.forEach(button => {
    button.addEventListener('click', () => {
        // Desactivar todas las pestañas
        tabButtons.forEach(btn => btn.classList.remove('active'));
        document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));
        
        // Activar la pestaña seleccionada
        button.classList.add('active');
        document.getElementById(button.dataset.tab).classList.add('active');
    });
});

// Cargar Excel
btnLoad.addEventListener('click', () => {
    status.textContent = ''; 
    preview.innerHTML = '';
    
    if(!fileInput.files[0]) { 
        showAlert('Selecciona un archivo Excel primero.', 'error'); 
        return; 
    }
    
    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            workbook = XLSX.read(data, {type: 'array'});
            sheetsData = {};
            
            workbook.SheetNames.forEach(name => {
                sheetsData[name.toUpperCase()] = XLSX.utils.sheet_to_json(workbook.Sheets[name], {defval: ''});
            });
            
            log('Excel cargado: ' + workbook.SheetNames.join(', '));
            log('Filas en INGRESOS: ' + ((sheetsData['INGRESOS'] || sheetsData['ING'] || []).length));
            log('Filas en BASE: ' + ((sheetsData['BASE'] || []).length));
            log('Filas en ISSS: ' + ((sheetsData['ISSS'] || []).length));
            log('Filas en DUI: ' + ((sheetsData['DUI'] || []).length));
            log('Filas en SUBSIDIO: ' + ((sheetsData['SUBSIDIO'] || []).length));
            
            btnCalc.disabled = false;
            showAlert('Archivo cargado correctamente. Revise el registro de actividad.', 'success');
        } catch(err) {
            log('Error al cargar Excel: ' + err.message);
            showAlert('Error al cargar archivo: ' + err.message, 'error');
        }
    };
    reader.onerror = () => { 
        log('Error al leer archivo'); 
        showAlert('Error al leer el archivo', 'error'); 
    };
    reader.readAsArrayBuffer(fileInput.files[0]);
});

// función para computar monto de subsidio (intenta usar campos de la hoja SUBSIDIO; si faltan datos deja 0)
function computeSubsidioAmount(subRow, contextRow) {
    // buscar montos ya calculados en la fila de subsidio
    const montoDirect = getField(subRow, ['MONTO','MONTO_SUBSIDIO','PAGO','PAGO_SUBSIDIO','VALOR','SUBSIDIO','IMPORTE']);
    if(montoDirect !== undefined && montoDirect !== '') {
        const m = toNumber(montoDirect);
        log(`Subsidio: monto directo encontrado = ${m} para codigo ${getField(subRow,['CODIGO','COD'])||getField(subRow,['Codigo Empleado','Codigo'])}`);
        return m;
    }

    // intentar construir desde salario + dias + porcentaje
    let salario = getField(subRow, ['SALARIO','SALARIOS','SUELDO','SUELDO_MENSUAL','SALARIO_MENSUAL','SALARIO_PROMEDIO','BASE']);
    let dias = getField(subRow, ['DIAS_INCAPACIDAD','DIAS','DIAS_SUBSIDIO','DIAS_INCAP','DIAS_SUB']);
    let porcentaje = getField(subRow, ['PORCENTAJE','PORCENT','TASA','PORCENTAJE_PAGO','PORCENTAJE_SUBSIDIO']);

    // si los datos no estan en la fila de subsidio, intentar buscarlos en la fila de ingresos (contextRow)
    if((salario === undefined || salario === '') && contextRow) {
        salario = getField(contextRow, ['SALARIO','SALARIOS','SUELDO','SUELDO_MENSUAL','SALARIO_MENSUAL']);
    }
    if((dias === undefined || dias === '') && contextRow) {
        dias = getField(contextRow, ['DIAS_INCAPACIDAD','DIAS','DIAS_TRABAJADOS','DIAS_INCAP']);
    }

    const numSal = toNumber(salario);
    const numDias = Math.round(toNumber(dias));
    let pct = (porcentaje !== undefined && porcentaje !== '') ? toNumber(porcentaje)/100 : null;

    // detectar si salario es diario (nombres con 'DIARIO')
    let salarioIsDaily = false;
    for(const k of Object.keys(subRow)) {
        if(k.toString().toUpperCase().includes('DIARIO')) salarioIsDaily = true;
    }
    if(contextRow) {
        for(const k of Object.keys(contextRow)) {
            if(k.toString().toUpperCase().includes('DIARIO')) salarioIsDaily = true;
        }
    }

    if(numSal > 0 && numDias > 0) {
        const sueldoDiario = salarioIsDaily ? numSal : (numSal/30);
        if(pct === null) pct = 0.67; // por defecto 67% si no hay porcentaje (ajustable)
        const calc = Number((sueldoDiario * numDias * pct).toFixed(2));
        log(`Subsidio calculado: salario=${numSal} (diario=${sueldoDiario.toFixed(2)}), dias=${numDias}, pct=${(pct*100).toFixed(2)}% => monto=${calc}`);
        return calc;
    }

    log(`Subsidio: no se pudo calcular (faltan datos) para codigo ${getField(subRow,['CODIGO','COD'])||getField(subRow,['Codigo Empleado','Codigo'])} — monto 0`);
    return 0;
}

// Calcular
btnCalc.addEventListener('click', () => {
    status.textContent = ''; 
    preview.innerHTML = ''; 
    results = [];
    
    if(!mesInput.value || !mesInput.value.match(/^\d{6}$/)) { 
        showAlert('Ingresa un mes valido (mmyyyy) con 6 digitos', 'error'); 
        return; 
    }
    
    updateProgress(0); 
    log('Inicio de calculo');

    const mapISSS = {}, mapDUI = {}, mapAFP = {};
    const baseRows = sheetsData['BASE'] || [];
    const isssRows = sheetsData['ISSS'] || [];
    const duiRows = sheetsData['DUI'] || [];
    const ingRows = sheetsData['INGRESOS'] || sheetsData['ING'] || [];
    const subRows = sheetsData['SUBSIDIO'] || [];

    function proc(arr, fn) { 
        arr.forEach((r, i) => {
            if (i % 50 === 0) updateProgress((i / arr.length) * 33);
            fn(r, i);
        }); 
    }

    // BASE -> AFP
    proc(baseRows, row => {
        const code = String(getField(row, ['CODIGO','COD','Codigo Empleado','Codigo']) || '').trim();
        if(!code) return;
        const afpVal = String(getField(row, ['AFP','AFP_NAME','AFPS']) || '').toUpperCase();
        mapAFP[code] = (afpVal === 'CRECER') ? 'MAX' : (afpVal === 'CONFIA' ? 'COF' : (afpVal === 'IPSFA' ? 'IPS' : 'COF'));
    });

    // ISSS
    proc(isssRows, row => {
        const code = String(getField(row, ['CODIGO','COD','Codigo Empleado','Codigo']) || '').trim();
        if(!code) return;
        mapISSS[code] = padLeft(digitsOnly(getField(row, ['ISSS','ISSS_ID','NUM_ISSS']) || ''), 9);
    });

    // DUI
    proc(duiRows, row => {
        const code = String(getField(row, ['CODIGO','COD','Codigo Empleado','Codigo']) || '').trim();
        if(!code) return;
        mapDUI[code] = padLeft(digitsOnly(getField(row, ['DUI','DUI_NUM','NUM_DUI']) || ''), 9);
    });

    updateProgress(33);
    log('Mapas de referencia creados');

    // INGRESOS - Procesar filas de ingresos
    proc(ingRows, (row, i) => {
        updateProgress(33 + (i / ingRows.length) * 33);
        
        const code = String(getField(row, ['CODIGO','COD','Codigo Empleado','Codigo']) || '').trim();
        if(!code) return;
        
        const sal = toNumber(getField(row, ['SALARIOS','SALARIO','SUELDO']) || 0);
        const otros = toNumber(getField(row, ['PAGOS ADICIONALES','PAGOS','OTROS']) || 0);
        const vac = toNumber(getField(row, ['VACACION','VACACION','VACACIONES']) || 0);
        const diasTrabajados = Math.round(toNumber(getField(row, ['DIAS_TRABAJADOS','DIAS TRABAJADOS','DIAS']) || 0));
        const obsText = String(getField(row, ['observacion','OBSERVACION','OBS']) || '').toLowerCase();
        const [pn, sn, pa, sa] = splitName(getField(row, ['Nombre Empleado','NOMBRE','NOMBRE EMPLEADO']) || '');
        let obsCodes = [];

        // PRIORIDAD: Si es jubilado/pensionado (codigo 04)
        if(obsText.includes('pensionado') || obsText.includes('jubilado')) {
            obsCodes.push('04');
        } else {
            // Otras observaciones
            if(vac > 0 && otros > 0) obsCodes.push('10');
            else if(vac > 0) obsCodes.push('09');
            else if(otros > 0) obsCodes.push('02');
            
            if(obsText.includes('licencia') || obsText.includes('permiso')) obsCodes.push('05');
            if(obsText.includes('incapacidad') || getField(row, ['DIAS_INCAPACIDAD','DIAS_INCAP'])) obsCodes.push('06');
            if(obsText.includes('retiro')) obsCodes.push('07');
            if(obsText.includes('subsidio')) obsCodes.push('12');
            
            if(obsCodes.length === 0) obsCodes.push('01');
        }

        // Vacaciones: asignar dias_vacacion = 15 si hay monto vacacion, sino 0
        let diasVac = (vac > 0) ? 15 : 0;

        // construir objeto resultante
        results.push({
            CODIGO: code,
            PN: pn, SN: sn, PA: pa, SA: sa,
            SALARIOS: sal.toFixed(2),
            PAGOS_ADICIONALES: otros.toFixed(2),
            VACACION: vac.toFixed(2),
            DIAS: diasTrabajados,
            DIAS_VACACION: diasVac,
            AFP: mapAFP[code] || 'COF',
            ISSS: mapISSS[code] || '',
            DUI: mapDUI[code] || '',
            OBS_ARRAY: obsCodes, // temporal, para procesar obs1/obs2 luego
            _ROW_SOURCE: 'ING'
        });
    });

    updateProgress(66);
    log('Procesados ' + results.length + ' registros de ingresos');

    // SUBSIDIO -> agregar al final, calcular monto_subsidio y dias_incapacidad si aplica
    proc(subRows, (row, i) => {
        updateProgress(66 + (i / subRows.length) * 34);
        
        const code = String(getField(row, ['CODIGO','COD','Codigo Empleado','Codigo']) || '').trim();
        if(!code) return;
        const [pn, sn, pa, sa] = splitName(getField(row, ['Nombre Empleado','NOMBRE','NOMBRE EMPLEADO']) || '');
        const diasIncap = Math.round(toNumber(getField(row, ['DIAS_INCAPACIDAD','DIAS','DIAS_SUBSIDIO','DIAS_INCAP']) || 0));
        const montoSub = computeSubsidioAmount(row, null); // intentamos solo con la fila de subsidio
        const obsArr = ['12']; // por defecto subsidio como observacion principal

        if(diasIncap > 0) {
            // si hay dias de incapacidad ademas del subsidio, agregamos codigo 06 (incapacidad) como secundario
            obsArr.push('06');
        }

        // si usuario incluyo texto que indique incapacidad en la fila de subsidio
        const obsText = String(getField(row, ['observacion','OBSERVACION','OBS']) || '').toLowerCase();
        if(obsText.includes('incapacidad') && !obsArr.includes('06')) obsArr.push('06');

        // construir resultado: SALARIOS en 0, dias_trabajados = diasIncap (cuando aplique)
        results.push({
            CODIGO: code,
            PN: pn, SN: sn, PA: pa, SA: sa,
            SALARIOS: '0.00',
            PAGOS_ADICIONALES: '0.00',
            VACACION: '0.00',
            DIAS: diasIncap,
            DIAS_VACACION: 0,
            AFP: mapAFP[code] || 'COF',
            ISSS: mapISSS[code] || '',
            DUI: mapDUI[code] || '',
            OBS_ARRAY: obsArr,
            MONTO_SUBSIDIO: montoSub.toFixed(2),
            _ROW_SOURCE: 'SUBSIDIO'
        });
    });

    // Normalizar observaciones (obs1, obs2): si falta alguno, poner '00'
    results = results.map(r => {
        const arr = Array.isArray(r.OBS_ARRAY) ? r.OBS_ARRAY.slice() : [];
        const obs1 = arr[0] || '00';
        const obs2 = arr[1] || '00';
        return Object.assign({}, r, {OBS1: obs1, OBS2: obs2});
    });

    updateProgress(100); 
    log('Calculo finalizado. Total de registros: ' + results.length);

    // Vista previa
    if (results.length > 0) {
        let html = '<table><thead><tr>' + 
            ['CODIGO','PN','SN','PA','SA','SALARIOS','PAGOS_ADICIONALES','VACACION','DIAS','DIAS_VACACION','AFP','ISSS','DUI','OBS1','OBS2','MONTO_SUBSIDIO']
            .map(h => `<th>${h}</th>`).join('') + 
            '</tr></thead><tbody>';
        
        results.slice(0, 40).forEach(r => {
            let trClass = '';
            if (String(r.OBS1 || '').includes('04')) trClass = 'jubilado';
            else if (String(r.OBS1 || '').includes('06') || (r.OBS2 && String(r.OBS2).includes('06'))) trClass = 'incapacidad';
            else if (String(r.OBS1 || '').includes('09') || (r.OBS2 && String(r.OBS2).includes('09'))) trClass = 'vacacion';
            else if (String(r.OBS1 || '').includes('12') || (r.OBS2 && String(r.OBS2).includes('12'))) trClass = 'subsidio';
            
            html += `<tr class="${trClass}">` + 
                ['CODIGO','PN','SN','PA','SA','SALARIOS','PAGOS_ADICIONALES','VACACION','DIAS','DIAS_VACACION','AFP','ISSS','DUI','OBS1','OBS2','MONTO_SUBSIDIO']
                .map(k => `<td>${r[k] !== undefined ? r[k] : ''}</td>`).join('') + 
                '</tr>';
        });
        
        html += '</tbody></table>'; 
        preview.innerHTML = html;
        
        if (results.length > 40) {
            preview.innerHTML += `<p>Mostrando 40 de ${results.length} registros. Todos se incluiran en la exportacion.</p>`;
        }
        
        btnExport.disabled = false;
        showAlert('Calculo completado. Revise la vista previa antes de exportar.', 'success');
    } else {
        showAlert('No se encontraron datos para procesar.', 'error');
    }
});

// Exportar CSV/XLSX
btnExport.addEventListener('click', () => {
    if(results.length === 0) { 
        showAlert('No hay datos para exportar', 'error'); 
        return; 
    }
    
    const numeroPatronal = "501760146";
    const mesanyo = mesInput.value;
    const codigo001 = "001";
    const codigo01 = "01";
    const campoVacio = "";
    const horasLaboradas = 8;
    const tipoFila = "Normal";

    log('Iniciando exportacion de ' + results.length + ' registros');
    updateProgress(0);

    const ssfRows = results.map((r, i) => {
        if (i % 50 === 0) updateProgress((i / results.length) * 50);
        
        const obs1 = String(r.OBS1 || '00');
        const obs2 = String(r.OBS2 || '00');
        const obs1txt = obsCatalog[obs1] || '';
        const obs2txt = obsCatalog[obs2] || '';
        
        return {
            numero_patronal: String(numeroPatronal),
            mesanyo: String(mesanyo),
            codigo_001: String(codigo001),
            dui: String(r.DUI || ''),
            codigo_01: String(codigo01),
            isss: String(r.ISSS || ''),
            afp: String(r.AFP || ''),
            primer_nombre: String(r.PN || ''),
            segundo_nombre: String(r.SN || ''),
            primer_apellido: String(r.PA || ''),
            segundo_apellido: String(r.SA || ''),
            campo_vacio: String(campoVacio),
            monto_salario: String(r.SALARIOS || '0.00'),
            monto_pagos_adicionales: String(r.PAGOS_ADICIONALES || '0.00'),
            monto_vacacion: String(r.VACACION || '0.00'),
            dias_trabajados: String(r.DIAS || 0),
            horas_laboradas: String(horasLaboradas),
            dias_vacacion: String(r.DIAS_VACACION || 0),
            codigo_observacion: obs1,
            comentario_observacion: obs1txt,
            codigo_observacion2: obs2,
            comentario_observacion2: obs2txt,
            tipo_fila: String(tipoFila),
            monto_subsidio: String(r.MONTO_SUBSIDIO || '0.00'),
            _code: String(r.CODIGO || '')
        };
    });

    updateProgress(50);
    log('Generando archivo CSV');

    // --- CSV ---
    const csvHeaders = Object.keys(ssfRows[0]).filter(k => !k.startsWith('_'));
    let csvContent = csvHeaders.join(',') + '\n';
    
    ssfRows.forEach((r, i) => {
        if (i % 50 === 0) updateProgress(50 + (i / ssfRows.length) * 25);
        csvContent += csvHeaders.map(h => `"${String(r[h] || '')}"`).join(',') + '\n';
    });
    
    const csvBlob = new Blob([csvContent], {type: 'text/csv;charset=utf-8;'});
    const csvUrl = URL.createObjectURL(csvBlob);
    const csvLink = document.createElement('a');
    csvLink.href = csvUrl;
    csvLink.download = `SSF_INGRESOS_${mesanyo}.csv`;
    csvLink.click();
    URL.revokeObjectURL(csvUrl);

    updateProgress(75);
    log('Generando archivo XLSX');

    // --- XLSX: forzar todas las celdas a texto ---
    // Construir un arreglo AOA (array of arrays) con encabezado + filas (todo como strings)
    const aoa = [];
    aoa.push(csvHeaders); // encabezado
    
    ssfRows.forEach(r => {
        const rowData = csvHeaders.map(h => {
            // Forzar todos los valores a string para preservar ceros a la izquierda
            const value = String(r[h] || '');
            return value;
        });
        aoa.push(rowData);
    });

    const ws = XLSX.utils.aoa_to_sheet(aoa);

    // Forzar tipo 's' (string) en todas las celdas para preservar ceros a la izquierda
    const range = XLSX.utils.decode_range(ws['!ref']);
    for(let R = range.s.r; R <= range.e.r; ++R) {
        for(let C = range.s.c; C <= range.e.c; ++C) {
            const cell_address = {c:C, r:R};
            const cell_ref = XLSX.utils.encode_cell(cell_address);
            if(!ws[cell_ref]) continue;
            
            // Forzar todas las celdas a tipo texto
            ws[cell_ref].t = 's';
            
            // Asegurar que los valores numéricos con ceros a la izquierda se mantengan como texto
            if(ws[cell_ref].v !== undefined && typeof ws[cell_ref].v === 'number') {
                ws[cell_ref].v = String(ws[cell_ref].v);
            }
        }
    }

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'SSF_INGRESOS');
    
    updateProgress(95);
    XLSX.writeFile(wb, `SSF_INGRESOS_${mesanyo}.xlsx`);
    
    updateProgress(100);
    log('Exportacion completada: SSF_INGRESOS_' + mesanyo + '.csv y SSF_INGRESOS_' + mesanyo + '.xlsx');
    
    showAlert('Archivos exportados correctamente', 'success');
    
    // Ocultar barra de progreso después de un momento
    setTimeout(() => {
        progressBar.style.display = 'none';
    }, 2000);
});