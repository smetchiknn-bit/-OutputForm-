import * as XLSX from 'xlsx-js-style';

export interface SheetData {
  name: string;
  data: any[][];
  formulas?: any[][];
}

export interface ProcessingResult {
  success: boolean;
  message: string;
  workbook?: XLSX.WorkBook;
  sheets?: SheetData[];
}

export interface ProcessingLog {
  step: string;
  status: 'success' | 'error' | 'warning';
  message: string;
}

export function validateWorkbook(wb: XLSX.WorkBook): { valid: boolean; message: string } {
  const sheetNames = wb.SheetNames;
  
  if (sheetNames.length < 3) {
    return { valid: false, message: 'Файл должен содержать минимум 3 листа' };
  }

  if (sheetNames[0] !== 'Свод') {
    return { valid: false, message: `Имя листа 1 ("${sheetNames[0]}") не соответствует инструкции. Ожидается "Свод"` };
  }
  if (sheetNames[1] !== 'СМР уник') {
    return { valid: false, message: `Имя листа 2 ("${sheetNames[1]}") не соответствует инструкции. Ожидается "СМР уник"` };
  }
  if (sheetNames[2] !== 'ТМЦ уник') {
    return { valid: false, message: `Имя листа 3 ("${sheetNames[2]}") не соответствует инструкции. Ожидается "ТМЦ уник"` };
  }

  return { valid: true, message: 'Структура файла проверена успешно' };
}

export function performNumbering(inputData: any[][]): { data: any[][]; nElements: number } {
  const data = inputData;
  const nElements = data.length - 1;
  let nO = 0, nK = 0, nC = 0, nU = 0, nE = 0;
  let nL1 = 0, nL2 = 0, nL3 = 0, nL4 = 0;
  let nGr = 0, nKER = 0, nTMZ = 0;
  let r1 = '', r2 = '', r3 = '', r4 = '', r5 = '';
  let r6 = '', r7 = '', r8 = '', r9 = '';
  let r10 = '', r11 = '', r12 = '';

  for (let i = 0; i < data.length; i++) {
    while (data[i].length < 19) {
      data[i].push(null);
    }
  }

  data[0][0] = 'Код ССР\nИД КЕР\nИД ТМЦ';
  data[0][1] = 'Статья ССР\nНаименование КЕР\nНаименование ТМЦ';
  data[0][2] = 'ЕдИзм КЕР\nЕдИзм ТМЦ';
  data[0][3] = 'Объем СМР\nРасход ТМЦ';
  data[0][4] = 'Цена СМР\nза ЕдИзм КЕР\nОбъем ТМЦ';
  data[0][5] = 'Цена ТМЦ\nза ЕдИзм КЕР\nза ЕдИзм ТМЦ';
  data[0][6] = 'Цена\nСМР+ТМЦ\nза ЕдИзм КЕР';
  data[0][7] = 'СМР всего\nруб с НДС';
  data[0][8] = 'ТМЦ всего\nруб с НДС';
  data[0][9] = 'Стоимость\nвсего\nруб с НДС';
  
  data[0][10] = 'ТА';
  data[0][11] = 'КЕР.в.СР';
  data[0][12] = 'Было.Объем.КЕР_ТМЦ';
  data[0][13] = 'Количество точек';
  data[0][14] = 'Иерархия';
  data[0][15] = 'Признак КЕР';
  data[0][16] = 'Признак ТМЦ';
  data[0][17] = 'ИД.ДС';
  data[0][18] = 'ИД.акта.ДС.ФСК';

  for (let i = 1; i <= nElements; i++) {
    const row = data[i];
    const type = row[10];

    if (type === 'О') {
      nO++; r1 = nO + '.';
      row[12] = nO;
      nK = 0; nC = 0; nU = 0; nE = 0; nL1 = 0; nL2 = 0; nL3 = 0; nL4 = 0; nGr = 0; nKER = 0; nTMZ = 0;
    } else if (type === 'К') {
      nK++; r2 = nK + '.';
      row[12] = r1 + nK;
      nC = 0; nU = 0; nE = 0; nL1 = 0; nL2 = 0; nL3 = 0; nL4 = 0; nGr = 0; nKER = 0; nTMZ = 0;
    } else if (type === 'С') {
      nC++; r3 = nC + '.';
      row[12] = r1 + r2 + nC;
      nU = 0; nE = 0; nL1 = 0; nL2 = 0; nL3 = 0; nL4 = 0; nGr = 0; nKER = 0; nTMZ = 0;
    } else if (type === 'У') {
      nU++; r4 = nU + '.';
      row[12] = r1 + r2 + r3 + nU;
      nE = 0; nL1 = 0; nL2 = 0; nL3 = 0; nL4 = 0; nGr = 0; nKER = 0; nTMZ = 0;
    } else if (type === 'Э') {
      nE++; r5 = nE + '.';
      row[12] = r1 + r2 + r3 + r4 + nE;
      nL1 = 0; nL2 = 0; nL3 = 0; nL4 = 0; nGr = 0; nKER = 0; nTMZ = 0;
    } else if (type === 'Л1') {
      nL1++; r6 = nL1 + '.';
      row[12] = r1 + r2 + r3 + r4 + r5 + nL1;
      nL2 = 0; nL3 = 0; nL4 = 0; nGr = 0; nKER = 0; nTMZ = 0;
    } else if (type === 'Л2') {
      nL2++; r7 = nL2 + '.';
      row[12] = r1 + r2 + r3 + r4 + r5 + r6 + nL2;
      nL3 = 0; nL4 = 0; nGr = 0; nKER = 0; nTMZ = 0;
    } else if (type === 'Л3') {
      nL3++; r8 = nL3 + '.';
      row[12] = r1 + r2 + r3 + r4 + r5 + r6 + r7 + nL3;
      nL4 = 0; nGr = 0; nKER = 0; nTMZ = 0;
    } else if (type === 'Л4') {
      nL4++; r9 = nL4 + '.';
      row[12] = r1 + r2 + r3 + r4 + r5 + r6 + r7 + r8 + nL4;
      nGr = 0; nKER = 0; nTMZ = 0;
    } else if (type === 'ГР') {
      nGr++; r10 = nGr + '.';
      row[12] = r1 + r2 + r3 + r4 + r5 + r6 + r7 + r8 + r9 + nGr;
      nKER = 0; nTMZ = 0;
    } else if (type === 'КЕР') {
      nKER++; r11 = nKER + '.';
      row[12] = r1 + r2 + r3 + r4 + r5 + r6 + r7 + r8 + r9 + r10 + nKER;
      row[15] = 1;
      if (row[0] !== null && row[0] !== undefined) {
        const numVal = Number(row[0]);
        if (!isNaN(numVal)) row[0] = numVal;
      }
    } else if (type === 'ТМЦ') {
      nTMZ++; r12 = nTMZ + '.';
      row[12] = r1 + r2 + r3 + r4 + r5 + r6 + r7 + r8 + r9 + r10 + r11 + nTMZ;
      row[16] = 1;
      if (row[0] !== null && row[0] !== undefined) {
        const numVal = Number(row[0]);
        if (!isNaN(numVal)) row[0] = numVal;
      }
    }

    if (row[12] !== null && row[12] !== undefined) {
      const str = String(row[12]);
      row[13] = (str.match(/\./g) || []).length;
      const dots = row[13];
      if (dots > 0) {
        const parts = str.split('.');
        let hier = '';
        for (let d = 0; d < dots; d++) {
          hier += parts[d] + '.';
        }
        row[14] = hier;
      } else {
        row[14] = str;
      }
    }
  }

  return { data, nElements };
}

export function computeFormulas(inputData: any[][], nElements: number, smrUnikData: any[][], tmzUnikData: any[][]): { data: any[][]; formulas: any[][] } {
  const data = inputData;
  const smrDict: Map<number, { price: number; volume: number }> = new Map();
  const tmzDict: Map<number, { price: number; volume: number }> = new Map();

  for (let i = 1; i < smrUnikData.length; i++) {
    const id = Number(smrUnikData[i][0]);
    if (!isNaN(id)) {
      smrDict.set(id, { price: Number(smrUnikData[i][4]) || 0, volume: Number(smrUnikData[i][3]) || 0 });
    }
  }

  for (let i = 1; i < tmzUnikData.length; i++) {
    const id = Number(tmzUnikData[i][0]);
    if (!isNaN(id)) {
      tmzDict.set(id, { price: Number(tmzUnikData[i][4]) || 0, volume: Number(tmzUnikData[i][3]) || 0 });
    }
  }

  const formulas: any[][] = data.map(row => new Array(row.length).fill(null));

  for (let i = 1; i <= nElements; i++) {
    const row = data[i];
    const formulaRow = formulas[i];
    const isKER = row[15] === 1;
    const isTMZ = row[16] === 1;
    const rowNum = i + 1;

    if (isKER) {
      const kodSSR = Number(row[0]);
      const volume = Number(row[3]) || 0;
      
      formulaRow[4] = 'IFERROR(ROUND(VLOOKUP(A' + rowNum + ",'СМР уник'!A:F,5,FALSE),2)*P" + rowNum + ',0)';
      if (!isNaN(kodSSR) && smrDict.has(kodSSR)) {
        const smrInfo = smrDict.get(kodSSR)!;
        row[4] = Math.round(smrInfo.price * volume * 100) / 100;
      } else {
        row[4] = 0;
      }

      const hier = row[12];
      let sumG = 0;
      let countChildren = 0;
      for (let j = 1; j <= nElements; j++) {
        if (j === i) continue;
        const childHier = String(data[j][12] || '');
        if (childHier.startsWith(String(hier) + '.') && data[j][16] === 1) {
          sumG += Number(data[j][6]) || 0;
          countChildren++;
        }
      }
      formulaRow[5] = 'ROUND(SUMIF(O:O,M' + rowNum + '&".",G:G)/D' + rowNum + ',2)';
      if (countChildren > 0 && Number(row[3]) !== 0) {
        row[5] = Math.round(sumG / Number(row[3]) * 100) / 100;
      } else {
        row[5] = 0;
      }

      formulaRow[6] = 'ROUND(E' + rowNum + '+F' + rowNum + ',2)';
      row[6] = Math.round((Number(row[4]) + Number(row[5])) * 100) / 100;

      formulaRow[7] = 'ROUND(D' + rowNum + '*E' + rowNum + ',2)';
      row[7] = Math.round(Number(row[3]) * Number(row[4]) * 100) / 100;

      let tmzTotal = 0;
      for (let j = 1; j <= nElements; j++) {
        if (j === i) continue;
        const childHier = String(data[j][12] || '');
        if (childHier.startsWith(String(hier) + '.') && data[j][16] === 1) {
          tmzTotal += Number(data[j][6]) || 0;
        }
      }
      formulaRow[8] = 'IFERROR(ROUND(IF(P' + rowNum + '=1,SUM(OFFSET($G' + rowNum + ',1,0,COUNTIF($O:$O,$M' + rowNum + '&"."))),0),2),0)';
      row[8] = Math.round(tmzTotal * 100) / 100;

      formulaRow[9] = 'ROUND(H' + rowNum + '+I' + rowNum + ',2)';
      row[9] = Math.round((Number(row[7]) + Number(row[8])) * 100) / 100;

    } else if (isTMZ) {
      const kodSSR = Number(row[0]);
      
      formulaRow[5] = 'IFERROR(ROUND(VLOOKUP(A' + rowNum + ",'ТМЦ уник'!A:E,5,FALSE),2)*Q" + rowNum + ',0)';
      if (!isNaN(kodSSR) && tmzDict.has(kodSSR)) {
        const tmzInfo = tmzDict.get(kodSSR)!;
        row[5] = Math.round(tmzInfo.price * Number(row[3]) * 100) / 100;
      } else {
        row[5] = 0;
      }

      formulaRow[6] = 'ROUND(E' + rowNum + '*F' + rowNum + ',2)';
      row[6] = Math.round(Number(row[4]) * Number(row[5]) * 100) / 100;

    } else {
      const hier = row[12];
      if (hier !== null && hier !== undefined) {
        formulaRow[7] = 'ROUND(SUMIF(O:O,M' + rowNum + '&".",H:H),2)';
        let smrTotal = 0;
        let tmzTotal = 0;
        for (let j = 1; j <= nElements; j++) {
          const childHier = String(data[j][12] || '');
          if (childHier.startsWith(String(hier) + '.')) {
            if (data[j][15] === 1) {
              smrTotal += Number(data[j][7]) || 0;
              tmzTotal += Number(data[j][8]) || 0;
            }
          }
        }
        row[7] = Math.round(smrTotal * 100) / 100;

        formulaRow[8] = 'ROUND(SUMIF(O:O,M' + rowNum + '&".",I:I),2)';
        row[8] = Math.round(tmzTotal * 100) / 100;

        formulaRow[9] = 'ROUND(H' + rowNum + '+I' + rowNum + ',2)';
        row[9] = Math.round((smrTotal + tmzTotal) * 100) / 100;
      }
    }
  }

  return { data, formulas };
}

export function processSMRUnik(smrData: any[][], svodData: any[][]): { data: any[][]; formulas: any[][] } {
  const nElements = smrData.length - 1;

  for (let i = 0; i < smrData.length; i++) {
    while (smrData[i].length < 15) smrData[i].push(null);
  }

  // Устанавливаем заголовки для всех колонок
  smrData[0][0] = 'ИД.КЕР';
  smrData[0][1] = 'Наименование\nКЕР';
  smrData[0][2] = 'ЕдИзм КЕР';
  smrData[0][3] = 'Объем СМР\nВсего в СР';
  smrData[0][4] = 'СМР за ЕдИзм';
  smrData[0][5] = 'СМР Всего';
  smrData[0][6] = 'Состав Работ';
  smrData[0][7] = 'ФЕР для КЕР';
  smrData[0][8] = 'Цена СМР\nмедиана';
  smrData[0][9] = 'Цена СМР\nсредняя';
  smrData[0][10] = 'Резервный\nстолбец';
  smrData[0][11] = '';
  smrData[0][12] = 'Объем СМР\nиз Свод';
  smrData[0][13] = 'Проверка';

  const formulas: any[][] = smrData.map(row => new Array(row.length).fill(null));

  const svodSums: Map<number, number> = new Map();
  for (let i = 1; i < svodData.length; i++) {
    if (svodData[i][15] === 1) {
      const id = Number(svodData[i][0]);
      if (!isNaN(id)) {
        const vol = Number(svodData[i][3]) || 0;
        svodSums.set(id, (svodSums.get(id) || 0) + vol);
      }
    }
  }

  let totalSMR = 0;
  for (let i = 1; i <= nElements; i++) {
    const row = smrData[i];
    const formulaRow = formulas[i];
    const id = Number(row[0]);
    const rowNum = i + 1;
    
    if (!isNaN(id) && svodSums.has(id)) {
      row[12] = svodSums.get(id);
    } else {
      row[12] = 0;
    }

    const localVol = Number(row[3]) || 0;
    row[13] = Math.round((Number(row[12]) - localVol) * 100) / 100;
    
    formulaRow[5] = 'ROUND(D' + rowNum + '*E' + rowNum + ',2)';
    
    row[5] = Math.round(Number(row[3]) * Number(row[4]) * 100) / 100;
    totalSMR += Number(row[5]) || 0;
  }

  smrData[nElements + 1] = smrData[nElements + 1] || [];
  while (smrData[nElements + 1].length < 15) smrData[nElements + 1].push(null);
  
  const totalRow = nElements + 1;
  formulas[totalRow] = formulas[totalRow] || new Array(15).fill(null);
  formulas[totalRow][5] = 'SUM(F2:F' + (nElements + 1) + ')';
  smrData[totalRow][5] = Math.round(totalSMR * 100) / 100;
  smrData[totalRow][1] = 'ИТОГО';

  return { data: smrData, formulas };
}

export function processTMZUnik(tmzData: any[][], svodData: any[][]): { data: any[][]; formulas: any[][] } {
  const nElements = tmzData.length - 1;

  for (let i = 0; i < tmzData.length; i++) {
    while (tmzData[i].length < 14) tmzData[i].push(null);
  }

  // Устанавливаем заголовки для всех колонок
  tmzData[0][0] = 'ИД.ТМЦ';
  tmzData[0][1] = 'Наименование\nТМЦ';
  tmzData[0][2] = 'ЕдИзм ТМЦ';
  tmzData[0][3] = 'Объем ТМЦ\nВсего в СР';
  tmzData[0][4] = 'Цена ТМЦ\nза ЕдИзм';
  tmzData[0][5] = 'ТМЦ Всего';
  tmzData[0][6] = 'ИД.Поставщика';
  tmzData[0][7] = 'Имя.Поставщика';
  tmzData[0][8] = 'Номинация';
  tmzData[0][9] = 'Резервный\nстолбец';
  tmzData[0][10] = '';
  tmzData[0][11] = 'Объем ТМЦ\nиз Свод';
  tmzData[0][12] = 'Проверка';

  const formulas: any[][] = tmzData.map(row => new Array(row.length).fill(null));

  const svodSums: Map<number, number> = new Map();
  for (let i = 1; i < svodData.length; i++) {
    if (svodData[i][16] === 1) {
      const id = Number(svodData[i][0]);
      if (!isNaN(id)) {
        const vol = Number(svodData[i][3]) || 0;
        svodSums.set(id, (svodSums.get(id) || 0) + vol);
      }
    }
  }

  let totalTMZ = 0;
  for (let i = 1; i <= nElements; i++) {
    const row = tmzData[i];
    const formulaRow = formulas[i];
    const id = Number(row[0]);
    const rowNum = i + 1;
    
    if (!isNaN(id) && svodSums.has(id)) {
      row[11] = svodSums.get(id);
    } else {
      row[11] = 0;
    }

    const localVol = Number(row[3]) || 0;
    row[12] = Math.round((Number(row[11]) - localVol) * 100) / 100;
    
    formulaRow[5] = 'ROUND(D' + rowNum + '*E' + rowNum + ',2)';
    
    row[5] = Math.round(Number(row[3]) * Number(row[4]) * 100) / 100;
    totalTMZ += Number(row[5]) || 0;
  }

  tmzData[nElements + 1] = tmzData[nElements + 1] || [];
  while (tmzData[nElements + 1].length < 14) tmzData[nElements + 1].push(null);
  
  const totalRow = nElements + 1;
  formulas[totalRow] = formulas[totalRow] || new Array(14).fill(null);
  formulas[totalRow][5] = 'SUM(F2:F' + (nElements + 1) + ')';
  tmzData[totalRow][5] = Math.round(totalTMZ * 100) / 100;
  tmzData[totalRow][1] = 'ИТОГО';

  return { data: tmzData, formulas };
}

export function getRowType(data: any[][], rowIndex: number): string {
  if (rowIndex === 0) return 'header';
  return data[rowIndex]?.[10] || '';
}

export function getRowStyle(type: string, colIndex?: number): string {
  const baseStyles: { [key: string]: string } = {
    'О': 'bg-[#FFF2CB]', // Кремовый
    'К': 'bg-[#D9E2F3]', // Светло-голубой
    'С': 'bg-[#D9E2F3]',
    'У': 'bg-[#D9E2F3]',
    'Э': 'bg-[#D9E2F3]',
    'Л1': 'bg-[#F2F2F2]', // Очень светло-серый
    'Л2': 'bg-[#F2F2F2]',
    'Л3': 'bg-[#F2F2F2]',
    'ГР': 'bg-[#FFD965] font-bold', // Золотистый
    'КЕР': 'bg-white', // Белый
    'ТМЦ': 'bg-[#E2EFD9]', // Светло-зеленый
    'header': 'bg-[#dcfce7]', // Светло-зеленый (новая палитра)
  };

  if (colIndex !== undefined && colIndex > 9) {
    return '';
  }

  // Курсив для колонок B-J для типов О, К, С, У, Э, Л1, Л2, Л3
  const italicTypes = ['О', 'К', 'С', 'У', 'Э', 'Л1', 'Л2', 'Л3'];
  if (colIndex !== undefined && colIndex > 0 && italicTypes.includes(type)) {
    return (baseStyles[type] || 'bg-white') + ' italic';
  }

  return baseStyles[type] || 'bg-white';
}

export function getSheet1CRowStyle(data: any[][], rowIndex: number): string {
  if (rowIndex === 0) return 'header';
  const rowTypes = (data as any)._rowTypes;
  if (!rowTypes) return '';
  const type = rowTypes[rowIndex - 1];
  
  switch (type) {
    case 'О': return 'bg-[#D9D9D9] font-bold text-sm';
    case 'К': return 'font-bold text-base pl-4';
    case 'С': return 'pl-8';
    case 'У': return 'pl-12';
    case 'Э': return 'pl-16';
    case 'Л1': return 'text-xs pl-20';
    case 'Л2': return 'text-xs pl-24';
    case 'Л3': return 'text-xs pl-28';
    case 'КЕР': return 'bg-[#FFFFCC] font-bold';
    case 'КЕР_ТМЦ': 
    case 'ТМЦ': return 'bg-[#C6EFCE]';
    default: return '';
  }
}

// Применение стилей к листу Свод
function applySvodStyles(ws: XLSX.WorkSheet, data: any[][]) {
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
  
  // Обрабатываем все 19 колонок (A-S)
  for (let R = range.s.r; R <= range.e.r; R++) {
    for (let C = range.s.c; C <= Math.min(range.e.c, 18); C++) {
      const cellRef = XLSX.utils.encode_cell({ r: R, c: C });
      
      // Создаём ячейку если её нет (для применения заливки к пустым ячейкам)
      if (!ws[cellRef]) {
        ws[cellRef] = { t: 's', v: '', w: '' };
      }
      const cell = ws[cellRef];
      
      // Создаем стиль для ячейки
      const style: any = {
        font: { name: 'Calibri', sz: 8, color: { rgb: '000000' } },
        alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
        border: {
          top: { style: 'thin', color: { rgb: '000000' } },
          bottom: { style: 'thin', color: { rgb: '000000' } },
          left: { style: 'thin', color: { rgb: '000000' } },
          right: { style: 'thin', color: { rgb: '000000' } }
        }
      };
      
      // Заголовок (ТА)
      if (R === 0) {
        style.fill = { fgColor: { rgb: 'D8D8D8' } }; // Серый
        style.font = { name: 'Calibri', sz: 8, color: { rgb: '000000' } };
      } else {
        const type = data[R]?.[10];
        
        // Цвета по типу строки (только для колонок A-J)
        if (C <= 9) {
          switch (type) {
            case 'О':
              style.fill = { fgColor: { rgb: 'FFF2CB' } }; // Кремовый
              // Курсив для B-J
              if (C > 0) {
                style.font.italic = true;
              }
              break;
            case 'К':
            case 'С':
            case 'У':
            case 'Э':
              style.fill = { fgColor: { rgb: 'D9E2F3' } }; // Светло-голубой
              // Курсив для B-J
              if (C > 0) {
                style.font.italic = true;
              }
              break;
            case 'Л1':
            case 'Л2':
            case 'Л3':
              style.fill = { fgColor: { rgb: 'F2F2F2' } }; // Очень светло-серый
              // Курсив для B-J
              if (C > 0) {
                style.font.italic = true;
              }
              break;
            case 'ГР':
              style.fill = { fgColor: { rgb: 'FFD965' } }; // Светло-золотистый
              style.font.bold = true; // Полужирный для всех колонок
              break;
            case 'КЕР':
              style.fill = { fgColor: { rgb: 'FFFFFF' } }; // Белый
              // Обычный шрифт для всех колонок
              break;
            case 'ТМЦ':
              style.fill = { fgColor: { rgb: 'E2EFD9' } }; // Светло-зеленый
              // Обычный шрифт для всех колонок
              break;
          }
        }
        
        // Выравнивание по колонкам
        if (C === 0 || C === 1) {
          style.alignment.horizontal = 'left';
        } else if (C >= 3 && C <= 9) {
          style.alignment.horizontal = 'right';
        }
        
        // Формат чисел
        if (C >= 3 && C <= 9) {
          cell.z = '#,##0.00';
        }
      }
      
      cell.s = style;
    }
  }
  
  // Обновляем диапазон листа после добавления ячеек (все 19 колонок A-S)
  ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: data.length - 1, c: 18 } });
}

// Применение стилей к листу СМР уник
function applySMRStyles(ws: XLSX.WorkSheet) {
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
  
  // Обрабатываем все 14 колонок (A-N, 0-13)
  for (let R = range.s.r; R <= range.e.r; R++) {
    for (let C = range.s.c; C <= Math.min(range.e.c, 13); C++) {
      const cellRef = XLSX.utils.encode_cell({ r: R, c: C });
      
      // Создаём ячейку если её нет
      if (!ws[cellRef]) {
        ws[cellRef] = { t: 's', v: '', w: '' };
      }
      const cell = ws[cellRef];
      
      const style: any = {
        font: { name: 'Calibri', sz: 8, color: { rgb: '000000' } },
        alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
        border: {
          top: { style: 'thin', color: { rgb: '000000' } },
          bottom: { style: 'thin', color: { rgb: '000000' } },
          left: { style: 'thin', color: { rgb: '000000' } },
          right: { style: 'thin', color: { rgb: '000000' } }
        }
      };
      
      // Заголовок
      if (R === 0) {
        style.fill = { fgColor: { rgb: 'D8D8D8' } };
      } else {
        // Выравнивание по колонкам
        if (C === 1 || C === 6) {
          style.alignment.horizontal = 'left';
        } else if (C >= 3 && C <= 5) {
          style.alignment.horizontal = 'right';
          cell.z = '#,##0.00';
        }
      }
      
      cell.s = style;
    }
  }
  
  // Обновляем диапазон листа (все 14 колонок A-N)
  const maxRow = Math.max(...Object.keys(ws).filter(k => k.match(/^[A-Z]+\d+$/)).map(k => parseInt(k.match(/\d+/)?.[0] || '0')));
  ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: maxRow, c: 13 } });
}

// Применение стилей к листу ТМЦ уник
function applyTMZStyles(ws: XLSX.WorkSheet) {
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
  
  // Обрабатываем все 13 колонок (A-M, 0-12)
  for (let R = range.s.r; R <= range.e.r; R++) {
    for (let C = range.s.c; C <= Math.min(range.e.c, 12); C++) {
      const cellRef = XLSX.utils.encode_cell({ r: R, c: C });
      
      // Создаём ячейку если её нет
      if (!ws[cellRef]) {
        ws[cellRef] = { t: 's', v: '', w: '' };
      }
      const cell = ws[cellRef];
      
      const style: any = {
        font: { name: 'Calibri', sz: 8, color: { rgb: '000000' } },
        alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
        border: {
          top: { style: 'thin', color: { rgb: '000000' } },
          bottom: { style: 'thin', color: { rgb: '000000' } },
          left: { style: 'thin', color: { rgb: '000000' } },
          right: { style: 'thin', color: { rgb: '000000' } }
        }
      };
      
      // Заголовок
      if (R === 0) {
        style.fill = { fgColor: { rgb: 'D8D8D8' } };
      } else {
        // Выравнивание по колонкам
        if (C === 1) {
          style.alignment.horizontal = 'left';
        } else if (C >= 3 && C <= 5) {
          style.alignment.horizontal = 'right';
          cell.z = '#,##0.00';
        }
      }
      
      cell.s = style;
    }
  }
  
  // Обновляем диапазон листа (все 13 колонок A-M)
  const maxRow = Math.max(...Object.keys(ws).filter(k => k.match(/^[A-Z]+\d+$/)).map(k => parseInt(k.match(/\d+/)?.[0] || '0')));
  ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: maxRow, c: 12 } });
}

// Временно отключено - создание листа 1С
export function processWorkbookWith1C(file: ArrayBuffer, addSheet1C: boolean = false): { result: ProcessingResult; logs: ProcessingLog[] } {
  const logs: ProcessingLog[] = [];
  
  try {
    logs.push({ step: 'Чтение файла', status: 'success', message: 'Файл загружен успешно' });
    const wb = XLSX.read(file, { type: 'array' });

    const validation = validateWorkbook(wb);
    if (!validation.valid) {
      logs.push({ step: 'Проверка структуры', status: 'error', message: validation.message });
      return { result: { success: false, message: validation.message }, logs };
    }
    logs.push({ step: 'Проверка структуры', status: 'success', message: validation.message });

    const svodSheet = wb.Sheets['Свод'];
    const smrSheet = wb.Sheets['СМР уник'];
    const tmzSheet = wb.Sheets['ТМЦ уник'];

    let svodData = XLSX.utils.sheet_to_json<any[]>(svodSheet, { header: 1, defval: null });
    let smrData = XLSX.utils.sheet_to_json<any[]>(smrSheet, { header: 1, defval: null });
    let tmzData = XLSX.utils.sheet_to_json<any[]>(tmzSheet, { header: 1, defval: null });

    const { data: numberedData, nElements } = performNumbering(svodData);
    logs.push({ step: 'Нумерация', status: 'success', message: `Обработано ${nElements} строк` });

    const formulaResult = computeFormulas(numberedData, nElements, smrData, tmzData);
    svodData = formulaResult.data;
    const svodFormulas = formulaResult.formulas;
    logs.push({ step: 'Вычисление формул', status: 'success', message: 'Формулы вычислены' });

    const smrResult = processSMRUnik(smrData, svodData);
    smrData = smrResult.data;
    const smrFormulas = smrResult.formulas;
    logs.push({ step: 'Обработка СМР уник', status: 'success', message: 'Лист СМР уник обработан' });

    const tmzResult = processTMZUnik(tmzData, svodData);
    tmzData = tmzResult.data;
    const tmzFormulas = tmzResult.formulas;
    logs.push({ step: 'Обработка ТМЦ уник', status: 'success', message: 'Лист ТМЦ уник обработан' });

    const newWb = XLSX.utils.book_new();
    
    const svodWs = XLSX.utils.aoa_to_sheet(svodData);
    svodWs['!cols'] = [
      { wch: 12 }, { wch: 50 }, { wch: 10 }, { wch: 12 }, { wch: 12 },
      { wch: 12 }, { wch: 12 }, { wch: 14 }, { wch: 14 }, { wch: 14 },
      { wch: 5, hidden: true }, { wch: 12, hidden: true }, { wch: 18, hidden: true },
      { wch: 12, hidden: true }, { wch: 18, hidden: true }, { wch: 10, hidden: true },
      { wch: 10, hidden: true }, { wch: 10, hidden: true }, { wch: 15, hidden: true },
    ];
    
    // Применяем форматирование
    applySvodStyles(svodWs, svodData);
    
    // Записываем формулы
    if (svodFormulas) {
      for (let r = 0; r < svodFormulas.length; r++) {
        for (let c = 0; c < svodFormulas[r].length; c++) {
          if (svodFormulas[r][c]) {
            const cellRef = XLSX.utils.encode_cell({ r, c });
            if (svodWs[cellRef]) {
              svodWs[cellRef].f = svodFormulas[r][c];
            }
          }
        }
      }
    }
    
    XLSX.utils.book_append_sheet(newWb, svodWs, 'Свод');

    const smrWs = XLSX.utils.aoa_to_sheet(smrData);
    smrWs['!cols'] = [
      { wch: 12 }, { wch: 50 }, { wch: 10 }, { wch: 12 }, { wch: 12 },
      { wch: 12 }, { wch: 60 }, { wch: 12 },
      { wch: 12, hidden: true }, { wch: 12, hidden: true }, { wch: 10, hidden: true },
      { wch: 10, hidden: true }, { wch: 12, hidden: true }, { wch: 12, hidden: true },
    ];
    
    // Применяем форматирование
    applySMRStyles(smrWs);
    
    // Записываем формулы
    if (smrFormulas) {
      for (let r = 0; r < smrFormulas.length; r++) {
        for (let c = 0; c < smrFormulas[r].length; c++) {
          if (smrFormulas[r][c]) {
            const cellRef = XLSX.utils.encode_cell({ r, c });
            if (smrWs[cellRef]) {
              smrWs[cellRef].f = smrFormulas[r][c];
            }
          }
        }
      }
    }
    
    XLSX.utils.book_append_sheet(newWb, smrWs, 'СМР уник');
    
    const tmzWs = XLSX.utils.aoa_to_sheet(tmzData);
    tmzWs['!cols'] = [
      { wch: 12 }, { wch: 50 }, { wch: 10 }, { wch: 12 }, { wch: 12 }, { wch: 12 },
      { wch: 12, hidden: true }, { wch: 12, hidden: true }, { wch: 12, hidden: true },
      { wch: 10, hidden: true }, { wch: 10, hidden: true }, { wch: 12, hidden: true },
      { wch: 12, hidden: true },
    ];
    
    // Применяем форматирование
    applyTMZStyles(tmzWs);
    
    // Записываем формулы
    if (tmzFormulas) {
      for (let r = 0; r < tmzFormulas.length; r++) {
        for (let c = 0; c < tmzFormulas[r].length; c++) {
          if (tmzFormulas[r][c]) {
            const cellRef = XLSX.utils.encode_cell({ r, c });
            if (tmzWs[cellRef]) {
              tmzWs[cellRef].f = tmzFormulas[r][c];
            }
          }
        }
      }
    }
    
    XLSX.utils.book_append_sheet(newWb, tmzWs, 'ТМЦ уник');

    const sheets: SheetData[] = [
      { name: 'Свод', data: svodData, formulas: svodFormulas },
      { name: 'СМР уник', data: smrData, formulas: smrFormulas },
      { name: 'ТМЦ уник', data: tmzData, formulas: tmzFormulas },
    ];

    // Временно отключено - создание листа 1С
    // if (addSheet1C) {
    //   logs.push({ step: 'Лист 1С', status: 'success', message: 'Лист 1С добавлен' });
    // }

    logs.push({ step: 'Завершение', status: 'success', message: `Обработка завершена. Листов: ${sheets.length}` });

    return {
      result: {
        success: true,
        message: `Файл обработан успешно. Создано листов: ${sheets.length}`,
        workbook: newWb,
        sheets
      },
      logs
    };

  } catch (error) {
    logs.push({ step: 'Ошибка', status: 'error', message: `Ошибка обработки: ${(error as Error).message}` });
    return {
      result: { success: false, message: `Ошибка обработки: ${(error as Error).message}` },
      logs
    };
  }
}

export function exportToExcel(wb: XLSX.WorkBook, filename: string = 'Выходная_форма_Результат.xlsx') {
  XLSX.writeFile(wb, filename, { bookType: 'xlsx' });
}
