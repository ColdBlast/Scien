/**
 * Complete 3D Spreadsheet Engine with Range Functions (AVERAGE, STDEV),
 * Powers, SQRT, Quadrature Error Propagation, Unit Simplification,
 * and Temperature Reading/Difference Formatting (°C, °F, °R, K).
 */

function onOpen() {
  const mode = getUncertaintyMode();
  const modeLabel = mode === 'WORST_CASE' ? 'Worst Case' : 'Random (RSS)';

  SpreadsheetApp.getUi()
    .createMenu('Scien Tools')
    .addItem(`Toggle Uncertainty Mode (Current: ${modeLabel})`, 'toggleUncertaintyMode')
    .addSeparator()
    .addItem('Undo Last 3D Edit', 'undoLastAction')
    .addItem('Re-Sync All Sheets', 'syncBackendSheets')
    .addItem('Restore Formula for Editing', 'restoreFormulaForActiveCell')
    .addItem('Plot Selected Data (X, Y)', 'openPlotDialog')
    .addToUi();
}

function getUncertaintyMode() {
  if (cachedUncertaintyMode) return cachedUncertaintyMode;
  cachedUncertaintyMode = PropertiesService.getDocumentProperties().getProperty('UNCERTAINTY_MODE') || 'QUADRATURE';
  return cachedUncertaintyMode;
}

let cachedUncertaintyMode = null;

function toggleUncertaintyMode() {
  const props = PropertiesService.getDocumentProperties();
  const currentMode = getUncertaintyMode();
  const newMode = (currentMode === 'WORST_CASE') ? 'QUADRATURE' : 'WORST_CASE';
  
  props.setProperty('UNCERTAINTY_MODE', newMode);
  cachedUncertaintyMode = newMode;
  
  const label = (newMode === 'WORST_CASE') 
    ? 'Worst Case (Linear Sum)' 
    : 'Random Error (Quadrature / RSS)';
    
  SpreadsheetApp.getActiveSpreadsheet().toast(`Mode changed to: ${label}`, 'Scien Settings', 5);
  
  onOpen();
  syncBackendSheets();
}

function onEdit(e) {
  if (!e || !e.range) return;

  const range = e.range;
  const sheet = range.getSheet();
  
  if (sheet.getName() !== 'Input Interface') return;
  if (range.getNumRows() > 1 || range.getNumColumns() > 1) return;

  const row = range.getRow();
  const col = range.getColumn();

  saveUndoState(e, row, col);

  const formula = range.getFormula();
  const rawInput = String(e.value || range.getDisplayValue()).trim();

  if (formula) {
    processFormulaEntry(e, row, col, formula);
    return;
  }

  if (!rawInput) {
    range.clearNote();
    updateTargetSheets(row, col, '', '', '', '');
    return;
  }

  range.clearNote();
  const parsed = parseMeasurement(rawInput);
  if (parsed) {
    const siUnits = convertToSIBase(parsed.unit);
    updateTargetSheets(row, col, parsed.value, parsed.uncertainty, siUnits, rawInput);
    
    const displayString = formatMeasurementDisplay(parsed.value, parsed.uncertainty, parsed.unit, parsed.uncStr);
    range.setValue(displayString);
  }
}

function processFormulaEntry(e, targetRow, targetCol, formulaStr) {
  const cleanFormula = formulaStr.replace(/^[=+]/, '').trim();
  const result = evaluate3DExpression(cleanFormula);

  if (result.error) {
    updateTargetSheets(targetRow, targetCol, 'ERR', 'ERR', 'ERR', formulaStr);
    e.range.setNote(`Error in ${formulaStr}: ${result.error}`);
    e.range.setValue('#UNIT!');
    return;
  }

  updateTargetSheets(targetRow, targetCol, result.v, result.u, result.unit, formulaStr);
  e.range.setNote(`Calculated from: ${formulaStr}`);

  const displayUnit = simplifyAndConvertUnits(result.unit);
  const displayStr = formatMeasurementDisplay(result.v, result.u, displayUnit, null, result.isDifference);
  e.range.setValue(displayStr);
}

function evaluate3DExpression(expression) {
  let cleanExpr = expression.replace(/\s+/g, '').trim();

  const funcResults = {};
  let funcIndex = 0;
  
  let rangeFuncPattern = /([A-Z.]+)\(([A-Z]+\d+:[A-Z]+\d+)\)/i;
  let match;
  
  while ((match = cleanExpr.match(rangeFuncPattern))) {
    const funcName = match[1].toUpperCase();
    const argStr = match[2];
    
    const rangeData = get3DRangeData(argStr);
    if (rangeData.error) return rangeData;
    
    const items = rangeData.items;
    if (items.length === 0) return { error: `Empty range provided in ${funcName}` };

    let resultObj;
    if (funcName === 'AVERAGE') {
      resultObj = calcAverage(items);
    } else if (['STDEV', 'STDEV.S', 'STDEV.P'].includes(funcName)) {
      const isPopulation = (funcName === 'STDEV.P');
      resultObj = calcStdev(items, isPopulation);
    } else {
      return { error: `Unknown function: ${funcName}` };
    }
    
    if (resultObj.error) return resultObj;

    const placeholder = `_FUNC_${funcIndex}_`;
    funcResults[placeholder] = resultObj;
    
    cleanExpr = cleanExpr.replace(match[0], placeholder);
    funcIndex++;
  }

  let singleFuncPattern = /([A-Z.]+)\(([A-Z]+\d+)\)/i;
  while ((match = cleanExpr.match(singleFuncPattern))) {
    const funcName = match[1].toUpperCase();
    const cellRef = match[2];
    
    const cellData = get3DCellData(cellRef);
    if (!cellData) return { error: `Cell ${cellRef} not found` };

    let resultObj;
    if (funcName === 'SQRT') {
      resultObj = Propagate3D.sqrt(cellData);
    } else {
      return { error: `Unknown function: ${funcName}` };
    }

    if (resultObj.error) return resultObj;

    const placeholder = `_FUNC_${funcIndex}_`;
    funcResults[placeholder] = resultObj;
    
    cleanExpr = cleanExpr.replace(match[0], placeholder);
    funcIndex++;
  }

  cleanExpr = cleanExpr.replace(/(^|\()(-)/g, '$10$2');
  
  const tokens = cleanExpr.match(/(_FUNC_\d+_|[A-Z]+\d+|(?:\d*\.\d+|\d+)|[+\-*/^()])/gi);
  if (!tokens) return { error: "Invalid expression syntax" };

  const precedence = { '+': 1, '-': 1, '*': 2, '/': 2, '^': 3 };
  const associativity = { '^': 'right' };
  
  const output = [];
  const opStack = [];

  for (let token of tokens) {
    const upperToken = token.toUpperCase();
    if (upperToken.startsWith('_FUNC_')) {
       output.push(funcResults[upperToken]);
    } else if (/^[A-Z]+\d+$/.test(upperToken)) {
      const cellData = get3DCellData(upperToken);
      if (!cellData) return { error: `Cell ${upperToken} not found` };
      output.push(cellData);
    } else if (!isNaN(parseFloat(token))) {
      output.push({ v: parseFloat(token), u: 0, unit: '' }); 
    } else if (['+', '-', '*', '/', '^'].includes(token)) {
      while (
        opStack.length > 0 && 
        opStack[opStack.length - 1] !== '(' &&
        (
          precedence[opStack[opStack.length - 1]] > precedence[token] ||
          (precedence[opStack[opStack.length - 1]] === precedence[token] && associativity[token] !== 'right')
        )
      ) {
        output.push(opStack.pop());
      }
      opStack.push(token);
    } else if (token === '(') {
      opStack.push(token);
    } else if (token === ')') {
      while (opStack.length > 0 && opStack[opStack.length - 1] !== '(') {
        output.push(opStack.pop());
      }
      if (opStack.length === 0 || opStack[opStack.length - 1] !== '(') {
        return { error: "Mismatched parentheses" };
      }
      opStack.pop(); 
    }
  }

  while (opStack.length > 0) {
    const op = opStack.pop();
    if (op === '(' || op === ')') return { error: "Mismatched parentheses" };
    output.push(op);
  }

  const evalStack = [];
  for (let token of output) {
    if (typeof token === 'object') {
      evalStack.push(token);
    } else {
      if (evalStack.length < 2) return { error: "Incomplete mathematical equation" };
      const b = evalStack.pop();
      const a = evalStack.pop();
      
      let res;
      if (token === '+') res = Propagate3D.add(a, b);
      else if (token === '-') res = Propagate3D.subtract(a, b);
      else if (token === '*') res = Propagate3D.multiply(a, b);
      else if (token === '/') res = Propagate3D.divide(a, b);
      else if (token === '^') res = Propagate3D.power(a, b.v);

      if (res && res.error) return res;
      evalStack.push(res);
    }
  }

  if (evalStack.length !== 1) return { error: "Could not evaluate expression completely" };
  return evalStack[0];
}

const Propagate3D = {
  add: function(a, b) {
    if (!areUnitsCompatibleForAddSub(a.unit, b.unit)) return { error: `Cannot add ${a.unit || 'none'} and ${b.unit || 'none'}` };
    const mode = getUncertaintyMode();
    const u = (mode === 'WORST_CASE') 
      ? Math.abs(a.u) + Math.abs(b.u) 
      : Math.sqrt(Math.pow(a.u, 2) + Math.pow(b.u, 2));
    return { v: a.v + b.v, u: u, unit: resolveAddSubUnit(a.unit, b.unit, false) };
  },

  subtract: function(a, b) {
    if (!areUnitsCompatibleForAddSub(a.unit, b.unit)) return { error: `Cannot subtract ${b.unit || 'none'} from ${a.unit || 'none'}` };
    const mode = getUncertaintyMode();
    const u = (mode === 'WORST_CASE') 
      ? Math.abs(a.u) + Math.abs(b.u) 
      : Math.sqrt(Math.pow(a.u, 2) + Math.pow(b.u, 2));
    // Subtracting absolute temperatures produces a temperature difference
    const isTemp = /°?[CFRK]/i.test(a.unit);
    return { v: a.v - b.v, u: u, unit: resolveAddSubUnit(a.unit, b.unit, true) };
  },

  multiply: function(a, b) {
    const val = a.v * b.v;
    const relA = a.v !== 0 ? a.u / a.v : 0;
    const relB = b.v !== 0 ? b.u / b.v : 0;
    
    const mode = getUncertaintyMode();
    const relResult = (mode === 'WORST_CASE') 
      ? Math.abs(relA) + Math.abs(relB) 
      : Math.sqrt(Math.pow(relA, 2) + Math.pow(relB, 2));
      
    const unc = Math.abs(val) * relResult;
    const newUnit = UnitMath.format(UnitMath.operate(a.unit, b.unit, 'multiply'));
    return { v: val, u: unc, unit: newUnit };
  },

  divide: function(a, b) {
    if (b.v === 0) return { error: 'Division by zero' };
    const val = a.v / b.v;
    const relA = a.v !== 0 ? a.u / a.v : 0;
    const relB = b.v !== 0 ? b.u / b.v : 0;
    
    const mode = getUncertaintyMode();
    const relResult = (mode === 'WORST_CASE') 
      ? Math.abs(relA) + Math.abs(relB) 
      : Math.sqrt(Math.pow(relA, 2) + Math.pow(relB, 2));
      
    const unc = Math.abs(val) * relResult;
    const newUnit = UnitMath.format(UnitMath.operate(a.unit, b.unit, 'divide'));
    return { v: val, u: unc, unit: newUnit };
  },

  power: function(a, n) {
    if (a.v === 0 && n < 0) return { error: 'Division by zero in power' };
    const val = Math.pow(a.v, n);
    const relA = a.v !== 0 ? a.u / a.v : 0;
    const unc = Math.abs(val) * Math.abs(n) * Math.abs(relA);
    const newUnit = UnitMath.format(UnitMath.power(a.unit, n));
    return { v: val, u: unc, unit: newUnit };
  },

  sqrt: function(a) {
    if (a.v < 0) return { error: 'Square root of negative number' };
    return this.power(a, 0.5);
  }
};

function get3DRangeData(rangeStr) {
  const match = rangeStr.toUpperCase().match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
  if (!match) return { error: `Invalid range format: ${rangeStr}` };

  const colStart = match[1];
  const rowStart = parseInt(match[2], 10);
  const colEnd = match[3];
  const rowEnd = parseInt(match[4], 10);

  const items = [];
  let baseUnit = null;

  if (colStart !== colEnd) return { error: "Multi-column ranges not currently supported" };

  const colNum = columnNameToIndex(colStart);
  for (let r = rowStart; r <= rowEnd; r++) {
    const a1 = indexToColumnName(colNum) + r;
    const data = get3DCellData(a1);
    if (!data || data.v === null || data.v === '') continue;

    if (baseUnit === null) baseUnit = data.unit;
    else if (data.unit !== baseUnit) return { error: `Unit mismatch in range: found '${data.unit}' and '${baseUnit}'` };

    items.push(data);
  }

  return { items: items, unit: baseUnit };
}

function calcAverage(items) {
  const n = items.length;
  let sumVal = 0;
  let sumSqUnc = 0;
  const unit = items[0].unit;

  for (let item of items) {
    sumVal += item.v;
    sumSqUnc += Math.pow(item.u, 2);
  }

  const avgVal = sumVal / n;
  const avgUnc = (1 / n) * Math.sqrt(sumSqUnc);

  return { v: avgVal, u: avgUnc, unit: unit, isDifference: items[0].isDifference };
}

function calcStdev(items, isPopulation) {
  const n = items.length;
  const divisor = isPopulation ? n : (n - 1);
  if (divisor <= 0) return { error: "Not enough data points for standard deviation calculation" };

  const unit = items[0].unit;
  let sumVal = 0;
  for (let item of items) sumVal += item.v;
  const mean = sumVal / n;

  let sumSqDiff = 0;
  let sumSqUnc = 0;
  for (let item of items) {
    const diff = item.v - mean;
    sumSqDiff += Math.pow(diff, 2);
    const dStdev_dx = diff / (Math.sqrt(divisor * sumSqDiff) || 1);
    sumSqUnc += Math.pow(dStdev_dx * item.u, 2);
  }

  const stdevVal = Math.sqrt(sumSqDiff / divisor);
  const stdevUnc = Math.sqrt(sumSqUnc);

  return { v: stdevVal, u: stdevUnc, unit: unit, isDifference: items[0].isDifference };
}

/**
 * Formats temperature units based on reading vs. difference rules.
 */
function formatTemperatureUnit(unitStr) {
  if (!unitStr) return '';
  return unitStr.trim();
}

function formatMeasurementDisplay(value, uncertainty, unit, rawUncStr = null, isDifference = false) {
  const formattedUnit = formatTemperatureUnit(unit);

  if (uncertainty === undefined || uncertainty === null || uncertainty === 0) {
    return `${value}` + (formattedUnit ? ` ${formattedUnit}` : '');
  }

  const absUnc = Math.abs(uncertainty);
  
  const [mantissaStr, expStr] = absUnc.toExponential().split('e');
  const firstDigit = parseInt(mantissaStr.replace('.', '')[0], 10);
  const magnitude = parseInt(expStr, 10);

  const maxSF = (firstDigit === 1 || firstDigit === 2) ? 2 : 1;
  let sf = maxSF;

  if (rawUncStr) {
    let str = String(rawUncStr).trim().replace(/^[+-]/, '').split(/e/i)[0];
    let digits = str.includes('.') 
      ? str.replace(/^0+\.?0*/, '').replace('.', '') 
      : str.replace(/^0+/, '');
    let typedSF = digits.length > 0 ? digits.length : 1;

    sf = Math.min(typedSF, maxSF);
  }

  const dp = sf - 1 - magnitude;
  const factor = Math.pow(10, dp);
  const roundedUnc = Math.round(absUnc * factor) / factor;
  const roundedVal = Math.round(value * factor) / factor;

  let valStr, uncStr;
  if (dp > 0) {
    valStr = roundedVal.toFixed(dp);
    uncStr = roundedUnc.toFixed(dp);
  } else {
    valStr = roundedVal.toFixed(0);
    uncStr = roundedUnc.toFixed(0);
  }

  return `${valStr} ± ${uncStr}` + (formattedUnit ? ` ${formattedUnit}` : '');
}

function saveUndoState(e, row, col) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const valSheet = ss.getSheetByName('Values');
    const uncSheet = ss.getSheetByName('Uncertainties');
    const unitSheet = ss.getSheetByName('Units');
    const formSheet = ss.getSheetByName('Formulas');

    const oldUiValue = e.oldValue !== undefined ? e.oldValue : '';
    const oldNote = e.range.getNote() || '';
    
    const undoData = {
      row: row,
      col: col,
      uiValue: oldUiValue,
      uiNote: oldNote,
      val: getSafeBackendValue(valSheet, row, col),
      unc: getSafeBackendValue(uncSheet, row, col),
      unit: getSafeBackendValue(unitSheet, row, col),
      form: getSafeBackendValue(formSheet, row, col)
    };

    PropertiesService.getDocumentProperties().setProperty('LAST_3D_UNDO_STATE', JSON.stringify(undoData));
  } catch (err) {
    console.error("Failed to save undo state: " + err);
  }
}

function undoLastAction() {
  const docProps = PropertiesService.getDocumentProperties();
  const rawState = docProps.getProperty('LAST_3D_UNDO_STATE');

  if (!rawState) {
    SpreadsheetApp.getUi().alert('No recent 3D edit available to undo.');
    return;
  }

  try {
    const state = JSON.parse(rawState);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const inputSheet = ss.getSheetByName('Input Interface');
    if (!inputSheet) return;

    const targetRange = inputSheet.getRange(state.row, state.col);
    targetRange.setValue(state.uiValue);
    if (state.uiNote) targetRange.setNote(state.uiNote);
    else targetRange.clearNote();

    updateTargetSheets(state.row, state.col, state.val, state.unc, state.unit, state.form);
    docProps.deleteProperty('LAST_3D_UNDO_STATE');
    
    SpreadsheetApp.flush(); 
    ss.toast('Last action restored successfully.', 'Undo Complete', 3);
  } catch (err) {
    SpreadsheetApp.getUi().alert('Error undoing action: ' + err.message);
  }
}

function restoreFormulaForActiveCell() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  
  if (sheet.getName() !== 'Input Interface') return;

  const cell = sheet.getActiveCell();
  const row = cell.getRow();
  const col = cell.getColumn();

  let formulaToRestore = '';
  const note = cell.getNote();
  
  if (note && note.includes('Calculated from: ')) {
    formulaToRestore = note.replace('Calculated from: ', '').trim();
  } else if (note && note.includes('Error in ')) {
    const match = note.match(/^Error in\s+(.+?):\s/i);
    if (match) {
      formulaToRestore = match[1].trim();
    } else {
      formulaToRestore = note.replace('Error in ', '').trim();
    }
  }

  if (!formulaToRestore) {
    const formSheet = ss.getSheetByName('Formulas');
    if (formSheet) {
      formulaToRestore = String(formSheet.getRange(row, col).getValue() || '').trim();
      if (formulaToRestore.startsWith("'")) {
        formulaToRestore = formulaToRestore.substring(1);
      }
    }
  }

  if (formulaToRestore) {
    cell.setValue(formulaToRestore);
  }
}

function syncBackendSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName('Input Interface');
  if (!inputSheet) return;

  const range = inputSheet.getDataRange();
  const values = range.getValues();
  const formulas = range.getFormulas();
  const notes = range.getNotes();

  ['Values', 'Uncertainties', 'Units', 'Formulas'].forEach(tabName => {
    const sheet = ss.getSheetByName(tabName);
    if (sheet) sheet.clearContents();
  });

  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      const cellValue = String(values[r][c] || '').trim();
      const cellFormula = formulas[r][c];
      const cellNote = notes[r][c];

      if (!cellValue) continue;

      const targetRow = r + 1;
      const targetCol = c + 1;

      let rawFormula = cellFormula;
      if (cellNote && cellNote.includes('Calculated from: ')) {
        rawFormula = cellNote.replace('Calculated from: ', '').trim();
      }

      if (rawFormula) {
        processFormulaEntry({ range: inputSheet.getRange(targetRow, targetCol) }, targetRow, targetCol, rawFormula);
      } else {
        const parsed = parseMeasurement(cellValue);
        if (parsed) {
          const siUnits = convertToSIBase(parsed.unit);
          updateTargetSheets(targetRow, targetCol, parsed.value, parsed.uncertainty, siUnits, cellValue);
          
          const displayString = formatMeasurementDisplay(parsed.value, parsed.uncertainty, parsed.unit, parsed.uncStr, parsed.isDifference);
          inputSheet.getRange(targetRow, targetCol).setValue(displayString);
        }
      }
    }
  }
  SpreadsheetApp.getUi().alert('Backend tabs successfully re-synced!');
}

function getSafeBackendValue(sheet, row, col) {
  if (!sheet || row > sheet.getMaxRows() || col > sheet.getMaxColumns()) return '';
  return sheet.getRange(row, col).getValue();
}

function safeSetBackendValue(sheet, row, col, value) {
  if (!sheet) return;
  try {
    const maxRows = sheet.getMaxRows();
    const maxCols = sheet.getMaxColumns();
    if (row > maxRows) sheet.insertRowsAfter(maxRows, row - maxRows + 5);
    if (col > maxCols) sheet.insertColumnsAfter(maxCols, col - maxCols + 2);
    sheet.getRange(row, col).setValue(value);
  } catch (e) {
    try { sheet.getRange(row, col).setValue(value); } catch(err) {}
  }
}

function updateTargetSheets(row, col, value, uncertainty, unit, formula = '') {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const valSheet = ss.getSheetByName('Values');
  const uncSheet = ss.getSheetByName('Uncertainties');
  const unitSheet = ss.getSheetByName('Units');
  const formSheet = ss.getSheetByName('Formulas');

  safeSetBackendValue(valSheet, row, col, value);
  safeSetBackendValue(uncSheet, row, col, uncertainty);
  safeSetBackendValue(unitSheet, row, col, unit);

  if (formSheet) {
    const storedText = (formula.startsWith('=') || formula.startsWith('+')) ? `'${formula}` : formula;
    safeSetBackendValue(formSheet, row, col, storedText);
  }
}

function get3DCellData(a1Notation) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const valSheet = ss.getSheetByName('Values');
  const uncSheet = ss.getSheetByName('Uncertainties');
  const unitSheet = ss.getSheetByName('Units');

  if (!valSheet || !uncSheet || !unitSheet) return null;

  const unit = String(unitSheet.getRange(a1Notation).getValue() || '').trim();

  return {
    v: Number(valSheet.getRange(a1Notation).getValue()) || 0,
    u: Number(uncSheet.getRange(a1Notation).getValue()) || 0,
    unit: unit,
  };
}

/**
 * Normalizes temperature units on input (readings vs. differences)
 * and converts 'deg' or 'o' text fallbacks.
 */
function normalizeTemperatureUnits(unitStr) {
  if (!unitStr) return '';
  let u = unitStr.trim();
  
  // 1. Temperature differences (Suffix °): Cdeg, Co, Fdeg, Fo, Rdeg, Ro, Kdeg, Ko -> C°, F°, R°, K°
  u = u.replace(/\b([CFRK])(?:°|deg|o|\u00B0)\b/gi, (m, p1) => p1.toUpperCase() + '°');
  
  // 2. Temperature readings (Prefix °): degC, oC, degF, oF, degR, oR -> °C, °F, °R
  u = u.replace(/\b(?:°|deg|o|\u00B0)([CFR])\b/gi, (m, p1) => '°' + p1.toUpperCase());
  
  // 3. Kelvin readings (No °): degK, oK -> K
  u = u.replace(/\b(?:°|deg|o|\u00B0)K\b/gi, 'K');
  
  // 4. Clean up any literal °K entered by mistake back to standard K
  u = u.replace(/°K\b/gi, 'K');
  
  return u;
}

/**
 * Maps a temperature unit (both absolute and difference) to its scale family base.
 * Celsius and Kelvin share a scale size. Fahrenheit and Rankine share a scale size.
 *
 * @param {string} unitStr - The normalized unit string (e.g., "°C", "K°", "m/s")
 * @return {string} The family identifier or the original unit if not a temperature.
 */
function getTemperatureFamilyBase(unitStr) {
  if (!unitStr) return '';
  let u = unitStr.trim().toUpperCase();
  
  // Check if it's a temperature unit
  const isTemp = u.includes('°') || u === 'K';
  if (!isTemp) return unitStr;
  
  // Strip out the degree symbol and spaces to isolate the scale letter
  const baseLetter = u.replace(/°/g, '').trim();
  
  if (baseLetter === 'C' || baseLetter === 'K') {
    return 'TEMP_CELSIUS_KELVIN';
  }
  if (baseLetter === 'F' || baseLetter === 'R') {
    return 'TEMP_FAHRENHEIT_RANKINE';
  }
  
  return unitStr;
}

/**
 * Checks if two units are identical or belong to the same temperature scale family.
 */
function areUnitsCompatibleForAddSub(unit1, unit2) {
  if (!unit1 && !unit2) return true;
  
  const family1 = getTemperatureFamilyBase(unit1);
  const family2 = getTemperatureFamilyBase(unit2);
  
  if (family1.startsWith('TEMP_') && family2.startsWith('TEMP_')) {
    return family1 === family2;
  }
  
  return unit1 === unit2;
}

/**
 * Determines the resulting unit when adding or subtracting two compatible units.
 * Applies thermodynamic rules for temperature absolutes vs differences.
 * 
 * @param {string} unit1 - The first normalized unit (e.g., "K°", "°C")
 * @param {string} unit2 - The second normalized unit
 * @param {boolean} isSubtraction - True if the operation is subtraction
 * @return {string} The resolved output unit
 */
function resolveAddSubUnit(unit1, unit2, isSub) {
  if (!unit1 || !unit2) return unit1 || unit2 || '';

  const fam = getTemperatureFamilyBase(unit1);
  if (!fam.startsWith('TEMP_')) return unit1;

  // Safe difference check that ignores hidden trailing spaces
  const isDiff1 = /([CFRK])°/i.test(unit1);
  const isDiff2 = /([CFRK])°/i.test(unit2);

  if (!isSub) {
    // ADDITION: If either unit is Absolute, the result is Absolute
    return (!isDiff1) ? unit1 : (!isDiff2) ? unit2 : unit1;
  } else {
    // SUBTRACTION: Absolute - Absolute = Difference (e.g. °C - °C = C°)
    if (!isDiff1 && !isDiff2) {
      const baseLetter = unit1.replace(/[^CFRK]/g, '');
      return baseLetter + '°';
    }
    // Absolute - Difference = Absolute; Difference - Difference = Difference
    return unit1;
  }
}

function parseMeasurement(inputStr) {
  const regex = /^([+-]?\d*\.?\d+(?:[eE][+-]?\d+)?)\s*(?:±|\+\/-|\+-)\s*([+-]?\d*\.?\d+(?:[eE][+-]?\d+)?)\s*(.*)$/;
  const match = inputStr.match(regex);
  
  if (!match) return null;
  
  const rawUnit = match[3] ? match[3].trim() : '';
  const normalizedUnit = normalizeTemperatureUnits(rawUnit);
  
  return {
    value: parseFloat(match[1]),
    uncertainty: match[2] ? parseFloat(match[2]) : 0,
    uncStr: match[2] ? match[2].trim() : '',
    unit: normalizedUnit,
  };
}

function convertToSIBase(unitStr) {
  if (!unitStr) return '';
  const siDerivedMap = {
    'N': 'kg·m/s²', 'J': 'kg·m²/s²', 'W': 'kg·m²/s³',
    'Pa': 'kg/(m·s²)', 'Hz': '1/s',
    'V': 'kg·m²/(A·s³)', 'F': 'A²·s⁴/(kg·m²)', 'Ω': 'kg·m²/(A·s³)'
  };
  let normalized = unitStr;
  for (const [derived, base] of Object.entries(siDerivedMap)) {
    normalized = normalized.replace(new RegExp(`\\b${derived}\\b`, 'g'), base);
  }
  // Convert Coulombs 'C' safely without altering °C or C° temperature units
  normalized = normalized.replace(/(?<![°])\bC\b(?![°])/g, 'A·s');
  return normalized;
}

function simplifyAndConvertUnits(unitStr) {
  if (!unitStr) return '';
  let u = unitStr.replace(/kg·m\/s²·m/g, 'kg·m²/s²').replace(/m·m/g, 'm²').replace(/s·s/g, 's²');
  const baseToDerivedMap = [
    { base: 'kg·m²/s²', derived: 'J' }, { base: 'kg·m²/s³', derived: 'W' },
    { base: 'kg·m/s²', derived: 'N' }, { base: 'kg/(m·s²)', derived: 'Pa' },
    { base: 'kg·m²/(A·s³)', derived: 'V' }, { base: 'A²·s⁴/(kg·m²)', derived: 'F' },
    { base: 'kg·m²/(A²·s³)', derived: 'Ω' }, { base: 'A·s', derived: 'C' },
    { base: '1/s', derived: 'Hz' }
  ];
  for (const item of baseToDerivedMap) {
    u = u.replace(new RegExp(item.base.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&'), 'g'), item.derived);
  }
  return u;
}

function columnNameToIndex(name) {
  let index = 0;
  for (let i = 0; i < name.length; i++) {
    index = index * 26 + (name.charCodeAt(i) - 64);
  }
  return index;
}

function indexToColumnName(index) {
  let name = '';
  while (index > 0) {
    let remainder = (index - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    index = Math.floor((index - 1) / 26);
  }
  return name;
}

function openPlotDialog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  
  if (sheet.getName() !== 'Input Interface') {
    SpreadsheetApp.getUi().alert('Please select data on the Input Interface tab.');
    return;
  }

  const range = sheet.getActiveRange();
  if (range.getNumColumns() !== 2) {
    SpreadsheetApp.getUi().alert('Please select exactly two adjacent columns (X on the left, Y on the right).');
    return;
  }

  const row = range.getRow();
  const col = range.getColumn();
  const numRows = range.getNumRows();

  const valSheet = ss.getSheetByName('Values');
  const uncSheet = ss.getSheetByName('Uncertainties');
  const unitSheet = ss.getSheetByName('Units');

  const xVals = valSheet.getRange(row, col, numRows, 1).getValues().flat();
  const yVals = valSheet.getRange(row, col + 1, numRows, 1).getValues().flat();
  const xUncs = uncSheet.getRange(row, col, numRows, 1).getValues().flat();
  const yUncs = uncSheet.getRange(row, col + 1, numRows, 1).getValues().flat();

  const xUnits = unitSheet.getRange(row, col, numRows, 1).getValues().flat().find(u => u) || '';
  const yUnits = unitSheet.getRange(row, col + 1, numRows, 1).getValues().flat().find(u => u) || '';

  const plotData = { xVals, yVals, xUncs, yUncs, xUnits, yUnits };

  const htmlTemplate = HtmlService.createTemplateFromFile('PlotDialog');
  htmlTemplate.plotData = JSON.stringify(plotData);

  const html = htmlTemplate.evaluate()
    .setWidth(850)
    .setHeight(650)
    .setTitle('SciEng Data Plot');

  SpreadsheetApp.getUi().showModalDialog(html, 'Data Plot with Uncertainties');
}

function applyPowerToUnit(unitString, power) {
  if (!unitString) return "";

  let normalized = String(unitString)
    .replace(/²/g, '^2')
    .replace(/³/g, '^3')
    .replace(/⁻¹/g, '^-1');
normalizeunit
  let result = normalized.replace(/([a-zA-Z°]+)(?:\^([-\d.]+))?/g, (match, base, currentExp) => {
    let exp = currentExp ? parseFloat(currentExp) : 1;
    let newExp = exp * power;
    
    if (newExp === 1) return base;
    if (newExp === 0) return "";
    return base + '^' + newExp;
  });

  return result
    .replace(/\^2(?!\d)/g, '²')
    .replace(/\^3(?!\d)/g, '³')
    .replace(/\^-1(?!\d)/g, '⁻¹');
}

const UnitMath = {
  parse: function(unitStr) {
    if (!unitStr) return {};
    const dict = {};
    
    const norm = String(unitStr).replace(/²/g, '^2').replace(/³/g, '^3').replace(/⁻¹/g, '^-1');
    const parts = norm.split('/');
    this._extractTerms(parts[0], 1, dict);
    
    if (parts.length > 1) {
      const denom = parts.slice(1).join('*'); 
      this._extractTerms(denom, -1, dict);
    }
    return this._clean(dict);
  },
  
  _extractTerms: function(str, sign, dict) {
    const regex = /([a-zA-Z°Ω]+)(?:\^([-\d.]+))?/g;
    let match;
    while ((match = regex.exec(str)) !== null) {
      const base = match[1];
      const exp = match[2] ? parseFloat(match[2]) : 1;
      dict[base] = (dict[base] || 0) + (exp * sign);
    }
  },

  _clean: function(dict) {
    for (let k in dict) {
      if (Math.abs(dict[k]) < 1e-10) delete dict[k];
    }
    return dict;
  },

  operate: function(u1, u2, operation) {
    const d1 = typeof u1 === 'string' ? this.parse(u1) : u1;
    const d2 = typeof u2 === 'string' ? this.parse(u2) : u2;
    const res = { ...d1 };
    
    for (let k in d2) {
      if (operation === 'multiply') res[k] = (res[k] || 0) + d2[k];
      else if (operation === 'divide') res[k] = (res[k] || 0) - d2[k];
    }
    return this._clean(res);
  },

  power: function(u, n) {
    const d = typeof u === 'string' ? this.parse(u) : u;
    const res = {};
    for (let k in d) res[k] = d[k] * n;
    return this._clean(res);
  },

  format: function(dict) {
    const num = [], den = [];
    
    for (let k of Object.keys(dict).sort()) {
      const exp = dict[k];
      if (exp > 0) num.push(exp === 1 ? k : `${k}^${exp}`);
      else if (exp < 0) den.push(Math.abs(exp) === 1 ? k : `${k}^${Math.abs(exp)}`);
    }
    
    const numStr = num.length > 0 ? num.join('·') : (den.length > 0 ? '1' : '');
    const denStr = den.length > 0 ? '/' + den.join('·') : '';
    const rawStr = numStr + denStr;
    
    const standardUnits = {
      'N·m': 'J', 'A·V': 'W', 'W·s': 'J', 'C·V': 'J', 'A·s': 'C', 'A·Ω': 'V'
    };
    
    if (standardUnits[rawStr]) return standardUnits[rawStr];

    return rawStr.replace(/\^2(?!\d)/g, '²')
                 .replace(/\^3(?!\d)/g, '³')
                 .replace(/\^-1(?!\d)/g, '⁻¹');
  }
};
