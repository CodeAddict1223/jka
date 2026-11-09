import { GridData } from '../types/excel';

// Coordinate helpers: 'A1' -> { col: 1, row: 1 }, 'B5' -> { col: 2, row: 5 }
export function colNameToIndex(colName: string): number {
  let index = 0;
  for (let i = 0; i < colName.length; i++) {
    index = index * 26 + (colName.toUpperCase().charCodeAt(i) - 64);
  }
  return index;
}

export function indexToColName(index: number): string {
  let col = '';
  while (index > 0) {
    const rem = (index - 1) % 26;
    col = String.fromCharCode(65 + rem) + col;
    index = Math.floor((index - 1) / 26);
  }
  return col || 'A';
}

export function parseCellAddress(addr: string): { col: string; row: number; colIdx: number } | null {
  const match = addr.trim().toUpperCase().match(/^([A-Z]+)([0-9]+)$/);
  if (!match) return null;
  return {
    col: match[1],
    row: parseInt(match[2], 10),
    colIdx: colNameToIndex(match[1]),
  };
}

// Expand "A1:B3" to ['A1', 'A2', 'A3', 'B1', 'B2', 'B3']
export function expandRange(rangeStr: string): string[] {
  const clean = rangeStr.replace(/\$/g, '').trim().toUpperCase();
  if (!clean.includes(':')) {
    return [clean];
  }
  const [startStr, endStr] = clean.split(':');
  const start = parseCellAddress(startStr);
  const end = parseCellAddress(endStr);
  if (!start || !end) return [];

  const minCol = Math.min(start.colIdx, end.colIdx);
  const maxCol = Math.max(start.colIdx, end.colIdx);
  const minRow = Math.min(start.row, end.row);
  const maxRow = Math.max(start.row, end.row);

  const cells: string[] = [];
  for (let c = minCol; c <= maxCol; c++) {
    const colName = indexToColName(c);
    for (let r = minRow; r <= maxRow; r++) {
      cells.push(`${colName}${r}`);
    }
  }
  return cells;
}

// Get raw or evaluated numeric/string value from cell in grid
export function getCellValue(grid: GridData, cellRef: string): any {
  const key = cellRef.replace(/\$/g, '').toUpperCase();
  const cell = grid[key];
  if (!cell) return '';
  return cell.value;
}

export function getNumericValue(grid: GridData, cellRef: string): number {
  const val = getCellValue(grid, cellRef);
  if (typeof val === 'number') return val;
  if (typeof val === 'string') {
    const clean = val.replace(/[$,%]/g, '').trim();
    const num = parseFloat(clean);
    return isNaN(num) ? 0 : num;
  }
  return 0;
}

// Evaluates formulas entered by the user
export function evaluateFormula(formula: string, grid: GridData): { result: any; error?: string } {
  if (!formula.startsWith('=')) {
    return { result: formula };
  }

  const expr = formula.substring(1).trim();
  const upper = expr.toUpperCase();

  try {
    // 1. Basic Arithmetic if no function name at start e.g. =A1+B1, =B2*1.05
    if (/^[A-Z0-9\$\+\-\*\/\^\(\)\.\s]+$/.test(upper) && !/[A-Z]{2,}\(/.test(upper)) {
      // Replace cell coordinates with numeric values
      const resolved = upper.replace(/\$?[A-Z]+\$?[0-9]+/g, (match) => {
        return String(getNumericValue(grid, match));
      });
      // Replace ^ with **
      const safeExpr = resolved.replace(/\^/g, '**');
      // eslint-disable-next-line no-eval
      const calc = Function(`"use strict"; return (${safeExpr})`)();
      return { result: Math.round(calc * 10000) / 10000 };
    }

    // 2. Parse Function Call e.g. SUM(A1:A5), IF(A1>10, "Yes", "No")
    const funcMatch = expr.match(/^([A-Za-z0-9_\.]+)\((.*)\)$/s);
    if (!funcMatch) {
      return { result: '#VALUE!' };
    }

    const funcName = funcMatch[1].toUpperCase();
    const rawArgs = splitArguments(funcMatch[2]);

    switch (funcName) {
      case 'SUM': {
        let sum = 0;
        for (const arg of rawArgs) {
          const cells = expandRange(arg);
          for (const c of cells) {
            sum += getNumericValue(grid, c);
          }
        }
        return { result: Math.round(sum * 10000) / 10000 };
      }

      case 'AVERAGE': {
        let sum = 0;
        let count = 0;
        for (const arg of rawArgs) {
          const cells = expandRange(arg);
          for (const c of cells) {
            const v = getCellValue(grid, c);
            if (typeof v === 'number' || (typeof v === 'string' && !isNaN(parseFloat(v)) && v.trim() !== '')) {
              sum += getNumericValue(grid, c);
              count++;
            }
          }
        }
        return { result: count === 0 ? 0 : Math.round((sum / count) * 100) / 100 };
      }

      case 'MIN': {
        const vals: number[] = [];
        for (const arg of rawArgs) {
          for (const c of expandRange(arg)) {
            vals.push(getNumericValue(grid, c));
          }
        }
        return { result: vals.length ? Math.min(...vals) : 0 };
      }

      case 'MAX': {
        const vals: number[] = [];
        for (const arg of rawArgs) {
          for (const c of expandRange(arg)) {
            vals.push(getNumericValue(grid, c));
          }
        }
        return { result: vals.length ? Math.max(...vals) : 0 };
      }

      case 'PRODUCT': {
        let prod = 1;
        for (const arg of rawArgs) {
          for (const c of expandRange(arg)) {
            prod *= getNumericValue(grid, c);
          }
        }
        return { result: Math.round(prod * 10000) / 10000 };
      }

      case 'ABS': {
        const val = resolveArg(rawArgs[0], grid);
        return { result: Math.abs(Number(val) || 0) };
      }

      case 'COUNT': {
        let count = 0;
        for (const arg of rawArgs) {
          for (const c of expandRange(arg)) {
            const v = getCellValue(grid, c);
            if (typeof v === 'number' || (typeof v === 'string' && !isNaN(parseFloat(v)) && v.trim() !== '')) {
              count++;
            }
          }
        }
        return { result: count };
      }

      case 'COUNTA': {
        let count = 0;
        for (const arg of rawArgs) {
          for (const c of expandRange(arg)) {
            const v = getCellValue(grid, c);
            if (v !== '' && v !== null && v !== undefined) {
              count++;
            }
          }
        }
        return { result: count };
      }

      case 'COUNTBLANK': {
        let count = 0;
        for (const arg of rawArgs) {
          for (const c of expandRange(arg)) {
            const v = getCellValue(grid, c);
            if (v === '' || v === null || v === undefined) {
              count++;
            }
          }
        }
        return { result: count };
      }

      case 'ROWS': {
        const cells = expandRange(rawArgs[0]);
        if (cells.length === 0) return { result: 0 };
        const rows = new Set(cells.map((c) => parseCellAddress(c)?.row));
        return { result: rows.size };
      }

      case 'COLUMNS': {
        const cells = expandRange(rawArgs[0]);
        if (cells.length === 0) return { result: 0 };
        const cols = new Set(cells.map((c) => parseCellAddress(c)?.col));
        return { result: cols.size };
      }

      case 'ROUND': {
        const val = Number(resolveArg(rawArgs[0], grid)) || 0;
        const decimals = parseInt(resolveArg(rawArgs[1] || '0', grid), 10) || 0;
        const factor = Math.pow(10, decimals);
        return { result: Math.round(val * factor) / factor };
      }

      case 'ROUNDUP': {
        const val = Number(resolveArg(rawArgs[0], grid)) || 0;
        const decimals = parseInt(resolveArg(rawArgs[1] || '0', grid), 10) || 0;
        const factor = Math.pow(10, decimals);
        return { result: Math.ceil(val * factor) / factor };
      }

      case 'ROUNDDOWN': {
        const val = Number(resolveArg(rawArgs[0], grid)) || 0;
        const decimals = parseInt(resolveArg(rawArgs[1] || '0', grid), 10) || 0;
        const factor = Math.pow(10, decimals);
        return { result: Math.floor(val * factor) / factor };
      }

      case 'MROUND': {
        const val = Number(resolveArg(rawArgs[0], grid)) || 0;
        const mult = Number(resolveArg(rawArgs[1], grid)) || 1;
        return { result: Math.round(val / mult) * mult };
      }

      case 'CEILING': {
        const val = Number(resolveArg(rawArgs[0], grid)) || 0;
        const mult = Number(resolveArg(rawArgs[1] || '1', grid)) || 1;
        return { result: Math.ceil(val / mult) * mult };
      }

      case 'FLOOR': {
        const val = Number(resolveArg(rawArgs[0], grid)) || 0;
        const mult = Number(resolveArg(rawArgs[1] || '1', grid)) || 1;
        return { result: Math.floor(val / mult) * mult };
      }

      case 'INT': {
        const val = Number(resolveArg(rawArgs[0], grid)) || 0;
        return { result: Math.floor(val) };
      }

      case 'TRUNC': {
        const val = Number(resolveArg(rawArgs[0], grid)) || 0;
        return { result: Math.trunc(val) };
      }

      case 'ODD': {
        let val = Math.ceil(Number(resolveArg(rawArgs[0], grid)) || 0);
        if (val % 2 === 0) val += val >= 0 ? 1 : -1;
        return { result: val };
      }

      case 'EVEN': {
        let val = Math.ceil(Number(resolveArg(rawArgs[0], grid)) || 0);
        if (Math.abs(val) % 2 !== 0) val += val >= 0 ? 1 : -1;
        return { result: val };
      }

      case 'IF': {
        const condition = evaluateCondition(rawArgs[0], grid);
        const trueVal = resolveArg(rawArgs[1], grid);
        const falseVal = rawArgs[2] !== undefined ? resolveArg(rawArgs[2], grid) : false;
        return { result: condition ? trueVal : falseVal };
      }

      case 'AND': {
        const conds = rawArgs.map((arg) => evaluateCondition(arg, grid));
        return { result: conds.every(Boolean) };
      }

      case 'OR': {
        const conds = rawArgs.map((arg) => evaluateCondition(arg, grid));
        return { result: conds.some(Boolean) };
      }

      case 'NOT': {
        return { result: !evaluateCondition(rawArgs[0], grid) };
      }

      case 'COUNTIF': {
        const range = expandRange(rawArgs[0]);
        const criteria = resolveArg(rawArgs[1], grid);
        let count = 0;
        for (const c of range) {
          const val = getCellValue(grid, c);
          if (matchesCriteria(val, criteria)) count++;
        }
        return { result: count };
      }

      case 'SUMIF': {
        const range = expandRange(rawArgs[0]);
        const criteria = resolveArg(rawArgs[1], grid);
        const sumRange = rawArgs[2] ? expandRange(rawArgs[2]) : range;
        let sum = 0;
        for (let i = 0; i < range.length; i++) {
          const checkVal = getCellValue(grid, range[i]);
          if (matchesCriteria(checkVal, criteria)) {
            const sumCell = sumRange[i] || range[i];
            sum += getNumericValue(grid, sumCell);
          }
        }
        return { result: sum };
      }

      case 'AVERAGEIF': {
        const range = expandRange(rawArgs[0]);
        const criteria = resolveArg(rawArgs[1], grid);
        const avgRange = rawArgs[2] ? expandRange(rawArgs[2]) : range;
        let sum = 0;
        let count = 0;
        for (let i = 0; i < range.length; i++) {
          const checkVal = getCellValue(grid, range[i]);
          if (matchesCriteria(checkVal, criteria)) {
            const avgCell = avgRange[i] || range[i];
            sum += getNumericValue(grid, avgCell);
            count++;
          }
        }
        return { result: count ? Math.round((sum / count) * 100) / 100 : 0 };
      }

      case 'SUMPRODUCT': {
        // e.g. SUMPRODUCT(A2:A4, B2:B4)
        const range1 = expandRange(rawArgs[0]);
        const range2 = expandRange(rawArgs[1]);
        let total = 0;
        for (let i = 0; i < Math.min(range1.length, range2.length); i++) {
          total += getNumericValue(grid, range1[i]) * getNumericValue(grid, range2[i]);
        }
        return { result: total };
      }

      case 'VLOOKUP': {
        // VLOOKUP(lookup_value, table_array, col_index, [exact_match])
        const lookup = resolveArg(rawArgs[0], grid);
        const tableRange = rawArgs[1];
        const colIdx = parseInt(resolveArg(rawArgs[2], grid), 10);
        const isExact = rawArgs[3] ? resolveArg(rawArgs[3], grid) === false || String(resolveArg(rawArgs[3], grid)).toLowerCase() === 'false' || resolveArg(rawArgs[3], grid) === 0 : true;

        const [startCell, endCell] = tableRange.replace(/\$/g, '').split(':');
        const start = parseCellAddress(startCell);
        const end = parseCellAddress(endCell);
        if (!start || !end) return { result: '#REF!' };

        const startCol = start.colIdx;
        const targetColName = indexToColName(startCol + colIdx - 1);
        const startColName = indexToColName(startCol);

        for (let r = start.row; r <= end.row; r++) {
          const valInFirstCol = getCellValue(grid, `${startColName}${r}`);
          if (String(valInFirstCol).toLowerCase() === String(lookup).toLowerCase()) {
            return { result: getCellValue(grid, `${targetColName}${r}`) };
          }
        }
        return { result: '#N/A' };
      }

      case 'XLOOKUP': {
        // XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found])
        const lookup = resolveArg(rawArgs[0], grid);
        const lookupCells = expandRange(rawArgs[1]);
        const returnCells = expandRange(rawArgs[2]);
        const notFound = rawArgs[3] ? resolveArg(rawArgs[3], grid) : '#N/A';

        for (let i = 0; i < lookupCells.length; i++) {
          const val = getCellValue(grid, lookupCells[i]);
          if (String(val).toLowerCase() === String(lookup).toLowerCase()) {
            return { result: getCellValue(grid, returnCells[i]) };
          }
        }
        return { result: notFound };
      }

      case 'INDEX': {
        // INDEX(array, row_num, [col_num])
        const cells = expandRange(rawArgs[0]);
        const rowNum = parseInt(resolveArg(rawArgs[1], grid), 10);
        return { result: getCellValue(grid, cells[rowNum - 1]) };
      }

      case 'MATCH': {
        // MATCH(lookup_value, lookup_array, [match_type])
        const lookup = resolveArg(rawArgs[0], grid);
        const cells = expandRange(rawArgs[1]);
        for (let i = 0; i < cells.length; i++) {
          if (String(getCellValue(grid, cells[i])).toLowerCase() === String(lookup).toLowerCase()) {
            return { result: i + 1 };
          }
        }
        return { result: '#N/A' };
      }

      case 'CONCAT': {
        let str = '';
        for (const arg of rawArgs) {
          for (const c of expandRange(arg)) {
            str += String(getCellValue(grid, c) ?? resolveArg(arg, grid));
          }
        }
        return { result: str };
      }

      case 'TEXTJOIN': {
        // TEXTJOIN(delimiter, ignore_empty, text1, text2...)
        const delim = resolveArg(rawArgs[0], grid);
        const ignoreEmpty = Boolean(resolveArg(rawArgs[1], grid));
        const parts: string[] = [];
        for (let i = 2; i < rawArgs.length; i++) {
          for (const c of expandRange(rawArgs[i])) {
            const v = getCellValue(grid, c);
            if (v !== '' || !ignoreEmpty) parts.push(String(v));
          }
        }
        return { result: parts.join(delim) };
      }

      case 'LEFT': {
        const text = String(resolveArg(rawArgs[0], grid));
        const len = parseInt(resolveArg(rawArgs[1] || '1', grid), 10);
        return { result: text.substring(0, len) };
      }

      case 'RIGHT': {
        const text = String(resolveArg(rawArgs[0], grid));
        const len = parseInt(resolveArg(rawArgs[1] || '1', grid), 10);
        return { result: text.substring(text.length - len) };
      }

      case 'MID': {
        const text = String(resolveArg(rawArgs[0], grid));
        const start = parseInt(resolveArg(rawArgs[1], grid), 10);
        const len = parseInt(resolveArg(rawArgs[2], grid), 10);
        return { result: text.substring(start - 1, start - 1 + len) };
      }

      case 'LEN': {
        const text = String(resolveArg(rawArgs[0], grid));
        return { result: text.length };
      }

      case 'UPPER': {
        return { result: String(resolveArg(rawArgs[0], grid)).toUpperCase() };
      }

      case 'LOWER': {
        return { result: String(resolveArg(rawArgs[0], grid)).toLowerCase() };
      }

      case 'PROPER': {
        const text = String(resolveArg(rawArgs[0], grid));
        return {
          result: text.replace(/\w\S*/g, (txt) => txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase()),
        };
      }

      case 'TRIM': {
        return { result: String(resolveArg(rawArgs[0], grid)).trim().replace(/\s+/g, ' ') };
      }

      case 'IFERROR': {
        const res = evaluateFormula('=' + rawArgs[0], grid);
        if (res.result === '#N/A' || res.result === '#VALUE!' || res.result === '#DIV/0!' || res.error) {
          return { result: resolveArg(rawArgs[1], grid) };
        }
        return res;
      }

      case 'TODAY': {
        const now = new Date();
        return { result: `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}` };
      }

      case 'YEAR': {
        const dateStr = String(resolveArg(rawArgs[0], grid));
        const d = new Date(dateStr);
        return { result: isNaN(d.getFullYear()) ? 2026 : d.getFullYear() };
      }

      case 'PMT': {
        // PMT(rate, nper, pv)
        const rate = Number(resolveArg(rawArgs[0], grid)) || 0;
        const nper = Number(resolveArg(rawArgs[1], grid)) || 1;
        const pv = Number(resolveArg(rawArgs[2], grid)) || 0;
        if (rate === 0) return { result: -(pv / nper) };
        const pmt = (pv * rate * Math.pow(1 + rate, nper)) / (Math.pow(1 + rate, nper) - 1);
        return { result: -Math.round(pmt * 100) / 100 };
      }

      default:
        return { result: `#NAME?` };
    }
  } catch (err: any) {
    return { result: '#ERROR!', error: err.message };
  }
}

// Helpers for arguments parsing
function splitArguments(argStr: string): string[] {
  const args: string[] = [];
  let depth = 0;
  let inQuotes = false;
  let current = '';

  for (let i = 0; i < argStr.length; i++) {
    const char = argStr[i];
    if (char === '"' && argStr[i - 1] !== '\\') {
      inQuotes = !inQuotes;
      current += char;
    } else if (char === '(' && !inQuotes) {
      depth++;
      current += char;
    } else if (char === ')' && !inQuotes) {
      depth--;
      current += char;
    } else if (char === ',' && depth === 0 && !inQuotes) {
      args.push(current.trim());
      current = '';
    } else {
      current += char;
    }
  }
  if (current.trim()) {
    args.push(current.trim());
  }
  return args;
}

function resolveArg(arg: string, grid: GridData): any {
  if (!arg) return '';
  const trimmed = arg.trim();

  // String literal
  if (trimmed.startsWith('"') && trimmed.endsWith('"')) {
    return trimmed.slice(1, -1);
  }

  // Number literal
  if (!isNaN(Number(trimmed))) {
    return Number(trimmed);
  }

  // Boolean
  if (trimmed.toUpperCase() === 'TRUE') return true;
  if (trimmed.toUpperCase() === 'FALSE') return false;

  // Coordinate reference
  if (/^\$?[A-Z]+\$?[0-9]+$/i.test(trimmed)) {
    return getCellValue(grid, trimmed);
  }

  // Nested formula evaluation
  if (trimmed.includes('(')) {
    const subRes = evaluateFormula('=' + trimmed, grid);
    return subRes.result;
  }

  return trimmed;
}

function evaluateCondition(condStr: string, grid: GridData): boolean {
  // Support operators: =, <>, >=, <=, >, <
  const operators = ['>=', '<=', '<>', '=', '>', '<'];
  for (const op of operators) {
    const idx = condStr.indexOf(op);
    if (idx !== -1) {
      const left = resolveArg(condStr.substring(0, idx), grid);
      const right = resolveArg(condStr.substring(idx + op.length), grid);
      const leftNum = Number(left);
      const rightNum = Number(right);
      const isNum = !isNaN(leftNum) && !isNaN(rightNum);

      switch (op) {
        case '=':
          return String(left).toLowerCase() === String(right).toLowerCase();
        case '<>':
          return String(left).toLowerCase() !== String(right).toLowerCase();
        case '>=':
          return isNum ? leftNum >= rightNum : String(left) >= String(right);
        case '<=':
          return isNum ? leftNum <= rightNum : String(left) <= String(right);
        case '>':
          return isNum ? leftNum > rightNum : String(left) > String(right);
        case '<':
          return isNum ? leftNum < rightNum : String(left) < String(right);
      }
    }
  }
  // Fallback boolean
  return Boolean(resolveArg(condStr, grid));
}

function matchesCriteria(val: any, criteria: any): boolean {
  const critStr = String(criteria).trim();
  if (critStr.startsWith('>=')) return Number(val) >= Number(critStr.slice(2));
  if (critStr.startsWith('<=')) return Number(val) <= Number(critStr.slice(2));
  if (critStr.startsWith('<>')) return String(val).toLowerCase() !== critStr.slice(2).toLowerCase();
  if (critStr.startsWith('>')) return Number(val) > Number(critStr.slice(1));
  if (critStr.startsWith('<')) return Number(val) < Number(critStr.slice(1));
  if (critStr.startsWith('=')) return String(val).toLowerCase() === critStr.slice(1).toLowerCase();

  // Wildcards *
  if (critStr.includes('*')) {
    const regex = new RegExp('^' + critStr.replace(/\*/g, '.*') + '$', 'i');
    return regex.test(String(val));
  }

  return String(val).toLowerCase() === critStr.toLowerCase();
}
