import { SlideData, GridData, CellFormat } from '../types/excel';
import { indexToColName, colNameToIndex } from './formulaEngine';

/**
 * Dynamic Question Generator Engine
 * Generates dynamically varied Excel practice problems with randomized coordinates,
 * values, data layouts (horizontal/vertical), bounds, constraints, and objectives.
 * Implements pedagogical patterns from user references:
 * 1. Cell Identification & Basic Formulas (Screenshot 190424, 190605)
 * 2. Range Functions (SUM / Ranges horizontal & vertical) (Screenshot 190553, 190710)
 * 3. Keyboard Navigation & Selection (Shift+Arrow, Ctrl+Arrow, gap jumping, block selection) (Screenshot 190754, 190816, 190844, 190904)
 */

// Random helper within range [min, max]
function randInt(min: number, max: number): number {
  return Math.floor(Math.random() * (max - min + 1)) + min;
}

// Pick random item from array
function pickRandom<T>(arr: T[]): T {
  return arr[Math.floor(Math.random() * arr.length)];
}

// Generate a random distinct seed string
export function generateVariationKey(): string {
  return `${Date.now().toString(36)}-${Math.random().toString(36).substring(2, 7)}`;
}

/**
 * Dynamically transforms a slide into a fresh, varied practice problem
 */
export function generateDynamicSlideVariation(baseSlide: SlideData, variationSeed?: number): SlideData {
  const seed = variationSeed ?? randInt(1, 1000);
  const slideId = baseSlide.id;

  // 1. PATTERN: Cell Addressing & Identification (Screenshot 190605)
  if (slideId === 'm1-cell-addressing' || slideId.includes('cell-address')) {
    const cols = ['B', 'C', 'D', 'E', 'F'];
    const rows = [2, 3, 4, 5, 6, 7];
    const targetCol = pickRandom(cols);
    const targetRow = pickRandom(rows);
    const targetCell = `${targetCol}${targetRow}`;

    const newGrid: GridData = {
      [targetCell]: {
        value: '',
        format: { bold: true, bgColor: '#1c3d2e' },
      },
    };

    return {
      ...baseSlide,
      title: 'Active Cell Identification',
      practiceQuestion: {
        instruction: `What is the coordinate address of the highlighted active cell shown in the grid? Type its column letter and row number (e.g. ${targetCell}).`,
        hint: `Look at the illuminated column header (${targetCol}) and row header (${targetRow}). Coordinates are written as Column Letter then Row Number.`,
        targetCell,
        practiceType: 'data_entry',
      },
      excel: {
        ...baseSlide.excel,
        defaultActiveCell: targetCell,
        initialGrid: newGrid,
      },
      solution: {
        expectedValue: [targetCell, targetCell.toLowerCase(), `${targetCol} ${targetRow}`],
        revealAnswerText: `Active Cell: ${targetCell}`,
        explanationText: `The active cell is located at the intersection of Column ${targetCol} and Row ${targetRow}, giving the cell address ${targetCell}.`,
        autofillAction: {
          targetCell,
          value: targetCell,
          format: { bold: true },
        },
      },
    };
  }

  // 2. PATTERN: Basic Formulas & Cell References (Screenshot 190424)
  if (slideId === 'm6-operators' || slideId === 'm1-formula-bar') {
    // Dynamically choose layout: Vertical or Horizontal or Shifted coordinates
    const isVertical = seed % 2 === 0;
    const num1 = randInt(12, 65);
    const num2 = randInt(10, 50);
    const op = pickRandom(['+', '*', '-']);

    let cell1 = 'A1';
    let cell2 = 'A2';
    let targetCell = 'A3';
    let label = 'Sum';

    if (isVertical) {
      const col = pickRandom(['B', 'C', 'D']);
      const startRow = randInt(2, 4);
      cell1 = `${col}${startRow}`;
      cell2 = `${col}${startRow + 1}`;
      targetCell = `${col}${startRow + 2}`;
    } else {
      const row = randInt(2, 5);
      const startColIdx = randInt(2, 3); // B or C
      cell1 = `${indexToColName(startColIdx)}${row}`;
      cell2 = `${indexToColName(startColIdx + 1)}${row}`;
      targetCell = `${indexToColName(startColIdx + 2)}${row}`;
    }

    let expectedVal = 0;
    let opSymbol = '+';
    if (op === '+') {
      expectedVal = num1 + num2;
      label = 'Total';
    } else if (op === '*') {
      expectedVal = num1 * num2;
      opSymbol = '*';
      label = 'Product';
    } else {
      expectedVal = Math.max(num1, num2) - Math.min(num1, num2);
      opSymbol = '-';
      label = 'Variance';
    }

    const grid: GridData = {
      [cell1]: { value: num1 },
      [cell2]: { value: num2 },
      [targetCell]: { value: '' },
    };

    const expectedFormula1 = `=${cell1}${opSymbol}${cell2}`;
    const expectedFormula2 = `=${cell2}${opSymbol}${cell1}`;

    return {
      ...baseSlide,
      practiceQuestion: {
        instruction: `In cell ${targetCell}, calculate the ${label.toLowerCase()} of cells ${cell1} (${num1}) and ${cell2} (${num2}) using the formula =${cell1}${opSymbol}${cell2}.`,
        hint: `Select cell ${targetCell}, type =${cell1}${opSymbol}${cell2} and press Enter.`,
        targetCell,
        practiceType: 'formula',
      },
      excel: {
        ...baseSlide.excel,
        defaultActiveCell: targetCell,
        initialGrid: grid,
      },
      solution: {
        expectedFormula: [expectedFormula1, expectedFormula2],
        expectedValue: expectedVal,
        revealAnswerText: `Formula: ${expectedFormula1}`,
        explanationText: `Excel references cells ${cell1} and ${cell2} to calculate ${num1} ${opSymbol} ${num2} = ${expectedVal}.`,
        autofillAction: {
          targetCell,
          formula: expectedFormula1,
          value: expectedVal,
          format: { bold: true },
        },
      },
    };
  }

  // 3. PATTERN: Range Functions (SUM / Ranges horizontal & vertical) (Screenshot 190553 & 190710)
  if (slideId === 'm6-sum-function' || slideId === 'm6-autosum' || slideId === 'm6-min-max') {
    const isHorizontal = seed % 2 === 0;
    const isMax = slideId === 'm6-min-max';
    const funcName = isMax ? 'MAX' : 'SUM';

    if (isHorizontal) {
      // Horizontal Range like Screenshot 190710: e.g. B3 to E3, numbers in row 3, target in E5 or F3
      const row = randInt(2, 4);
      const startColIdx = randInt(2, 3); // B or C
      const cellCount = randInt(3, 4);
      const endColIdx = startColIdx + cellCount - 1;

      const startCol = indexToColName(startColIdx);
      const endCol = indexToColName(endColIdx);
      const startCell = `${startCol}${row}`;
      const endCell = `${endCol}${row}`;
      const targetCell = `${endCol}${row + 2}`;

      const grid: GridData = {};
      const nums: number[] = [];

      for (let c = startColIdx; c <= endColIdx; c++) {
        const cName = indexToColName(c);
        const val = randInt(2, 25);
        nums.push(val);
        grid[`${cName}${row}`] = { value: val };
      }
      grid[targetCell] = { value: '' };

      const computedVal = isMax ? Math.max(...nums) : nums.reduce((a, b) => a + b, 0);
      const expectedFormula = `=${funcName}(${startCell}:${endCell})`;

      return {
        ...baseSlide,
        practiceQuestion: {
          instruction: `In cell ${targetCell}, enter a formula using the ${funcName} function to add up every cell from ${startCell} to ${endCell}: =${funcName}(${startCell}:${endCell}).`,
          hint: `Click cell ${targetCell}, type =${funcName}(${startCell}:${endCell}) and press Enter.`,
          targetCell,
          practiceType: 'formula',
          requiredShortcut: slideId === 'm6-autosum' ? 'Alt + =' : undefined,
        },
        excel: {
          ...baseSlide.excel,
          defaultActiveCell: targetCell,
          initialGrid: grid,
        },
        solution: {
          expectedFormula: [
            expectedFormula,
            `=${funcName}(${endCell}:${startCell})`,
            `=${funcName.toLowerCase()}(${startCell}:${endCell})`,
          ],
          expectedValue: computedVal,
          revealAnswerText: `Formula: ${expectedFormula}`,
          explanationText: `=${funcName}(${startCell}:${endCell}) computes the aggregate across cells [${nums.join(', ')}], resulting in ${computedVal}.`,
          autofillAction: {
            targetCell,
            formula: expectedFormula,
            value: computedVal,
            format: { bold: true },
          },
        },
      };
    } else {
      // Vertical Range like Screenshot 190553: e.g. B1:B4 or C2:C6
      const col = pickRandom(['B', 'C', 'D']);
      const startRow = randInt(1, 2);
      const rowCount = randInt(3, 5);
      const endRow = startRow + rowCount - 1;
      const targetRow = endRow + 1;

      const startCell = `${col}${startRow}`;
      const endCell = `${col}${endRow}`;
      const targetCell = `${col}${targetRow}`;

      const grid: GridData = {};
      const nums: number[] = [];

      // Add a header above if startRow > 1
      if (startRow > 1) {
        grid[`${col}1`] = { value: 'Scores', format: { bold: true } };
      }

      for (let r = startRow; r <= endRow; r++) {
        const val = randInt(10, 95);
        nums.push(val);
        grid[`${col}${r}`] = { value: val };
      }
      grid[targetCell] = { value: '' };

      const computedVal = isMax ? Math.max(...nums) : nums.reduce((a, b) => a + b, 0);
      const expectedFormula = `=${funcName}(${startCell}:${endCell})`;

      return {
        ...baseSlide,
        practiceQuestion: {
          instruction: `In cell ${targetCell}, enter a formula using the ${funcName} function to calculate all cells from ${startCell} to ${endCell}: =${funcName}(${startCell}:${endCell}).`,
          hint: `Select ${targetCell}, type =${funcName}(${startCell}:${endCell}) and press Enter.`,
          targetCell,
          practiceType: 'formula',
          requiredShortcut: slideId === 'm6-autosum' ? 'Alt + =' : undefined,
        },
        excel: {
          ...baseSlide.excel,
          defaultActiveCell: targetCell,
          initialGrid: grid,
        },
        solution: {
          expectedFormula: [
            expectedFormula,
            `=${funcName}(${endCell}:${startCell})`,
            `=${funcName.toLowerCase()}(${startCell}:${endCell})`,
          ],
          expectedValue: computedVal,
          revealAnswerText: `Formula: ${expectedFormula}`,
          explanationText: `Evaluates ${funcName} across rows ${startRow} through ${endRow} (${nums.join(' + ')}) = ${computedVal}.`,
          autofillAction: {
            targetCell,
            formula: expectedFormula,
            value: computedVal,
            format: { bold: true },
          },
        },
      };
    }
  }

  // 4. PATTERN: Dynamic Selection with Shift + Arrows (Screenshot 190754 & 190844)
  if (slideId === 'm2-shift-arrows') {
    const isVertical = seed % 2 === 0;

    if (isVertical) {
      // Vertical selection with exclusion (Screenshot 190844)
      const col = pickRandom(['B', 'C', 'D', 'E']);
      const startRow = randInt(2, 3);
      const count = 2; // select first two
      const stopRow = startRow + count - 1;
      const excludeRow = stopRow + 1;

      const startCell = `${col}${startRow}`;
      const stopCell = `${col}${stopRow}`;
      const targetRange = `${startCell}:${stopCell}`;

      const grid: GridData = {
        [startCell]: { value: 'Select me', format: { bold: true } },
        [`${col}${stopRow}`]: { value: 'and me' },
        [`${col}${excludeRow}`]: { value: 'BUT NOT ME', format: { textColor: '#f87171' } },
      };

      return {
        ...baseSlide,
        practiceQuestion: {
          instruction: `Select the first two cells (${startCell}:${stopCell}) using the keyboard (Hold Shift and press Down Arrow). Do not select ${col}${excludeRow}!`,
          hint: `Click ${startCell}, hold Shift, and press Down Arrow once to select down to ${stopCell}.`,
          targetRange,
          practiceType: 'selection',
          requiredShortcut: 'Shift + Down Arrow',
        },
        excel: {
          ...baseSlide.excel,
          defaultActiveCell: startCell,
          initialGrid: grid,
        },
        solution: {
          expectedShortcut: 'Shift + Down Arrow',
          expectedRange: targetRange,
          revealAnswerText: `Range: ${targetRange} (Shift + Down Arrow)`,
          explanationText: `Holding Shift and pressing Down Arrow expands your selection to include ${startCell} and ${stopCell} while safely leaving ${col}${excludeRow} unselected.`,
          autofillAction: {
            targetCell: startCell,
            value: 'Select me',
          },
        },
      };
    } else {
      // Horizontal selection (Screenshot 190754)
      const row = randInt(2, 4);
      const startColIdx = randInt(2, 3); // B or C
      const cellCount = randInt(2, 3);
      const endColIdx = startColIdx + cellCount - 1;

      const startCol = indexToColName(startColIdx);
      const endCol = indexToColName(endColIdx);
      const startCell = `${startCol}${row}`;
      const endCell = `${endCol}${row}`;
      const targetRange = `${startCell}:${endCell}`;

      const grid: GridData = {
        [startCell]: { value: 'Select me', format: { bold: true } },
      };
      for (let c = startColIdx + 1; c <= endColIdx; c++) {
        grid[`${indexToColName(c)}${row}`] = { value: c === endColIdx ? 'Me too!' : 'and me' };
      }

      return {
        ...baseSlide,
        practiceQuestion: {
          instruction: `Hold Shift and press Right Arrow to extend the selection across all populated cells from ${startCell} to ${endCell}.`,
          hint: `Start on ${startCell}, hold Shift, and press Right Arrow until ${endCell} is highlighted.`,
          targetRange,
          practiceType: 'selection',
          requiredShortcut: 'Shift + Right Arrow',
        },
        excel: {
          ...baseSlide.excel,
          defaultActiveCell: startCell,
          initialGrid: grid,
        },
        solution: {
          expectedShortcut: 'Shift + Right Arrow',
          expectedRange: targetRange,
          revealAnswerText: `Range: ${targetRange} (Shift + Right Arrow)`,
          explanationText: `Shift + Right Arrow dynamically expands selection cell by cell across the row.`,
          autofillAction: {
            targetCell: startCell,
            value: 'Select me',
          },
        },
      };
    }
  }

  // 5. PATTERN: Edge-of-Data Jumping with Ctrl + Arrows & Gap Testing (Screenshot 190816)
  if (slideId === 'm2-ctrl-arrows') {
    const isHorizontal = seed % 2 === 0;

    if (isHorizontal) {
      // Horizontal jump: e.g. B2 to E2 with "Jump here"
      const row = randInt(2, 4);
      const startColIdx = randInt(2, 3); // B or C
      const length = randInt(3, 4);
      const endColIdx = startColIdx + length - 1;

      const startCol = indexToColName(startColIdx);
      const endCol = indexToColName(endColIdx);
      const startCell = `${startCol}${row}`;
      const edgeCell = `${endCol}${row}`;

      const grid: GridData = {};
      for (let c = startColIdx; c < endColIdx; c++) {
        grid[`${indexToColName(c)}${row}`] = { value: c === startColIdx ? 'Start' : 'text' };
      }
      grid[edgeCell] = { value: 'Jump here', format: { bold: true, textColor: '#34d399' } };

      return {
        ...baseSlide,
        practiceQuestion: {
          instruction: `Starting at cell ${startCell}, press Ctrl + Right Arrow to jump immediately to the last populated cell at the edge of the data (${edgeCell}).`,
          hint: `Select ${startCell}, hold Ctrl, and press the Right Arrow key (or click Simulate Key).`,
          targetCell: edgeCell,
          practiceType: 'shortcut',
          requiredShortcut: 'Ctrl + Right Arrow',
        },
        excel: {
          ...baseSlide.excel,
          defaultActiveCell: startCell,
          initialGrid: grid,
        },
        solution: {
          expectedShortcut: 'Ctrl + Right Arrow',
          expectedRange: edgeCell,
          revealAnswerText: `Shortcut: Ctrl + Right Arrow (jumps from ${startCell} to ${edgeCell})`,
          explanationText: `Ctrl + Right Arrow detects continuous data along Row ${row} and leaps straight to the boundary cell ${edgeCell}.`,
          autofillAction: {
            targetCell: edgeCell,
            value: 'Jump here',
            format: { bold: true },
          },
        },
      };
    } else {
      // Vertical jump with variable depth
      const col = pickRandom(['B', 'C', 'D']);
      const startRow = randInt(1, 2);
      const depth = randInt(4, 6);
      const endRow = startRow + depth - 1;

      const startCell = `${col}${startRow}`;
      const edgeCell = `${col}${endRow}`;

      const grid: GridData = {};
      grid[startCell] = { value: 'Staff', format: { bold: true } };

      const sampleNames = ['Alice', 'Bob', 'Charlie', 'Dana', 'Evan', 'Fiona', 'Grace'];
      for (let r = startRow + 1; r < endRow; r++) {
        grid[`${col}${r}`] = { value: pickRandom(sampleNames) };
      }
      grid[edgeCell] = { value: 'Last Member', format: { bold: true, textColor: '#34d399' } };

      return {
        ...baseSlide,
        practiceQuestion: {
          instruction: `Starting with cell ${startCell} active, press Ctrl + Down Arrow to jump directly to the bottom edge cell (${edgeCell}).`,
          hint: `Click ${startCell}, then press Ctrl + Down Arrow.`,
          targetCell: edgeCell,
          practiceType: 'shortcut',
          requiredShortcut: 'Ctrl + Down Arrow',
        },
        excel: {
          ...baseSlide.excel,
          defaultActiveCell: startCell,
          initialGrid: grid,
        },
        solution: {
          expectedShortcut: 'Ctrl + Down Arrow',
          expectedRange: edgeCell,
          revealAnswerText: `Shortcut: Ctrl + Down Arrow (jumps from ${startCell} to ${edgeCell})`,
          explanationText: `Ctrl + Down Arrow moves through all populated rows and halts at the final contiguous cell ${edgeCell}.`,
          autofillAction: {
            targetCell: edgeCell,
            value: 'Last Member',
            format: { bold: true },
          },
        },
      };
    }
  }

  // 6. PATTERN: Block / Table Region Selection (Screenshot 190904)
  if (slideId === 'm2-ctrl-a') {
    // 3x3 or 4x3 block of cells
    const startColIdx = randInt(2, 3);
    const startRow = randInt(2, 3);
    const numCols = randInt(3, 4);
    const numRows = randInt(3, 4);

    const endColIdx = startColIdx + numCols - 1;
    const endRow = startRow + numRows - 1;

    const startCell = `${indexToColName(startColIdx)}${startRow}`;
    const endCell = `${indexToColName(endColIdx)}${endRow}`;
    const targetRange = `${startCell}:${endCell}`;

    const grid: GridData = {};
    for (let c = startColIdx; c <= endColIdx; c++) {
      const cName = indexToColName(c);
      for (let r = startRow; r <= endRow; r++) {
        grid[`${cName}${r}`] = { value: r === startRow ? `Col ${c - startColIdx + 1}` : 'data' };
      }
    }

    return {
      ...baseSlide,
      practiceQuestion: {
        instruction: `Select all the populated cells in the data region (${targetRange}) using the keyboard shortcut Ctrl + A (or Ctrl + Shift + Down).`,
        hint: `Click inside the table region and press Ctrl + A.`,
        targetRange,
        practiceType: 'shortcut',
        requiredShortcut: 'Ctrl + A',
      },
      excel: {
        ...baseSlide.excel,
        defaultActiveCell: startCell,
        initialGrid: grid,
      },
      solution: {
        expectedShortcut: 'Ctrl + A',
        expectedRange: targetRange,
        revealAnswerText: `Shortcut: Ctrl + A (selects ${targetRange})`,
        explanationText: `Ctrl + A senses the surrounding continuous table boundary and highlights the entire region in one keystroke.`,
        autofillAction: {
          targetCell: startCell,
          value: 'data',
        },
      },
    };
  }

  // 7. PATTERN: Fill Down (Ctrl + D) with variable rows and formulas
  if (slideId === 'm3-fill-down') {
    const colA = pickRandom(['A', 'B']);
    const colB = indexToColName(colNameToIndex(colA) + 1);
    const colC = indexToColName(colNameToIndex(colA) + 2);
    const startRow = randInt(2, 3);
    const nextRow = startRow + 1;

    const price1 = randInt(50, 200);
    const rate1 = 0.08;
    const price2 = randInt(150, 400);
    const rate2 = 0.08;

    const formulaAbove = `=${colA}${startRow}*${colB}${startRow}`;
    const formulaTarget = `=${colA}${nextRow}*${colB}${nextRow}`;
    const expectedVal = Math.round(price2 * rate2 * 100) / 100;

    const grid: GridData = {
      [`${colA}1`]: { value: 'Price', format: { bold: true } },
      [`${colB}1`]: { value: 'Tax Rate', format: { bold: true } },
      [`${colC}1`]: { value: 'Tax Amt', format: { bold: true } },
      [`${colA}${startRow}`]: { value: price1 },
      [`${colB}${startRow}`]: { value: rate1 },
      [`${colC}${startRow}`]: { value: Math.round(price1 * rate1 * 100) / 100, formula: formulaAbove },
      [`${colA}${nextRow}`]: { value: price2 },
      [`${colB}${nextRow}`]: { value: rate2 },
      [`${colC}${nextRow}`]: { value: '' },
    };

    const targetCell = `${colC}${nextRow}`;

    return {
      ...baseSlide,
      practiceQuestion: {
        instruction: `Select cell ${targetCell} directly below ${colC}${startRow}, then press Ctrl + D (or click Simulate Key) to Fill Down the formula from cell ${colC}${startRow} into cell ${targetCell}.`,
        hint: `Select cell ${targetCell}, then press Ctrl + D.`,
        targetCell,
        practiceType: 'shortcut',
        requiredShortcut: 'Ctrl + D',
      },
      excel: {
        ...baseSlide.excel,
        defaultActiveCell: targetCell,
        initialGrid: grid,
      },
      solution: {
        expectedShortcut: 'Ctrl + D',
        expectedFormula: [formulaTarget, `=${colB}${nextRow}*${colA}${nextRow}`],
        expectedValue: expectedVal,
        revealAnswerText: `Shortcut: Ctrl + D in cell ${targetCell} (produces ${formulaTarget})`,
        explanationText: `Ctrl + D copies the formula downward and adjusts relative row coordinates from row ${startRow} to row ${nextRow}.`,
        autofillAction: {
          targetCell,
          formula: formulaTarget,
          value: expectedVal,
        },
      },
    };
  }

  // 8. PATTERN: Percentage Formatting (Ctrl + Shift + %) with randomized decimals
  if (slideId === 'm3-number-formats') {
    const rawVal = pickRandom([0.15, 0.285, 0.34, 0.425, 0.68, 0.75]);
    const targetCell = pickRandom(['B2', 'B3', 'C2']);
    const pctInt = Math.round(rawVal * 100);
    const pctStr = `${pctInt}%`;

    const grid: GridData = {
      A1: { value: 'Metric', format: { bold: true } },
      B1: { value: 'Rate', format: { bold: true } },
      A2: { value: 'Margin' },
      [targetCell]: { value: String(rawVal) },
    };

    return {
      ...baseSlide,
      practiceQuestion: {
        instruction: `Select cell ${targetCell} (raw rate ${rawVal}), then press Ctrl + Shift + % (or click Simulate Key) to apply percentage formatting.`,
        hint: `Select cell ${targetCell}, then press Ctrl + Shift + % (or click Simulate Key).`,
        targetCell,
        practiceType: 'shortcut',
        requiredShortcut: 'Ctrl + Shift + %',
      },
      excel: {
        ...baseSlide.excel,
        defaultActiveCell: targetCell,
        initialGrid: grid,
      },
      solution: {
        expectedShortcut: 'Ctrl + Shift + %',
        expectedValue: [pctStr, `${rawVal * 100}%`, String(rawVal), rawVal],
        revealAnswerText: `Shortcut: Ctrl + Shift + % (formats as ${pctStr})`,
        explanationText: `Ctrl + Shift + % multiplies the decimal by 100 and applies the percentage symbol.`,
        autofillAction: {
          targetCell,
          value: pctStr,
          format: { format: 'percentage' },
        },
      },
    };
  }

  // 9. PATTERN: Row Deletion with dynamic row position
  if (slideId === 'm4-insert-delete') {
    const deleteRow = randInt(3, 4);
    const targetCell = `A${deleteRow}`;

    const grid: GridData = {
      A1: { value: 'Item', format: { bold: true } },
      B1: { value: 'Units', format: { bold: true } },
      A2: { value: 'Widget Alpha' },
      B2: { value: randInt(40, 100) },
      [`A${deleteRow}`]: { value: 'Duplicate / Unwanted Entry' },
      [`B${deleteRow}`]: { value: 0 },
      [`A${deleteRow + 1}`]: { value: 'Widget Beta' },
      [`B${deleteRow + 1}`]: { value: randInt(80, 150) },
    };

    return {
      ...baseSlide,
      practiceQuestion: {
        instruction: `Select row ${deleteRow} containing "Duplicate / Unwanted Entry" (cell ${targetCell}), then press Ctrl + - to delete the redundant row.`,
        hint: `Select ${targetCell}, press Shift + Space to select the row, then press Ctrl + - (or click Simulate Key).`,
        targetCell,
        practiceType: 'shortcut',
        requiredShortcut: 'Ctrl + -',
      },
      excel: {
        ...baseSlide.excel,
        defaultActiveCell: targetCell,
        initialGrid: grid,
      },
      solution: {
        expectedShortcut: 'Ctrl + -',
        revealAnswerText: `Shortcut: Ctrl + - on Row ${deleteRow}`,
        explanationText: `Ctrl + - removes the active row and shifts the remaining data rows upward.`,
        autofillAction: {
          targetCell,
          value: 'Widget Beta',
        },
      },
    };
  }

  // 10. Default: For general formula slides, randomize numbers slightly if numeric
  const newGrid: GridData = {};
  let numMultiplier = 1;

  // If initialGrid has simple numbers, randomize with a safe factor
  Object.entries(baseSlide.excel.initialGrid).forEach(([key, cell]) => {
    if (typeof cell.value === 'number' && cell.value > 10 && !cell.formula && key !== baseSlide.practiceQuestion.targetCell) {
      const randomizedNum = Math.round(cell.value * (0.8 + Math.random() * 0.4));
      newGrid[key] = { ...cell, value: randomizedNum };
    } else {
      newGrid[key] = { ...cell };
    }
  });

  return {
    ...baseSlide,
    excel: {
      ...baseSlide.excel,
      initialGrid: newGrid,
    },
  };
}

/**
 * Adaptive Live Feedback Evaluator
 * Evaluates live student input against target objectives and returns guidance hints.
 */
export interface AdaptiveLiveFeedback {
  status: 'empty' | 'typing_formula' | 'syntax_warning' | 'range_incomplete' | 'ready_to_submit' | 'shortcut_hint';
  message: string;
  suggestion?: string;
}

export function evaluateLiveInput(
  input: string,
  slide: SlideData,
  selectedRange: string[],
  activeCell: string
): AdaptiveLiveFeedback {
  const trimmed = input.trim();
  const practiceType = slide.practiceQuestion.practiceType;

  // 1. Shortcut Practice
  if (practiceType === 'shortcut') {
    const required = slide.practiceQuestion.requiredShortcut || '';
    return {
      status: 'shortcut_hint',
      message: `Required action: Press ${required} on your keyboard or click the on-screen shortcut button.`,
      suggestion: required,
    };
  }

  // 2. Selection Practice
  if (practiceType === 'selection') {
    const targetRange = slide.practiceQuestion.targetRange || slide.solution.expectedRange;
    const currentSelectionStr = selectedRange.length > 1 ? `${selectedRange[0]}:${selectedRange[selectedRange.length - 1]}` : (activeCell || selectedRange[0]);
    if (targetRange && currentSelectionStr.toUpperCase() === targetRange.toUpperCase()) {
      return {
        status: 'ready_to_submit',
        message: `✓ Perfect! You highlighted range ${targetRange}. Click Check or press Enter.`,
      };
    }
    return {
      status: 'range_incomplete',
      message: `Currently selected: ${currentSelectionStr}. Target: ${targetRange}. Hold Shift and press arrow keys to expand.`,
    };
  }

  // 3. Empty input
  if (!trimmed) {
    return {
      status: 'empty',
      message: `Type directly in cell ${slide.practiceQuestion.targetCell || activeCell} or use the Formula Bar.`,
    };
  }

  // 4. Formula Analysis
  if (practiceType === 'formula') {
    if (!trimmed.startsWith('=')) {
      return {
        status: 'syntax_warning',
        message: `Every Excel formula must start with an equals sign '='. Example: =SUM(...)`,
        suggestion: `=${trimmed}`,
      };
    }

    const upper = trimmed.toUpperCase();

    // Check open parenthesis
    if (/[A-Z]{2,}/.test(upper) && !upper.includes('(')) {
      return {
        status: 'typing_formula',
        message: `Good function name! Now open parenthesis '(' to specify your cell arguments.`,
      };
    }

    // Check unclosed parenthesis
    const openParens = (upper.match(/\(/g) || []).length;
    const closeParens = (upper.match(/\)/g) || []).length;
    if (openParens > closeParens) {
      return {
        status: 'typing_formula',
        message: `Formula parenthesis is open. Don't forget to close with ')' when finished.`,
      };
    }

    // Check comma vs colon in ranges: if =SUM(B3,E3) while target has colon
    if (upper.startsWith('=SUM') && upper.includes(',') && !upper.includes(':')) {
      const expected = String(slide.solution.expectedFormula || '');
      if (expected.includes(':')) {
        return {
          status: 'syntax_warning',
          message: `Notice: A comma ',' only adds individual cells. To add the continuous range, use a colon ':' (e.g. B3:E3).`,
        };
      }
    }

    return {
      status: 'ready_to_submit',
      message: `Formula looks valid! Press Enter or click Check to verify calculations.`,
    };
  }

  // 5. Data Entry
  return {
    status: 'ready_to_submit',
    message: `Press Enter or click Check to submit "${trimmed}".`,
  };
}
