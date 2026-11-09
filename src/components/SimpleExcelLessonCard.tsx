import React, { useState, useEffect, useRef } from 'react';
import { SlideData, GridData, CheckResult, CellFormat } from '../types/excel';
import { evaluateFormula, parseCellAddress } from '../engine/formulaEngine';
import { 
  generateDynamicSlideVariation, 
  evaluateLiveInput 
} from '../engine/dynamicQuestionGenerator';
import { 
  Check, Lightbulb, RotateCcw, ArrowRight, ArrowLeft, 
  HelpCircle, Eye, Keyboard, Play, Sparkles, Shuffle
} from 'lucide-react';

interface SimpleExcelLessonCardProps {
  slide: SlideData;
  isCompleted: boolean;
  onCompleteSlide: (slideId: string) => void;
  onNavigatePrev: () => void;
  onNavigateNext: () => void;
  hasPrev: boolean;
  hasNext: boolean;
}

// 7 Columns (A-G) and 8 Rows (1-8) matching the pedagogical reference screenshots
const DEFAULT_COLS = ['A', 'B', 'C', 'D', 'E', 'F', 'G'];
const DEFAULT_ROWS = [1, 2, 3, 4, 5, 6, 7, 8];

export const SimpleExcelLessonCard: React.FC<SimpleExcelLessonCardProps> = ({
  slide,
  isCompleted,
  onCompleteSlide,
  onNavigatePrev,
  onNavigateNext,
  hasPrev,
  hasNext,
}) => {
  // Dynamically generated problem variation (parameters, bounds, layouts, numbers change dynamically)
  const [currentSlide, setCurrentSlide] = useState<SlideData>(() => generateDynamicSlideVariation(slide));
  const [grid, setGrid] = useState<GridData>(() => currentSlide.excel.initialGrid);
  
  const targetCell = currentSlide.practiceQuestion.targetCell || currentSlide.excel.defaultActiveCell || 'B5';
  const [activeCell, setActiveCell] = useState<string>(targetCell);
  const [currentText, setCurrentText] = useState<string>('');
  const [isEditingInCell, setIsEditingInCell] = useState<boolean>(true);
  const [selectedRange, setSelectedRange] = useState<string[]>([targetCell]);
  
  const [checkResult, setCheckResult] = useState<CheckResult | null>(null);
  const [showHint, setShowHint] = useState<boolean>(false);
  const [showReveal, setShowReveal] = useState<boolean>(false);
  const [variationCount, setVariationCount] = useState<number>(1);
  
  // Shortcut tracking
  const [lastKeysPressed, setLastKeysPressed] = useState<string>('');
  const [shortcutFeedback, setShortcutFeedback] = useState<string | null>(null);
  const [selectionAnchor, setSelectionAnchor] = useState<string>('B2');
  const [selectionHead, setSelectionHead] = useState<string>('B2');

  const inCellInputRef = useRef<HTMLInputElement>(null);
  const formulaInputRef = useRef<HTMLInputElement>(null);
  const cardContainerRef = useRef<HTMLDivElement>(null);

  // Sync state on slide change: dynamically vary layout, coordinates, and numbers!
  useEffect(() => {
    const dynamicInstance = generateDynamicSlideVariation(slide);
    setCurrentSlide(dynamicInstance);
    setGrid(dynamicInstance.excel.initialGrid);
    const startCell = dynamicInstance.practiceQuestion.targetCell || dynamicInstance.excel.defaultActiveCell || 'B2';
    setActiveCell(startCell);
    setSelectedRange([startCell]);
    setSelectionAnchor(startCell);
    setSelectionHead(startCell);
    
    const cell = dynamicInstance.excel.initialGrid[startCell];
    const initialText = cell?.formula || (cell?.value !== undefined ? String(cell.value) : '');
    setCurrentText(initialText);
    
    setCheckResult(null);
    setShowHint(false);
    setShowReveal(false);
    setLastKeysPressed('');
    setShortcutFeedback(null);
    setIsEditingInCell(true);
    setVariationCount(1);

    // Auto-focus the cell after a tick
    setTimeout(() => {
      inCellInputRef.current?.focus();
    }, 100);
  }, [slide.id]);

  // Handler to generate a new dynamic variation of the same lesson
  const handleGenerateNewVariation = () => {
    const dynamicInstance = generateDynamicSlideVariation(slide, Date.now());
    setCurrentSlide(dynamicInstance);
    setGrid(dynamicInstance.excel.initialGrid);
    const startCell = dynamicInstance.practiceQuestion.targetCell || dynamicInstance.excel.defaultActiveCell || 'B2';
    setActiveCell(startCell);
    setSelectedRange([startCell]);
    setSelectionAnchor(startCell);
    setSelectionHead(startCell);
    
    const cell = dynamicInstance.excel.initialGrid[startCell];
    const initialText = cell?.formula || (cell?.value !== undefined ? String(cell.value) : '');
    setCurrentText(initialText);
    
    setCheckResult(null);
    setShowHint(false);
    setShowReveal(false);
    setLastKeysPressed('');
    setShortcutFeedback(null);
    setIsEditingInCell(true);
    setVariationCount((v) => v + 1);

    setTimeout(() => {
      inCellInputRef.current?.focus();
    }, 50);
  };

  // When activeCell changes, load its content into currentText and focus the cell input
  const handleSelectCell = (cellRef: string, editMode = true) => {
    // If currently typing an incomplete formula with an open paren e.g. =SUM(, clicking another cell inserts the reference
    if (currentText.trim().startsWith('=') && (currentText.endsWith('(') || currentText.endsWith(','))) {
      const updated = currentText + cellRef;
      setCurrentText(updated);
      updateCellContent(activeCell, updated);
      return;
    }

    // Otherwise, switch active cell
    setActiveCell(cellRef);
    setSelectedRange([cellRef]);
    setSelectionAnchor(cellRef);
    setSelectionHead(cellRef);
    const cell = grid[cellRef];
    const val = cell?.formula || (cell?.value !== undefined ? String(cell.value) : '');
    setCurrentText(val);
    setIsEditingInCell(editMode);

    setTimeout(() => {
      inCellInputRef.current?.focus();
    }, 50);
  };

  // Update cell and grid dynamically as user types
  const updateCellContent = (cellKey: string, text: string) => {
    setCurrentText(text);
    const isFormula = text.trim().startsWith('=');
    let evaluatedVal: any = text;

    if (isFormula) {
      const evalRes = evaluateFormula(text.trim(), grid);
      evaluatedVal = evalRes.result;
    } else {
      const parsedNum = parseFloat(text.trim());
      if (!isNaN(parsedNum) && text.trim() === String(parsedNum)) {
        evaluatedVal = parsedNum;
      }
    }

    setGrid((prev) => ({
      ...prev,
      [cellKey]: {
        ...(prev[cellKey] || {}),
        value: evaluatedVal,
        formula: isFormula ? text.trim() : undefined,
      },
    }));
  };

  // Validate selection range dynamically against target objective
  const validateSelectionRange = (_cells: string[], actualRangeStr: string, shortcutName: string) => {
    setLastKeysPressed(shortcutName);

    const targetRange = (currentSlide.practiceQuestion.targetRange || currentSlide.solution.expectedRange || '').toUpperCase().trim();
    const actualUpper = actualRangeStr.toUpperCase().trim();

    if (targetRange && actualUpper === targetRange) {
      setShortcutFeedback(`✓ Perfect: ${actualRangeStr} reached!`);
      setCheckResult({
        correct: true,
        message: `Great job! You highlighted ${actualRangeStr} using ${shortcutName}.`,
        detail: currentSlide.solution.explanationText,
      });
      onCompleteSlide(currentSlide.id);
    } else {
      const targetParts = targetRange.split(':');
      const tStart = parseCellAddress(targetParts[0] || '');
      const tEnd = parseCellAddress(targetParts[1] || targetParts[0] || '');
      const actualParts = actualUpper.split(':');
      const aStart = parseCellAddress(actualParts[0] || '');
      const aEnd = parseCellAddress(actualParts[1] || actualParts[0] || '');

      let progressHint = `Press ${shortcutName} again to extend to ${targetRange}.`;

      if (tStart && tEnd && aStart && aEnd) {
        if (aEnd.colIdx > tEnd.colIdx || aEnd.row > tEnd.row) {
          const shrinkKey = aEnd.colIdx > tEnd.colIdx ? 'Shift + Left Arrow' : 'Shift + Up Arrow';
          progressHint = `Extended past ${targetRange}. Press ${shrinkKey} to shrink back.`;
        } else if (aEnd.colIdx < tEnd.colIdx || aEnd.row < tEnd.row) {
          const remainingSteps = Math.max(tEnd.colIdx - aEnd.colIdx, tEnd.row - aEnd.row);
          progressHint = `Currently at ${actualRangeStr} (${remainingSteps} more press needed to reach ${targetRange}).`;
        }
      }

      setShortcutFeedback(`Detected: ${shortcutName} (${actualRangeStr}). ${progressHint}`);
      setCheckResult({
        correct: false,
        message: `Progress: currently highlighted ${actualRangeStr}. Target is ${targetRange}. ${progressHint}`,
      });
    }
  };

  // Validate edge-of-data jump target
  const validateJumpTarget = (landedCell: string, shortcutName: string) => {
    setLastKeysPressed(shortcutName);

    const expected = (currentSlide.practiceQuestion.targetCell || currentSlide.solution.expectedRange || '').toUpperCase().trim();
    const landedUpper = landedCell.toUpperCase().trim();

    if (expected && landedUpper === expected) {
      setShortcutFeedback(`✓ Jumped to boundary cell ${landedCell}!`);
      setCheckResult({
        correct: true,
        message: `Great job! You executed ${shortcutName} and jumped to boundary cell ${landedCell}.`,
        detail: currentSlide.solution.explanationText,
      });
      onCompleteSlide(currentSlide.id);
    } else {
      setShortcutFeedback(`Detected: ${shortcutName}. Landed on ${landedCell}, target is ${expected}.`);
      setCheckResult({
        correct: false,
        message: `Landed on cell ${landedCell}. The target edge cell is ${expected}. Start from the continuous data and press ${shortcutName}.`,
      });
    }
  };

  // Execute Edge of Data Jump (Ctrl + Arrow)
  const executeEdgeJump = (direction: 'up' | 'down' | 'left' | 'right') => {
    const parsed = parseCellAddress(activeCell);
    if (!parsed) return;
    const { col, row, colIdx } = parsed;

    if (direction === 'down') {
      let targetRow = row;
      const currentHasValue = grid[`${col}${row}`]?.value !== undefined && grid[`${col}${row}`]?.value !== '';
      if (currentHasValue) {
        for (let r = row + 1; r <= DEFAULT_ROWS.length; r++) {
          const val = grid[`${col}${r}`]?.value;
          if (val !== undefined && val !== '') {
            targetRow = r;
          } else {
            break;
          }
        }
      } else {
        for (let r = row + 1; r <= DEFAULT_ROWS.length; r++) {
          const val = grid[`${col}${r}`]?.value;
          if (val !== undefined && val !== '') {
            targetRow = r;
            break;
          }
        }
      }
      const newRef = `${col}${targetRow}`;
      handleSelectCell(newRef, true);
      validateJumpTarget(newRef, 'Ctrl + Down Arrow');
    } else if (direction === 'up') {
      let targetRow = 1;
      const currentHasValue = grid[`${col}${row}`]?.value !== undefined && grid[`${col}${row}`]?.value !== '';
      if (currentHasValue) {
        for (let r = row - 1; r >= 1; r--) {
          const val = grid[`${col}${r}`]?.value;
          if (val !== undefined && val !== '') {
            targetRow = r;
          } else {
            break;
          }
        }
      }
      const newRef = `${col}${targetRow}`;
      handleSelectCell(newRef, true);
      validateJumpTarget(newRef, 'Ctrl + Up Arrow');
    } else if (direction === 'right') {
      let targetColIdx = DEFAULT_COLS.length;
      for (let c = colIdx + 1; c <= DEFAULT_COLS.length; c++) {
        const cName = DEFAULT_COLS[c - 1];
        const val = grid[`${cName}${row}`]?.value;
        if (val === undefined || val === '') {
          targetColIdx = c - 1;
          break;
        }
      }
      const newRef = `${DEFAULT_COLS[targetColIdx - 1] || 'A'}${row}`;
      handleSelectCell(newRef, true);
      validateJumpTarget(newRef, 'Ctrl + Right Arrow');
    } else if (direction === 'left') {
      const newRef = `A${row}`;
      handleSelectCell(newRef, true);
      validateJumpTarget(newRef, 'Ctrl + Left Arrow');
    }
  };

  // Execute Dynamic Range Selection (Shift + Arrow) cell by cell with multi-step validation
  const executeRangeExpand = (direction: 'right' | 'left' | 'down' | 'up') => {
    const anchorParsed = parseCellAddress(selectionAnchor) || parseCellAddress(activeCell);
    const headParsed = parseCellAddress(selectionHead) || parseCellAddress(activeCell);
    if (!anchorParsed || !headParsed) return;

    let newHeadColIdx = headParsed.colIdx;
    let newHeadRow = headParsed.row;

    if (direction === 'right') {
      newHeadColIdx = Math.min(DEFAULT_COLS.length, headParsed.colIdx + 1);
    } else if (direction === 'left') {
      newHeadColIdx = Math.max(1, headParsed.colIdx - 1);
    } else if (direction === 'down') {
      newHeadRow = Math.min(DEFAULT_ROWS.length, headParsed.row + 1);
    } else if (direction === 'up') {
      newHeadRow = Math.max(1, headParsed.row - 1);
    }

    const newHeadCell = `${DEFAULT_COLS[newHeadColIdx - 1]}${newHeadRow}`;
    setSelectionHead(newHeadCell);

    const minCol = Math.min(anchorParsed.colIdx, newHeadColIdx);
    const maxCol = Math.max(anchorParsed.colIdx, newHeadColIdx);
    const minRow = Math.min(anchorParsed.row, newHeadRow);
    const maxRow = Math.max(anchorParsed.row, newHeadRow);

    const newRangeCells: string[] = [];
    for (let r = minRow; r <= maxRow; r++) {
      for (let c = minCol; c <= maxCol; c++) {
        newRangeCells.push(`${DEFAULT_COLS[c - 1]}${r}`);
      }
    }
    setSelectedRange(newRangeCells);

    const actualRangeStr =
      minCol === maxCol && minRow === maxRow
        ? `${DEFAULT_COLS[minCol - 1]}${minRow}`
        : `${DEFAULT_COLS[minCol - 1]}${minRow}:${DEFAULT_COLS[maxCol - 1]}${maxRow}`;

    const shortcutLabel =
      direction === 'right'
        ? 'Shift + Right Arrow'
        : direction === 'left'
        ? 'Shift + Left Arrow'
        : direction === 'down'
        ? 'Shift + Down Arrow'
        : 'Shift + Up Arrow';

    validateSelectionRange(newRangeCells, actualRangeStr, shortcutLabel);
  };

  // Execute Fill Down (Ctrl + D)
  const executeFillDown = () => {
    const target = (currentSlide.practiceQuestion.targetCell || '').toUpperCase().trim();
    const isTarget = activeCell.toUpperCase().trim() === target;

    const parsed = parseCellAddress(activeCell);
    if (!parsed || parsed.row <= 1) return;
    const aboveRef = `${parsed.col}${parsed.row - 1}`;
    const aboveCell = grid[aboveRef];
    if (!aboveCell) return;

    let newFormula = aboveCell.formula;
    let newVal = aboveCell.value;

    if (newFormula) {
      newFormula = newFormula.replace(/([A-Z])(\d+)/g, (_match, colLetter, rowNum) => {
        return `${colLetter}${parseInt(rowNum, 10) + 1}`;
      });
      const evalRes = evaluateFormula(newFormula, grid);
      newVal = evalRes.result;
    }

    setGrid((prev) => ({
      ...prev,
      [activeCell]: {
        ...(prev[activeCell] || {}),
        value: newVal,
        formula: newFormula,
      },
    }));
    setCurrentText(newFormula || String(newVal));
    setLastKeysPressed('Ctrl + D');

    if (isTarget) {
      setShortcutFeedback(`✓ Filled down into ${activeCell}!`);
      setCheckResult({
        correct: true,
        message: `Great job! You used Ctrl + D to copy the formula into ${activeCell}.`,
        detail: currentSlide.solution.explanationText,
      });
      onCompleteSlide(currentSlide.id);
    } else {
      setShortcutFeedback(`Filled down into ${activeCell}. Target cell is ${target}.`);
      setCheckResult({
        correct: false,
        message: `You filled down into ${activeCell}. The target cell is ${target}. Please select ${target} directly below ${parsed.col}${parsed.row - 1}.`,
      });
    }
  };

  // Execute Row Deletion (Ctrl + -)
  const executeDeleteRow = () => {
    const parsed = parseCellAddress(activeCell);
    const rowToDelete = parsed ? parsed.row : 3;
    const targetCellRef = currentSlide.practiceQuestion.targetCell || 'A3';
    const targetParsed = parseCellAddress(targetCellRef);
    const targetRow = targetParsed ? targetParsed.row : 3;

    setGrid((prev) => {
      const nextGrid: GridData = {};
      Object.entries(prev).forEach(([key, cellData]) => {
        const cp = parseCellAddress(key);
        if (!cp) return;
        if (cp.row < rowToDelete) {
          nextGrid[key] = cellData;
        } else if (cp.row > rowToDelete) {
          const shiftedKey = `${cp.col}${cp.row - 1}`;
          nextGrid[shiftedKey] = cellData;
        }
      });
      return nextGrid;
    });

    setLastKeysPressed('Ctrl + -');

    if (rowToDelete === targetRow) {
      setShortcutFeedback(`✓ Deleted duplicate Row ${rowToDelete}!`);
      setCheckResult({
        correct: true,
        message: `Great job! You deleted the duplicate row ${rowToDelete} using Ctrl + -.`,
        detail: currentSlide.solution.explanationText,
      });
      onCompleteSlide(currentSlide.id);
    } else {
      setShortcutFeedback(`Deleted Row ${rowToDelete}. Target row was Row ${targetRow}.`);
      setCheckResult({
        correct: false,
        message: `You deleted Row ${rowToDelete}. The target duplicate row to delete was Row ${targetRow}. Click Restart to restore data and try again.`,
      });
    }
  };

  // Execute Percentage Format (Ctrl + Shift + %)
  const executePercentFormat = () => {
    const target = (currentSlide.practiceQuestion.targetCell || '').toUpperCase().trim();
    const isTarget = activeCell.toUpperCase().trim() === target;

    const cell = grid[activeCell];
    const rawNum = typeof cell?.value === 'number' ? cell.value : parseFloat(String(cell?.value || '0.285'));
    const percentStr = `${Math.round(rawNum * 100)}%`;

    setGrid((prev) => ({
      ...prev,
      [activeCell]: {
        ...(prev[activeCell] || {}),
        value: percentStr,
        format: {
          ...(prev[activeCell]?.format || {}),
          format: 'percentage',
        },
      },
    }));
    setCurrentText(percentStr);
    setLastKeysPressed('Ctrl + Shift + %');

    if (isTarget) {
      setShortcutFeedback(`✓ Formatted ${activeCell} as percentage!`);
      setCheckResult({
        correct: true,
        message: `Great job! You applied percentage format to ${activeCell} using Ctrl + Shift + %.`,
        detail: currentSlide.solution.explanationText,
      });
      onCompleteSlide(currentSlide.id);
    } else {
      setShortcutFeedback(`Formatted ${activeCell}. Target is ${target}.`);
      setCheckResult({
        correct: false,
        message: `Formatted ${activeCell}. The target cell is ${target}. Select ${target} and press Ctrl + Shift + %.`,
      });
    }
  };

  // Execute AutoSum (Alt + =)
  const executeAutoSum = () => {
    const target = (currentSlide.practiceQuestion.targetCell || '').toUpperCase().trim();
    const isTarget = !target || activeCell.toUpperCase().trim() === target;

    const parsed = parseCellAddress(activeCell);
    if (!parsed) return;
    const col = parsed.col;
    const row = parsed.row;
    if (row > 1) {
      const formula = `=SUM(${col}2:${col}${row - 1})`;
      updateCellContent(activeCell, formula);
      setLastKeysPressed('Alt + =');

      if (isTarget) {
        setShortcutFeedback(`✓ Inserted ${formula}!`);
        setCheckResult({
          correct: true,
          message: `Great job! You used Alt + = to insert ${formula}.`,
          detail: currentSlide.solution.explanationText,
        });
        onCompleteSlide(currentSlide.id);
      } else {
        setShortcutFeedback(`Inserted in ${activeCell}. Target is ${target}.`);
        setCheckResult({
          correct: false,
          message: `Inserted in ${activeCell}. Please select cell ${target} and press Alt + =.`,
        });
      }
    }
  };

  // Toggle Bold (Ctrl + B)
  const toggleBold = () => {
    const target = (currentSlide.practiceQuestion.targetCell || '').toUpperCase().trim();
    const isTarget = !target || activeCell.toUpperCase().trim() === target;

    const current = grid[activeCell]?.format?.bold;
    const newBold = !current;
    setGrid((prev) => ({
      ...prev,
      [activeCell]: {
        ...(prev[activeCell] || { value: '' }),
        format: {
          ...(prev[activeCell]?.format || {}),
          bold: newBold,
        },
      },
    }));
    setLastKeysPressed('Ctrl + B');

    if (newBold && isTarget) {
      setShortcutFeedback(`✓ Bold applied to ${activeCell}!`);
      setCheckResult({
        correct: true,
        message: `Great job! You applied bold formatting to ${activeCell} using Ctrl + B.`,
        detail: currentSlide.solution.explanationText,
      });
      onCompleteSlide(currentSlide.id);
    } else if (!isTarget) {
      setShortcutFeedback(`Bold applied to ${activeCell}. Target is ${target}.`);
      setCheckResult({
        correct: false,
        message: `You bolded ${activeCell}. The target cell to bold is ${target}. Select ${target} and press Ctrl + B.`,
      });
    }
  };

  // Undo (Ctrl + Z)
  const executeUndo = () => {
    setLastKeysPressed('Ctrl + Z');
    if (currentSlide.id === 'm1-quick-access') {
      const tCell = currentSlide.practiceQuestion.targetCell || 'B3';
      const expVal = currentSlide.solution.expectedValue ?? 85;
      const finalVal = Array.isArray(expVal) ? expVal[0] : expVal;
      setGrid((prev) => ({
        ...prev,
        [tCell]: { ...(prev[tCell] || {}), value: typeof finalVal === 'boolean' ? String(finalVal) : finalVal },
      }));
      if (activeCell === tCell) setCurrentText(String(finalVal));
      setShortcutFeedback(`✓ Reverted ${tCell} to ${finalVal}!`);
      setCheckResult({
        correct: true,
        message: `Great job! You used Ctrl + Z to undo and revert cell ${tCell}.`,
        detail: currentSlide.solution.explanationText,
      });
      onCompleteSlide(currentSlide.id);
    } else {
      handleRestart();
      setShortcutFeedback('✓ Action undone (Ctrl + Z)');
    }
  };

  // Cycle Dollar Anchors ($) (F4)
  const cycleDollarAnchors = () => {
    if (currentText.includes(':')) {
      const parts = currentText.split(':');
      const toggled = parts.map((p) => (p.includes('$') ? p.replace(/\$/g, '') : `$${p[0]}$${p.slice(1)}`)).join(':');
      updateCellContent(activeCell, toggled);
    } else if (currentText.startsWith('=')) {
      const formulaBody = currentText.slice(1);
      const toggled = formulaBody.includes('$') ? formulaBody.replace(/\$/g, '') : `$${formulaBody[0]}$${formulaBody.slice(1)}`;
      updateCellContent(activeCell, `=${toggled}`);
    }
  };

  // Select Entire Row (Shift + Space)
  const executeSelectRow = () => {
    const parsed = parseCellAddress(activeCell);
    if (!parsed) return;
    const rowCells = DEFAULT_COLS.map((c) => `${c}${parsed.row}`);
    setSelectedRange(rowCells);
    setLastKeysPressed('Shift + Space');

    const targetRow = currentSlide.practiceQuestion.targetCell
      ? parseCellAddress(currentSlide.practiceQuestion.targetCell)?.row
      : undefined;

    if (!targetRow || parsed.row === targetRow) {
      setShortcutFeedback(`✓ Row ${parsed.row} selected!`);
      setCheckResult({
        correct: true,
        message: `Great job! You selected Row ${parsed.row} using Shift + Space.`,
        detail: currentSlide.solution.explanationText,
      });
      onCompleteSlide(currentSlide.id);
    } else {
      setShortcutFeedback(`Row ${parsed.row} selected. Target is Row ${targetRow}.`);
      setCheckResult({
        correct: false,
        message: `You selected Row ${parsed.row}. Please click a cell in Row ${targetRow} and press Shift + Space.`,
      });
    }
  };

  // Select Entire Column (Ctrl + Space)
  const executeSelectColumn = () => {
    const parsed = parseCellAddress(activeCell);
    if (!parsed) return;
    const colCells = DEFAULT_ROWS.map((r) => `${parsed.col}${r}`);
    setSelectedRange(colCells);
    setLastKeysPressed('Ctrl + Space');

    const targetCol = currentSlide.practiceQuestion.targetCell
      ? parseCellAddress(currentSlide.practiceQuestion.targetCell)?.col
      : undefined;

    if (!targetCol || parsed.col === targetCol) {
      setShortcutFeedback(`✓ Column ${parsed.col} selected!`);
      setCheckResult({
        correct: true,
        message: `Great job! You selected Column ${parsed.col} using Ctrl + Space.`,
        detail: currentSlide.solution.explanationText,
      });
      onCompleteSlide(currentSlide.id);
    } else {
      setShortcutFeedback(`Column ${parsed.col} selected. Target is Column ${targetCol}.`);
      setCheckResult({
        correct: false,
        message: `You selected Column ${parsed.col}. Please click a cell in Column ${targetCol} and press Ctrl + Space.`,
      });
    }
  };

  // Select Entire Table Region (Ctrl + A)
  const executeSelectAll = () => {
    const allPopulated = Object.keys(grid).filter((k) => grid[k]?.value !== undefined && grid[k]?.value !== '');
    const selectedCells = allPopulated.length > 0 ? allPopulated : ['A1', 'B1', 'C1', 'A2', 'B2', 'C2'];
    setSelectedRange(selectedCells);
    setLastKeysPressed('Ctrl + A');

    setShortcutFeedback(`✓ Entire table region selected!`);
    setCheckResult({
      correct: true,
      message: `Great job! You selected the entire data region using Ctrl + A.`,
      detail: currentSlide.solution.explanationText,
    });
    onCompleteSlide(currentSlide.id);
  };

  // Unified shortcut runner for on-screen buttons
  const runShortcutAction = (shortcutName: string) => {
    const s = shortcutName.toLowerCase().replace(/\s+/g, '');
    if (s.includes('shift+rightarrow') || s.includes('shift+arrowright')) {
      executeRangeExpand('right');
    } else if (s.includes('shift+leftarrow') || s.includes('shift+arrowleft')) {
      executeRangeExpand('left');
    } else if (s.includes('shift+downarrow') || s.includes('shift+arrowdown')) {
      executeRangeExpand('down');
    } else if (s.includes('shift+uparrow') || s.includes('shift+arrowup')) {
      executeRangeExpand('up');
    } else if (s.includes('ctrl+downarrow') || s.includes('ctrl+arrowdown')) {
      executeEdgeJump('down');
    } else if (s.includes('ctrl+uparrow') || s.includes('ctrl+arrowup')) {
      executeEdgeJump('up');
    } else if (s.includes('ctrl+rightarrow') || s.includes('ctrl+arrowright')) {
      executeEdgeJump('right');
    } else if (s.includes('ctrl+leftarrow') || s.includes('ctrl+arrowleft')) {
      executeEdgeJump('left');
    } else if (s.includes('ctrl+d')) {
      executeFillDown();
    } else if (s.includes('ctrl+-')) {
      executeDeleteRow();
    } else if (s.includes('ctrl+shift+%') || s.includes('ctrl+shift+5')) {
      executePercentFormat();
    } else if (s.includes('ctrl+b')) {
      toggleBold();
    } else if (s.includes('ctrl+z')) {
      executeUndo();
    } else if (s.includes('alt+=')) {
      executeAutoSum();
    } else if (s.includes('shift+space')) {
      executeSelectRow();
    } else if (s.includes('ctrl+space')) {
      executeSelectColumn();
    } else if (s.includes('ctrl+a') || s.includes('ctrl+shift+down')) {
      executeSelectAll();
    } else if (s.includes('f2')) {
      setIsEditingInCell(true);
      inCellInputRef.current?.focus();
      setLastKeysPressed('F2');
      setShortcutFeedback('✓ Edit Mode active (F2)');
      setCheckResult({
        correct: true,
        message: 'Great job! You entered in-cell edit mode using F2.',
        detail: currentSlide.solution.explanationText,
      });
      onCompleteSlide(currentSlide.id);
    } else if (s.includes('f4')) {
      cycleDollarAnchors();
      setLastKeysPressed('F4');
      setShortcutFeedback('✓ Anchored references with $ (F4)');
      setCheckResult({
        correct: true,
        message: 'Great job! You toggled absolute cell references using F4.',
        detail: currentSlide.solution.explanationText,
      });
      onCompleteSlide(currentSlide.id);
    }
  };

  // Keyboard shortcut & navigation handler
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      const isAlt = e.altKey;
      const isCtrl = e.ctrlKey || e.metaKey;
      const isShift = e.shiftKey;
      const keyUpper = e.key.toUpperCase();
      const code = e.code;

      // Track key combination for display
      const comboParts: string[] = [];
      if (isCtrl) comboParts.push('Ctrl');
      if (isAlt) comboParts.push('Alt');
      if (isShift) comboParts.push('Shift');
      if (!['CONTROL', 'ALT', 'SHIFT', 'META'].includes(keyUpper)) {
        comboParts.push(keyUpper === ' ' ? 'Space' : keyUpper);
      }
      const combo = comboParts.join(' + ');
      if (comboParts.length > 1 || ['F2', 'F4', 'F5', 'ENTER', 'ESCAPE'].includes(keyUpper)) {
        setLastKeysPressed(combo);
      }

      // 1. Edge of Data Jump: Ctrl + ArrowDown
      if (isCtrl && (e.key === 'ArrowDown' || code === 'ArrowDown')) {
        e.preventDefault();
        e.stopPropagation();
        executeEdgeJump('down');
        return;
      }

      // 2. Edge of Data Jump: Ctrl + ArrowUp
      if (isCtrl && (e.key === 'ArrowUp' || code === 'ArrowUp')) {
        e.preventDefault();
        e.stopPropagation();
        executeEdgeJump('up');
        return;
      }

      // 3. Edge of Data Jump: Ctrl + ArrowRight
      if (isCtrl && (e.key === 'ArrowRight' || code === 'ArrowRight')) {
        e.preventDefault();
        e.stopPropagation();
        executeEdgeJump('right');
        return;
      }

      // 4. Edge of Data Jump: Ctrl + ArrowLeft
      if (isCtrl && (e.key === 'ArrowLeft' || code === 'ArrowLeft')) {
        e.preventDefault();
        e.stopPropagation();
        executeEdgeJump('left');
        return;
      }

      // 5. Dynamic Range Selection: Shift + ArrowRight
      if (isShift && !isCtrl && (e.key === 'ArrowRight' || code === 'ArrowRight')) {
        const isEditingText = currentText.length > 0 && (e.target as HTMLElement)?.tagName?.toLowerCase() === 'input';
        if (!isEditingText || currentSlide.practiceQuestion.requiredShortcut?.includes('Shift + Right Arrow') || currentSlide.practiceQuestion.practiceType === 'selection') {
          e.preventDefault();
          e.stopPropagation();
          executeRangeExpand('right');
          return;
        }
      }

      // 5b. Dynamic Range Selection: Shift + ArrowLeft (Shrink/expand left)
      if (isShift && !isCtrl && (e.key === 'ArrowLeft' || code === 'ArrowLeft')) {
        const isEditingText = currentText.length > 0 && (e.target as HTMLElement)?.tagName?.toLowerCase() === 'input';
        if (!isEditingText || currentSlide.practiceQuestion.requiredShortcut?.includes('Shift') || currentSlide.practiceQuestion.practiceType === 'selection') {
          e.preventDefault();
          e.stopPropagation();
          executeRangeExpand('left');
          return;
        }
      }

      // 6. Dynamic Range Selection: Shift + ArrowDown
      if (isShift && !isCtrl && (e.key === 'ArrowDown' || code === 'ArrowDown')) {
        e.preventDefault();
        e.stopPropagation();
        executeRangeExpand('down');
        return;
      }

      // 6b. Dynamic Range Selection: Shift + ArrowUp (Shrink/expand up)
      if (isShift && !isCtrl && (e.key === 'ArrowUp' || code === 'ArrowUp')) {
        e.preventDefault();
        e.stopPropagation();
        executeRangeExpand('up');
        return;
      }

      // 7. Fill Down: Ctrl + D
      if (isCtrl && (keyUpper === 'D' || code === 'KeyD')) {
        e.preventDefault();
        e.stopPropagation();
        executeFillDown();
        return;
      }

      // 8. Delete Row: Ctrl + -
      if (isCtrl && (e.key === '-' || code === 'Minus' || code === 'NumpadSubtract')) {
        e.preventDefault();
        e.stopPropagation();
        executeDeleteRow();
        return;
      }

      // 9. Percentage Format: Ctrl + Shift + % (or 5)
      if (isCtrl && isShift && (e.key === '%' || e.key === '5' || code === 'Digit5')) {
        e.preventDefault();
        e.stopPropagation();
        executePercentFormat();
        return;
      }

      // 10. AutoSum: Alt + = or Alt + Shift + =
      if (isAlt && (e.key === '=' || e.key === '+' || code === 'Equal')) {
        e.preventDefault();
        e.stopPropagation();
        executeAutoSum();
        return;
      }

      // 11. Bold: Ctrl + B
      if (isCtrl && (keyUpper === 'B' || code === 'KeyB')) {
        e.preventDefault();
        e.stopPropagation();
        toggleBold();
        return;
      }

      // 12. Select Entire Row: Shift + Space
      if (isShift && (code === 'Space' || e.key === ' ')) {
        const targetTag = (e.target as HTMLElement)?.tagName?.toLowerCase();
        if (targetTag !== 'input' || currentText === '' || currentSlide.practiceQuestion.requiredShortcut === 'Shift + Space') {
          e.preventDefault();
          e.stopPropagation();
          executeSelectRow();
          return;
        }
      }

      // 13. Select Entire Column: Ctrl + Space
      if (isCtrl && (code === 'Space' || e.key === ' ')) {
        e.preventDefault();
        e.stopPropagation();
        executeSelectColumn();
        return;
      }

      // 14. Select Region: Ctrl + A
      if (isCtrl && (keyUpper === 'A' || code === 'KeyA')) {
        e.preventDefault();
        e.stopPropagation();
        executeSelectAll();
        return;
      }

      // 15. Undo: Ctrl + Z
      if (isCtrl && (keyUpper === 'Z' || code === 'KeyZ')) {
        e.preventDefault();
        e.stopPropagation();
        executeUndo();
        return;
      }

      // 16. F2 (Edit cell mode)
      if (e.key === 'F2') {
        e.preventDefault();
        e.stopPropagation();
        setIsEditingInCell(true);
        setLastKeysPressed('F2');
        setShortcutFeedback('✓ Edit Mode active (F2)');
        setCheckResult({
          correct: true,
          message: 'Great job! You entered edit mode using F2.',
          detail: currentSlide.solution.explanationText,
        });
        onCompleteSlide(currentSlide.id);
        setTimeout(() => inCellInputRef.current?.focus(), 50);
        return;
      }

      // 17. F4 (Anchor $ references)
      if (e.key === 'F4') {
        e.preventDefault();
        e.stopPropagation();
        cycleDollarAnchors();
        setLastKeysPressed('F4');
        setShortcutFeedback('✓ Anchored references with $ (F4)');
        setCheckResult({
          correct: true,
          message: 'Great job! You toggled absolute cell references using F4.',
          detail: currentSlide.solution.explanationText,
        });
        onCompleteSlide(currentSlide.id);
        return;
      }

      // 18. Enter Key (Commit & Move down or Check Answer)
      if (e.key === 'Enter') {
        e.preventDefault();
        updateCellContent(activeCell, currentText);

        if (activeCell === targetCell && currentText.trim() !== '') {
          handleCheck();
        } else {
          const parsed = parseCellAddress(activeCell);
          if (parsed && parsed.row < DEFAULT_ROWS[DEFAULT_ROWS.length - 1]) {
            handleSelectCell(`${parsed.col}${parsed.row + 1}`);
          }
        }
        return;
      }

      // 19. Arrow keys navigation (when not actively editing text cursor)
      const targetTag = (e.target as HTMLElement)?.tagName?.toLowerCase();
      const isInput = targetTag === 'input';
      const isAtEdge = inCellInputRef.current?.selectionStart === inCellInputRef.current?.selectionEnd;

      if (!isInput || (isAtEdge && (e.key === 'ArrowUp' || e.key === 'ArrowDown'))) {
        const parsed = parseCellAddress(activeCell);
        if (!parsed) return;

        let newColIdx = parsed.colIdx;
        let newRow = parsed.row;

        if (e.key === 'ArrowUp') {
          e.preventDefault();
          newRow = Math.max(1, newRow - 1);
        } else if (e.key === 'ArrowDown') {
          e.preventDefault();
          newRow = Math.min(DEFAULT_ROWS.length, newRow + 1);
        } else if (e.key === 'ArrowLeft' && !isInput) {
          e.preventDefault();
          newColIdx = Math.max(1, newColIdx - 1);
        } else if (e.key === 'ArrowRight' && !isInput) {
          e.preventDefault();
          newColIdx = Math.min(DEFAULT_COLS.length, newColIdx + 1);
        } else {
          return;
        }

        const nextRef = `${DEFAULT_COLS[newColIdx - 1] || 'A'}${newRow}`;
        handleSelectCell(nextRef);
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [
    activeCell,
    currentText,
    grid,
    targetCell,
    selectionAnchor,
    selectionHead,
    selectedRange,
    currentSlide,
  ]);

  // Validate answer against dynamic currentSlide
  const handleCheck = () => {
    const { solution, practiceQuestion } = currentSlide;
    const cell = grid[targetCell];

    // 1. Shortcut practice check
    if (practiceQuestion.practiceType === 'shortcut') {
      const required = practiceQuestion.requiredShortcut?.toLowerCase().replace(/\s+/g, '');
      const pressed = lastKeysPressed.toLowerCase().replace(/\s+/g, '');
      const formatMet = solution.expectedFormat
        ? Object.entries(solution.expectedFormat).every(([k, v]) => cell?.format?.[k as keyof CellFormat] === v)
        : false;

      let valueMet = false;
      if (solution.expectedValue !== undefined) {
        const expectedVals = Array.isArray(solution.expectedValue) ? solution.expectedValue : [solution.expectedValue];
        valueMet = expectedVals.some((v) => String(cell?.value).trim().toLowerCase() === String(v).trim().toLowerCase());
      }

      let formulaMet = false;
      if (solution.expectedFormula) {
        const expectedForms = Array.isArray(solution.expectedFormula) ? solution.expectedFormula : [solution.expectedFormula];
        const cellFormula = (cell?.formula || String(cell?.value || '')).trim().toUpperCase().replace(/\s+/g, '');
        formulaMet = expectedForms.some((f) => cellFormula === f.trim().toUpperCase().replace(/\s+/g, ''));
      }

      let rangeMet = false;
      if (solution.expectedRange) {
        const actualRange = selectedRange.length > 1 ? `${selectedRange[0]}:${selectedRange[selectedRange.length - 1]}` : (activeCell || selectedRange[0]);
        rangeMet = actualRange.toUpperCase() === solution.expectedRange.toUpperCase() || activeCell.toUpperCase() === solution.expectedRange.toUpperCase();
      }

      if (pressed.includes(required || 'nomatch') || formatMet || valueMet || formulaMet || rangeMet) {
        setCheckResult({
          correct: true,
          message: `Correct! You successfully executed ${practiceQuestion.requiredShortcut}.`,
          detail: solution.explanationText,
        });
        onCompleteSlide(currentSlide.id);
        return;
      } else {
        setCheckResult({
          correct: false,
          message: `Shortcut not detected yet. Press ${practiceQuestion.requiredShortcut} on your keyboard or click the on-screen shortcut button.`,
        });
        return;
      }
    }

    // 2. Selection practice check
    if (practiceQuestion.practiceType === 'selection') {
      const expected = solution.expectedRange?.toUpperCase();
      const actual = selectedRange.length > 1 ? `${selectedRange[0]}:${selectedRange[selectedRange.length - 1]}` : selectedRange[0];

      if (actual === expected || selectedRange.join(',').includes(expected || '')) {
        setCheckResult({
          correct: true,
          message: `Great job! You highlighted ${solution.expectedRange}.`,
          detail: solution.explanationText,
        });
        onCompleteSlide(currentSlide.id);
        return;
      }
    }

    // 3. Formula & value check
    if (!cell || (cell.value === '' && !cell.formula)) {
      setCheckResult({
        correct: false,
        message: `Cell ${targetCell} is empty. Click cell ${targetCell} (or the practice box) and type your formula.`,
      });
      return;
    }

    let isFormulaMatch = false;
    if (solution.expectedFormula) {
      const entered = (cell.formula || String(cell.value)).trim().toUpperCase().replace(/\s+/g, '');
      const expectedList = Array.isArray(solution.expectedFormula)
        ? solution.expectedFormula
        : [solution.expectedFormula];

      isFormulaMatch = expectedList.some((exp) => {
        return entered === exp.trim().toUpperCase().replace(/\s+/g, '');
      });
    }

    let isValueMatch = false;
    if (solution.expectedValue !== undefined) {
      const enteredVal = cell.value;
      const expectedVals = Array.isArray(solution.expectedValue)
        ? solution.expectedValue
        : [solution.expectedValue];

      isValueMatch = expectedVals.some((exp) => {
        if (typeof exp === 'number' && typeof enteredVal === 'number') {
          return Math.abs(enteredVal - exp) < 0.05;
        }
        return String(enteredVal).trim().toLowerCase() === String(exp).trim().toLowerCase();
      });
    }

    if (isFormulaMatch || isValueMatch) {
      setCheckResult({
        correct: true,
        message: 'Correct! Outstanding work!',
        detail: solution.explanationText,
      });
      onCompleteSlide(currentSlide.id);
    } else {
      const enteredStr = cell.formula || String(cell.value);
      setCheckResult({
        correct: false,
        message: `In ${targetCell}, found "${enteredStr}". That's not quite right.`,
        detail: 'Tip: Check your formula syntax (e.g. =SUM(B2:B7)), coordinates, or click Hint to view guidance.',
      });
    }
  };

  const handleRestart = () => {
    setGrid(currentSlide.excel.initialGrid);
    const startCell = currentSlide.practiceQuestion.targetCell || currentSlide.excel.defaultActiveCell || 'A1';
    setActiveCell(startCell);
    setSelectedRange([startCell]);
    const cell = currentSlide.excel.initialGrid[startCell];
    const val = cell?.formula || (cell?.value !== undefined ? String(cell.value) : '');
    setCurrentText(val);
    setCheckResult(null);
    setShowHint(false);
    setShowReveal(false);
    setTimeout(() => inCellInputRef.current?.focus(), 50);
  };

  const handleRevealAnswer = () => {
    setShowReveal(true);
    if (currentSlide.solution.autofillAction) {
      const { targetCell: actCell, formula, value, format } = currentSlide.solution.autofillAction;
      setGrid((prev) => ({
        ...prev,
        [actCell]: {
          value,
          formula,
          format: format || prev[actCell]?.format,
        },
      }));
      setActiveCell(actCell);
      setCurrentText(formula || String(value));
      setCheckResult({
        correct: true,
        message: 'Answer revealed and applied to worksheet.',
        detail: currentSlide.solution.explanationText,
      });
      onCompleteSlide(currentSlide.id);
    }
  };

  const cols = DEFAULT_COLS;
  const rows = DEFAULT_ROWS;

  // Step-by-step adaptive feedback based on live student input
  const liveFeedback = evaluateLiveInput(currentText, currentSlide, selectedRange, activeCell);

  return (
    <div 
      ref={cardContainerRef}
      className="w-full max-w-3xl mx-auto bg-[#0b1517] text-[#e2e8f0] p-4 sm:p-7 rounded-2xl border border-[#1b3539] shadow-2xl space-y-5"
    >
      {/* 1. Header Navigation & Module Title */}
      <div className="flex items-center justify-between text-xs text-[#6e8a8d] border-b border-[#142629] pb-3">
        <div className="flex items-center gap-2">
          <span className="font-semibold text-emerald-400">{currentSlide.moduleTitle}</span>
          <span aria-hidden="true">·</span>
          <span>Topic {currentSlide.topicIndex}</span>
          {variationCount > 1 && (
            <span className="bg-[#102d29] text-emerald-300 border border-[#184841] px-1.5 py-0.2 rounded text-[10px] font-mono">
              Var #{variationCount}
            </span>
          )}
        </div>

        {/* Previous / Next Lesson controls */}
        <div className="flex items-center gap-2">
          <button
            onClick={onNavigatePrev}
            disabled={!hasPrev}
            className={`px-2.5 py-1 rounded text-xs flex items-center gap-1 transition-colors ${
              hasPrev
                ? 'text-[#8da3a6] hover:text-white hover:bg-[#142629] cursor-pointer'
                : 'text-[#31484b] cursor-not-allowed'
            }`}
          >
            <ArrowLeft className="w-3 h-3" />
            <span>Prev</span>
          </button>
          <button
            onClick={onNavigateNext}
            disabled={!hasNext}
            className={`px-2.5 py-1 rounded text-xs flex items-center gap-1 transition-colors ${
              hasNext
                ? 'bg-[#142b2e] hover:bg-[#1c3c40] text-emerald-300 font-medium cursor-pointer'
                : 'text-[#31484b] cursor-not-allowed'
            }`}
          >
            <span>Next</span>
            <ArrowRight className="w-3 h-3" />
          </button>
        </div>
      </div>

      {/* 2. TITLE (Bold, clean title) */}
      <div className="space-y-2">
        <div className="flex items-center justify-between gap-3 flex-wrap">
          <h1 className="text-2xl sm:text-3xl font-extrabold tracking-tight text-white">
            {currentSlide.title}
          </h1>

          {/* Quick Variation button in header */}
          <button
            onClick={handleGenerateNewVariation}
            title="Generate a new dynamic practice problem with randomized parameters"
            className="px-2.5 py-1 bg-[#122b28] hover:bg-[#193b37] text-emerald-300 border border-[#1c4b45] rounded-md text-xs font-medium flex items-center gap-1.5 transition-colors cursor-pointer shadow-xs active:scale-95"
          >
            <Shuffle className="w-3 h-3 text-emerald-400" />
            <span>New Variation</span>
          </button>
        </div>

        {/* 3. EXPLANATION */}
        <p className="text-[#9cb5b8] text-sm leading-relaxed">
          {currentSlide.explanation.overview}
        </p>

        {/* 4. EXAMPLE ROW */}
        <div className="flex flex-wrap items-center gap-2 text-xs pt-1">
          <span className="text-[#6e8a8d] font-medium">Example</span>
          {currentSlide.example.formula ? (
            <code className="bg-[#122326] border border-[#1e3c40] text-emerald-300 px-2 py-0.5 rounded font-mono font-semibold">
              {currentSlide.example.formula}
            </code>
          ) : (
            <code className="bg-[#122326] border border-[#1e3c40] text-emerald-300 px-2 py-0.5 rounded font-mono font-semibold">
              {currentSlide.solution.expectedFormula
                ? Array.isArray(currentSlide.solution.expectedFormula)
                  ? currentSlide.solution.expectedFormula[0]
                  : currentSlide.solution.expectedFormula
                : currentSlide.title}
            </code>
          )}

          {currentSlide.practiceQuestion.requiredShortcut && (
            <span className="bg-[#10302b] text-emerald-300 border border-[#184942] px-2 py-0.5 rounded font-mono font-semibold text-[11px] flex items-center gap-1">
              <Keyboard className="w-3 h-3" />
              <span>Shortcut: {currentSlide.practiceQuestion.requiredShortcut}</span>
            </span>
          )}
        </div>
      </div>

      {/* 5. INTERACTIVE PRACTICE QUESTION BOX (Clickable to focus target cell or execute shortcut immediately!) */}
      <div 
        onClick={() => {
          handleSelectCell(targetCell, true);
          if (currentSlide.practiceQuestion.practiceType === 'shortcut' && currentSlide.practiceQuestion.requiredShortcut) {
            runShortcutAction(currentSlide.practiceQuestion.requiredShortcut);
          }
        }}
        title="Click to start typing or execute shortcut"
        className="group bg-[#0c2220] hover:bg-[#0f2a27] border border-[#14423b] hover:border-emerald-500 rounded-lg px-4 py-3 text-emerald-100 text-sm font-medium leading-relaxed cursor-pointer transition-all flex items-center justify-between gap-3"
      >
        <div className="flex items-center gap-2">
          <span className="w-2 h-2 rounded-full bg-emerald-400 group-hover:animate-ping" />
          <span>{currentSlide.practiceQuestion.instruction}</span>
        </div>
        <div className="flex items-center gap-2 shrink-0">
          {currentSlide.practiceQuestion.requiredShortcut && (
            <button
              type="button"
              onClick={(e) => {
                e.stopPropagation();
                if (currentSlide.practiceQuestion.requiredShortcut) {
                  runShortcutAction(currentSlide.practiceQuestion.requiredShortcut);
                }
              }}
              className="text-[11px] font-mono bg-emerald-600 hover:bg-emerald-500 text-white font-semibold px-2.5 py-1 rounded shadow-xs flex items-center gap-1 cursor-pointer transition-colors"
            >
              <Play className="w-3 h-3 fill-current" />
              <span>Press {currentSlide.practiceQuestion.requiredShortcut}</span>
            </button>
          )}
          <span className="text-[11px] font-mono bg-[#143d36] text-emerald-300 px-2 py-0.5 rounded border border-[#1c544b]">
            Target: {targetCell} ✍️
          </span>
        </div>
      </div>

      {/* 6. FORMULA BAR (Synced live with active cell) */}
      <div className="flex items-center gap-2 bg-[#0e1d20] border border-[#1b3539] focus-within:border-emerald-500 rounded-lg px-3 py-1.5 transition-colors">
        <span className="font-mono font-bold text-xs text-white bg-[#142a2e] px-2 py-0.5 rounded">
          {activeCell}
        </span>
        <span className="font-serif italic font-bold text-xs text-[#527175] select-none">
          fx
        </span>
        <input
          ref={formulaInputRef}
          type="text"
          value={currentText}
          onChange={(e) => updateCellContent(activeCell, e.target.value)}
          placeholder="Type = and your formula, or edit cell directly..."
          className="flex-1 bg-transparent text-white font-mono text-xs focus:outline-none placeholder-[#456266]"
        />
        {shortcutFeedback ? (
          <span className="text-[11px] text-emerald-400 font-mono font-semibold animate-pulse">
            {shortcutFeedback}
          </span>
        ) : lastKeysPressed ? (
          <span className="text-[11px] text-[#527175] font-mono">
            Key: {lastKeysPressed}
          </span>
        ) : null}
      </div>

      {/* 6.1 ADAPTIVE LIVE FEEDBACK (Step-by-step guidance based on live input) */}
      {liveFeedback.status !== 'empty' && (
        <div className="flex items-center gap-2 px-3 py-2 bg-[#0c2325] border border-[#164448] rounded-lg text-xs text-emerald-300 animate-fade-in shadow-xs">
          <Sparkles className="w-3.5 h-3.5 text-emerald-400 shrink-0" />
          <span className="flex-1 font-medium">{liveFeedback.message}</span>
          {liveFeedback.suggestion && (
            <span className="font-mono text-white bg-[#143e3c] px-2 py-0.5 rounded border border-[#1e5c57] text-[11px] font-semibold">
              {liveFeedback.suggestion}
            </span>
          )}
        </div>
      )}

      {/* 7. FULLY INTERACTIVE EXCEL SPREADSHEET GRID */}
      <div className="overflow-x-auto rounded-lg border border-[#193236] bg-[#0c1719]">
        <table className="w-full border-collapse text-xs font-mono">
          <thead>
            <tr className="border-b border-[#193236]">
              {/* Corner spacer */}
              <th className="w-10 py-2 border-r border-[#193236] bg-[#0c1719] text-[#557376] font-normal text-center select-none">
                ◢
              </th>
              {cols.map((col) => {
                const isColActive = parseCellAddress(activeCell)?.col === col;
                return (
                  <th
                    key={col}
                    className={`py-2 px-3 text-center font-semibold border-r border-[#193236] min-w-[100px] transition-colors select-none ${
                      isColActive ? 'text-emerald-400 bg-[#0e2124]' : 'text-[#8ca4a7]'
                    }`}
                  >
                    {col}
                  </th>
                );
              })}
            </tr>
          </thead>
          <tbody>
            {rows.map((row) => {
              const isRowActive = parseCellAddress(activeCell)?.row === row;
              return (
                <tr key={row} className="border-b border-[#193236]">
                  {/* Row number */}
                  <td 
                    className={`w-10 py-2 text-center font-semibold border-r border-[#193236] select-none transition-colors ${
                      isRowActive ? 'text-emerald-400 bg-[#0e2124]' : 'text-[#557376]'
                    }`}
                  >
                    {row}
                  </td>

                  {/* Cells */}
                  {cols.map((col) => {
                    const cellRef = `${col}${row}`;
                    const cell = grid[cellRef];
                    const isCellActive = activeCell === cellRef;
                    const isTarget = targetCell === cellRef;
                    const isSelected = selectedRange.includes(cellRef);

                    // Cell content representation
                    let displayVal = '';
                    if (cell && cell.value !== undefined && cell.value !== '') {
                      displayVal = String(cell.value);
                    }

                    return (
                      <td
                        key={cellRef}
                        onClick={() => handleSelectCell(cellRef, true)}
                        className={`py-1.5 px-2 border-r border-[#193236] transition-colors relative cursor-cell text-left ${
                          isCellActive
                            ? 'bg-[#373d1f] text-white ring-2 ring-emerald-500 z-10'
                            : isSelected
                            ? 'bg-[#153136] text-white select-none'
                            : isTarget
                            ? 'bg-[#15291b] text-emerald-200 select-none'
                            : 'text-[#cbd5e1] hover:bg-[#112326] select-none'
                        } ${cell?.format?.bold ? 'font-bold' : ''}`}
                      >
                        {isCellActive ? (
                          // IN-CELL EDITING INPUT: User can type directly inside the cell!
                          <input
                            ref={inCellInputRef}
                            type="text"
                            value={currentText}
                            autoFocus
                            onChange={(e) => updateCellContent(cellRef, e.target.value)}
                            onKeyDown={(e) => {
                              if (e.key === 'Enter') {
                                e.preventDefault();
                                updateCellContent(activeCell, currentText);
                                handleCheck();
                              }
                            }}
                            placeholder={isTarget && !currentText ? 'type formula...' : ''}
                            className="w-full h-full p-0 m-0 border-0 outline-none font-mono text-xs bg-transparent text-white placeholder-emerald-400/50"
                          />
                        ) : (
                          <div className="flex items-center justify-between w-full h-full min-h-[18px]">
                            <span style={{ color: cell?.format?.textColor }}>{displayVal}</span>
                            {isTarget && !displayVal && (
                              <span className="text-[10px] text-emerald-400/80 font-sans italic animate-pulse">
                                click to type
                              </span>
                            )}
                          </div>
                        )}
                      </td>
                    );
                  })}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      {/* 8. TIP & KEYBOARD SHORTCUT HELPER STRIP */}
      <div className="flex flex-wrap items-center justify-between gap-2 text-[11px] text-[#6e8a8d] pt-1">
        <p className="leading-relaxed">
          <span className="font-semibold text-[#8da3a6]">Tip: </span>
          Click any cell to type directly inside it. Press <kbd className="bg-[#122326] px-1 rounded border border-[#1b3539] text-white">Enter</kbd> to calculate.
        </p>

        {/* Interactive shortcut simulator button */}
        {currentSlide.practiceQuestion.requiredShortcut && (
          <button
            onClick={() => {
              if (currentSlide.practiceQuestion.requiredShortcut) {
                runShortcutAction(currentSlide.practiceQuestion.requiredShortcut);
              }
            }}
            className="px-2.5 py-1 bg-[#173832] hover:bg-[#1e4a42] text-emerald-300 border border-[#23574e] rounded-md font-medium flex items-center gap-1.5 cursor-pointer transition-colors shadow-xs"
          >
            <Keyboard className="w-3.5 h-3.5" />
            <span>Simulate Key: <strong className="text-white">{currentSlide.practiceQuestion.requiredShortcut}</strong></span>
          </button>
        )}
      </div>

      {/* 9. CHECK RESULT FEEDBACK BANNER */}
      {checkResult && (
        <div
          className={`px-4 py-3 rounded-lg border text-xs space-y-1 transition-all ${
            checkResult.correct
              ? 'bg-[#0f2923] border-[#185245] text-emerald-200'
              : 'bg-[#291316] border-[#541f26] text-rose-200'
          }`}
        >
          <div className="font-bold flex items-center justify-between">
            <span className="flex items-center gap-1.5">
              {checkResult.correct ? '✓' : '⚠️'} {checkResult.message}
            </span>
            {checkResult.correct && hasNext && (
              <button
                onClick={onNavigateNext}
                className="px-3 py-1 bg-emerald-600 hover:bg-emerald-500 text-white rounded font-medium flex items-center gap-1 text-[11px] cursor-pointer"
              >
                <span>Next Lesson</span>
                <ArrowRight className="w-3 h-3" />
              </button>
            )}
          </div>
          {checkResult.detail && (
            <p className="text-[#a3b8b8] text-[11px] leading-relaxed">
              {checkResult.detail}
            </p>
          )}
        </div>
      )}

      {/* 10. ACTION BUTTONS */}
      <div className="flex flex-wrap items-center justify-between gap-3 pt-1">
        <div className="flex flex-wrap items-center gap-2">
          {/* Green Check button */}
          <button
            onClick={handleCheck}
            className="px-5 py-2 bg-[#00c274] hover:bg-[#00d882] text-slate-950 font-bold rounded-lg text-xs flex items-center gap-1.5 shadow-md transition-all active:scale-95 cursor-pointer"
          >
            <Check className="w-4 h-4 stroke-[2.5]" />
            <span>Check</span>
          </button>

          {/* New Variation button */}
          <button
            onClick={handleGenerateNewVariation}
            title="Generate a new dynamic problem variation with randomized parameters and coordinates"
            className="px-3.5 py-2 bg-[#12282b] hover:bg-[#18393d] text-emerald-300 border border-[#1f4a50] rounded-lg text-xs font-semibold flex items-center gap-1.5 transition-all cursor-pointer shadow-xs active:scale-95 hover:border-emerald-500"
          >
            <Shuffle className="w-3.5 h-3.5 text-emerald-400" />
            <span>New Variation {variationCount > 1 ? `(#${variationCount})` : ''}</span>
          </button>

          {/* Hint button */}
          <button
            onClick={() => setShowHint(!showHint)}
            className="px-4 py-2 bg-[#122326] hover:bg-[#183135] text-[#b8ced1] border border-[#1d3c40] rounded-lg text-xs font-medium flex items-center gap-1.5 transition-colors cursor-pointer"
          >
            <Lightbulb className="w-3.5 h-3.5 text-amber-400" />
            <span>{showHint ? 'Hide Hint' : 'Hint'}</span>
          </button>

          {/* Restart button */}
          <button
            onClick={handleRestart}
            className="px-4 py-2 bg-[#122326] hover:bg-[#183135] text-[#b8ced1] border border-[#1d3c40] rounded-lg text-xs font-medium flex items-center gap-1.5 transition-colors cursor-pointer"
          >
            <RotateCcw className="w-3.5 h-3.5" />
            <span>Restart</span>
          </button>

          {/* Reveal Answer button */}
          <button
            onClick={handleRevealAnswer}
            className="px-4 py-2 bg-[#122326] hover:bg-[#183135] text-[#b8ced1] border border-[#1d3c40] rounded-lg text-xs font-medium flex items-center gap-1.5 transition-colors cursor-pointer"
          >
            <Eye className="w-3.5 h-3.5 text-emerald-400" />
            <span>Reveal Answer</span>
          </button>
        </div>

        {/* Status Completed indicator */}
        {isCompleted && (
          <span className="text-xs text-emerald-400 font-semibold flex items-center gap-1">
            ✓ Completed
          </span>
        )}
      </div>

      {/* 11. HINT DRAWER */}
      {showHint && currentSlide.practiceQuestion.hint && (
        <div className="bg-[#122326] border border-amber-500/30 text-amber-200/90 rounded-lg p-3 text-xs leading-relaxed animate-fade-in">
          <span className="font-bold text-amber-400">💡 Hint: </span>
          {currentSlide.practiceQuestion.hint}
        </div>
      )}

      {/* 12. REVEAL ANSWER DRAWER */}
      {showReveal && (
        <div className="bg-[#0f2421] border border-emerald-500/40 text-emerald-100 rounded-lg p-3 text-xs space-y-1 animate-fade-in">
          <div className="font-mono font-bold text-emerald-300 flex items-center gap-2">
            <span>Solution: {currentSlide.solution.revealAnswerText}</span>
          </div>
          <p className="text-[#a0c2bf] text-[11px] leading-relaxed">
            {currentSlide.solution.explanationText}
          </p>
        </div>
      )}

      {/* 13. REPORT LINK FOOTER */}
      <div className="pt-2 border-t border-[#142629] flex items-center justify-between text-xs text-[#527175]">
        <button
          onClick={() => alert('Feedback noted! Every lesson dynamically randomizes coordinates and numbers for endless practice.')}
          className="hover:text-[#8da3a6] flex items-center gap-1.5 transition-colors cursor-pointer"
        >
          <HelpCircle className="w-3.5 h-3.5" />
          <span>Something wrong with this lesson? Report it</span>
        </button>

        <span className="text-[11px] font-mono">Module {currentSlide.moduleId}</span>
      </div>
    </div>
  );
};
