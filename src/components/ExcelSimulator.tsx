import React, { useState, useEffect, useRef } from 'react';
import { GridData, CellFormat } from '../types/excel';
import { evaluateFormula, parseCellAddress, expandRange } from '../engine/formulaEngine';
import { 
  Bold, Italic, Underline, DollarSign, Percent, 
  RotateCcw, RotateCw, Sigma, 
  Search, Keyboard
} from 'lucide-react';

interface ExcelSimulatorProps {
  initialGrid: GridData;
  targetCell?: string;
  targetRange?: string;
  requiredShortcut?: string;
  onGridChange?: (newGrid: GridData) => void;
  onShortcutTriggered?: (shortcut: string) => void;
  onSelectionChange?: (range: string, activeCell: string) => void;
  highlightCells?: string[];
  defaultActiveCell?: string;
}

const DEFAULT_COLS = ['A', 'B', 'C', 'D', 'E', 'F'];
const DEFAULT_ROWS = [1, 2, 3, 4, 5, 6, 7, 8];

export const ExcelSimulator: React.FC<ExcelSimulatorProps> = ({
  initialGrid,
  targetCell,
  targetRange,
  requiredShortcut,
  onGridChange,
  onShortcutTriggered,
  onSelectionChange,
  highlightCells = [],
  defaultActiveCell = 'A1',
}) => {
  const [grid, setGrid] = useState<GridData>(initialGrid);
  const [activeCell, setActiveCell] = useState<string>(defaultActiveCell);
  const [selectedRange, setSelectedRange] = useState<string[]>(targetRange ? expandRange(targetRange) : [defaultActiveCell]);
  const [isEditing, setIsEditing] = useState<boolean>(false);
  const [editValue, setEditValue] = useState<string>('');
  const [activeRibbonTab, setActiveRibbonTab] = useState<string>('Home');
  const [lastKeyPressed, setLastKeyPressed] = useState<string>('');
  const [shortcutFeedback, setShortcutFeedback] = useState<string | null>(null);

  const gridContainerRef = useRef<HTMLDivElement>(null);
  const formulaInputRef = useRef<HTMLInputElement>(null);
  const cellInputRef = useRef<HTMLInputElement>(null);

  // Sync state whenever initialGrid changes (e.g. slide change)
  useEffect(() => {
    setGrid(initialGrid);
    const startCell = defaultActiveCell || targetCell || 'A1';
    setActiveCell(startCell);
    setSelectedRange(targetRange ? expandRange(targetRange) : [startCell]);
    setIsEditing(false);
    setEditValue(String(initialGrid[startCell]?.formula || initialGrid[startCell]?.value || ''));
  }, [initialGrid, defaultActiveCell, targetCell, targetRange]);

  // Sync cell edit value when active cell changes
  useEffect(() => {
    const cell = grid[activeCell];
    const val = cell?.formula || (cell?.value !== undefined ? String(cell.value) : '');
    setEditValue(val);
    if (onSelectionChange) {
      const rangeStr = selectedRange.length > 1 ? `${selectedRange[0]}:${selectedRange[selectedRange.length - 1]}` : activeCell;
      onSelectionChange(rangeStr, activeCell);
    }
  }, [activeCell, selectedRange, grid, onSelectionChange]);

  // Compute live Status Bar statistics for selected range
  const calculateStatusBarMetrics = () => {
    const nums: number[] = [];
    let filledCount = 0;

    selectedRange.forEach((ref) => {
      const cell = grid[ref];
      if (cell && cell.value !== '' && cell.value !== undefined) {
        filledCount++;
        const num = typeof cell.value === 'number' ? cell.value : parseFloat(String(cell.value).replace(/[$,%]/g, ''));
        if (!isNaN(num)) {
          nums.push(num);
        }
      }
    });

    const sum = nums.reduce((acc, curr) => acc + curr, 0);
    const avg = nums.length > 0 ? Math.round((sum / nums.length) * 100) / 100 : 0;

    return {
      count: filledCount,
      sum: Math.round(sum * 100) / 100,
      avg,
      hasNumbers: nums.length > 0,
    };
  };

  const statusMetrics = calculateStatusBarMetrics();

  // Commit editing changes to grid
  const commitEdit = (valToCommit?: string) => {
    const text = valToCommit !== undefined ? valToCommit : editValue;
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

    const updatedGrid: GridData = {
      ...grid,
      [activeCell]: {
        ...(grid[activeCell] || {}),
        value: evaluatedVal,
        formula: isFormula ? text.trim() : undefined,
      },
    };

    setGrid(updatedGrid);
    setIsEditing(false);
    if (onGridChange) {
      onGridChange(updatedGrid);
    }
  };

  // Keyboard shortcut detection
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      // Do not capture if user is typing in regular text inputs outside simulator
      const targetTag = (e.target as HTMLElement)?.tagName?.toLowerCase();
      const isInput = targetTag === 'input' || targetTag === 'textarea';

      const keys: string[] = [];
      if (e.ctrlKey || e.metaKey) keys.push('Ctrl');
      if (e.altKey) keys.push('Alt');
      if (e.shiftKey) keys.push('Shift');
      
      const keyName = e.key.toUpperCase();
      if (!['CONTROL', 'ALT', 'SHIFT', 'META'].includes(keyName)) {
        keys.push(keyName === ' ' ? 'Space' : keyName);
      }

      const shortcutCombo = keys.join(' + ');
      if (keys.length > 1 || ['F2', 'F5', 'F4', 'ENTER', 'ESCAPE'].includes(keyName)) {
        setLastKeyPressed(shortcutCombo);
      }

      // Check specific shortcut mappings
      // 1. Ctrl + B
      if ((e.ctrlKey || e.metaKey) && keyName === 'B') {
        e.preventDefault();
        toggleFormat('bold');
        triggerShortcutAction('Ctrl + B');
        return;
      }

      // 2. Ctrl + Z
      if ((e.ctrlKey || e.metaKey) && keyName === 'Z') {
        e.preventDefault();
        triggerShortcutAction('Ctrl + Z');
        return;
      }

      // 3. Alt + = (AutoSum)
      if (e.altKey && (e.key === '=' || e.key === '+')) {
        e.preventDefault();
        triggerAutoSum();
        triggerShortcutAction('Alt + =');
        return;
      }

      // 4. Ctrl + A (Select Region)
      if ((e.ctrlKey || e.metaKey) && keyName === 'A' && !isInput) {
        e.preventDefault();
        const allPopulated = Object.keys(grid);
        setSelectedRange(allPopulated.length > 0 ? allPopulated : ['A1', 'B1', 'C1', 'A2', 'B2', 'C2']);
        triggerShortcutAction('Ctrl + A');
        return;
      }

      // 5. Shift + Space (Select Row)
      if (e.shiftKey && e.code === 'Space' && !isInput) {
        e.preventDefault();
        const parsed = parseCellAddress(activeCell);
        if (parsed) {
          const rowCells = DEFAULT_COLS.map((c) => `${c}${parsed.row}`);
          setSelectedRange(rowCells);
        }
        triggerShortcutAction('Shift + Space');
        return;
      }

      // 6. Ctrl + Space (Select Column)
      if ((e.ctrlKey || e.metaKey) && e.code === 'Space' && !isInput) {
        e.preventDefault();
        const parsed = parseCellAddress(activeCell);
        if (parsed) {
          const colCells = DEFAULT_ROWS.map((r) => `${parsed.col}${r}`);
          setSelectedRange(colCells);
        }
        triggerShortcutAction('Ctrl + Space');
        return;
      }

      // 7. F2 (Edit Cell)
      if (e.key === 'F2') {
        e.preventDefault();
        setIsEditing(true);
        triggerShortcutAction('F2');
        setTimeout(() => cellInputRef.current?.focus(), 50);
        return;
      }

      // 8. Escape (Cancel Edit)
      if (e.key === 'Escape' && isEditing) {
        e.preventDefault();
        setIsEditing(false);
        const cell = grid[activeCell];
        setEditValue(cell?.formula || String(cell?.value || ''));
        return;
      }

      // 9. Enter (Commit Edit & Move Down)
      if (e.key === 'Enter') {
        e.preventDefault();
        commitEdit();
        const parsed = parseCellAddress(activeCell);
        if (parsed && parsed.row < DEFAULT_ROWS[DEFAULT_ROWS.length - 1]) {
          const next = `${parsed.col}${parsed.row + 1}`;
          setActiveCell(next);
          setSelectedRange([next]);
        }
        return;
      }

      // 10. Arrow key navigation when not editing text
      if (!isEditing && ['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(e.key)) {
        const parsed = parseCellAddress(activeCell);
        if (!parsed) return;
        e.preventDefault();

        let newColIdx = parsed.colIdx;
        let newRow = parsed.row;

        if (e.key === 'ArrowUp') newRow = Math.max(1, newRow - 1);
        if (e.key === 'ArrowDown') newRow = Math.min(DEFAULT_ROWS.length, newRow + 1);
        if (e.key === 'ArrowLeft') newColIdx = Math.max(1, newColIdx - 1);
        if (e.key === 'ArrowRight') newColIdx = Math.min(DEFAULT_COLS.length, newColIdx + 1);

        const newCol = DEFAULT_COLS[newColIdx - 1] || 'A';
        const newRef = `${newCol}${newRow}`;
        setActiveCell(newRef);

        if (e.shiftKey) {
          setSelectedRange((prev) => [...prev, newRef]);
        } else {
          setSelectedRange([newRef]);
        }
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [activeCell, isEditing, editValue, grid, selectedRange]);

  const triggerShortcutAction = (shortcutName: string) => {
    setShortcutFeedback(`Executed: ${shortcutName}`);
    if (onShortcutTriggered) {
      onShortcutTriggered(shortcutName);
    }
    setTimeout(() => setShortcutFeedback(null), 2500);
  };

  const triggerAutoSum = () => {
    const parsed = parseCellAddress(activeCell);
    if (!parsed) return;
    const col = parsed.col;
    const row = parsed.row;
    
    // Look above for numbers
    if (row > 1) {
      const topRow = 2; // Usually row 1 is header
      const sumFormula = `=SUM(${col}${topRow}:${col}${row - 1})`;
      setEditValue(sumFormula);
      commitEdit(sumFormula);
    }
  };

  const toggleFormat = (formatKey: keyof CellFormat) => {
    const current = grid[activeCell]?.format?.[formatKey];
    const updatedGrid: GridData = {
      ...grid,
      [activeCell]: {
        ...(grid[activeCell] || { value: '' }),
        format: {
          ...(grid[activeCell]?.format || {}),
          [formatKey]: !current,
        },
      },
    };
    setGrid(updatedGrid);
    if (onGridChange) onGridChange(updatedGrid);
  };

  const applyNumberFormat = (fmt: 'currency' | 'percentage' | 'number' | 'general') => {
    const updatedGrid: GridData = {
      ...grid,
      [activeCell]: {
        ...(grid[activeCell] || { value: '' }),
        format: {
          ...(grid[activeCell]?.format || {}),
          format: fmt,
        },
      },
    };
    setGrid(updatedGrid);
    if (onGridChange) onGridChange(updatedGrid);
  };

  const formatDisplayValue = (cell?: { value: any; format?: CellFormat }) => {
    if (!cell || cell.value === undefined || cell.value === '') return '';
    const val = cell.value;
    const fmt = cell.format?.format;

    if (fmt === 'currency' && typeof val === 'number') {
      return `$${val.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
    }
    if (fmt === 'percentage' && typeof val === 'number') {
      return `${Math.round(val * 1000) / 10}%`;
    }
    if (typeof val === 'number') {
      return val.toLocaleString('en-US');
    }
    return String(val);
  };

  return (
    <div 
      ref={gridContainerRef}
      className="flex flex-col border border-slate-300 rounded-lg overflow-hidden bg-white shadow-sm font-sans select-none text-xs text-slate-800"
    >
      {/* 1. Excel Title Bar */}
      <div className="bg-[#107c41] text-white px-3 py-1.5 flex items-center justify-between text-xs font-medium">
        <div className="flex items-center gap-3">
          <div className="flex items-center gap-1.5 font-semibold tracking-wide">
            <span className="bg-white text-[#107c41] px-1 rounded font-bold text-[10px]">XLS</span>
            <span>Book1.xlsx - Microsoft Excel</span>
          </div>
          <span className="text-emerald-200 text-[11px] hidden sm:inline">AutoSave: On</span>
        </div>

        <div className="flex items-center gap-2">
          {requiredShortcut && (
            <div className="bg-emerald-900/60 border border-emerald-400/40 text-emerald-100 px-2 py-0.5 rounded text-[11px] flex items-center gap-1.5">
              <Keyboard className="w-3 h-3 text-emerald-300" />
              <span>Target Shortcut:</span>
              <kbd className="bg-emerald-800 px-1.5 py-0.2 rounded font-mono font-semibold text-white">
                {requiredShortcut}
              </kbd>
            </div>
          )}
          <div className="relative hidden md:block">
            <Search className="w-3 h-3 absolute left-2 top-1.5 text-emerald-200" />
            <input 
              type="text" 
              placeholder="Search or Alt+Q" 
              readOnly
              className="bg-emerald-800/60 placeholder-emerald-300 text-white rounded text-[11px] pl-6 pr-2 py-0.5 w-32 focus:outline-none"
            />
          </div>
        </div>
      </div>

      {/* 2. Ribbon Tabs */}
      <div className="bg-[#f3f4f6] border-b border-slate-200 px-2 pt-1 flex items-center gap-1 text-[11px]">
        {['File', 'Home', 'Insert', 'Page Layout', 'Formulas', 'Data', 'Review', 'View'].map((tab) => (
          <button
            key={tab}
            onClick={() => setActiveRibbonTab(tab)}
            className={`px-3 py-1 font-medium transition-colors rounded-t ${
              activeRibbonTab === tab
                ? 'bg-white text-[#107c41] font-semibold border-t-2 border-[#107c41]'
                : 'text-slate-600 hover:text-slate-900 hover:bg-slate-200/60'
            }`}
          >
            {tab}
          </button>
        ))}
      </div>

      {/* 3. Ribbon Command Toolbar (Home Tab) */}
      <div className="bg-white border-b border-slate-200 px-2 py-1.5 flex items-center gap-2 overflow-x-auto text-slate-700">
        {/* Undo/Redo */}
        <div className="flex items-center gap-0.5 pr-2 border-r border-slate-200">
          <button 
            onClick={() => triggerShortcutAction('Ctrl + Z')}
            title="Undo (Ctrl + Z)"
            className="p-1 hover:bg-slate-100 rounded text-slate-600 hover:text-slate-900"
          >
            <RotateCcw className="w-3.5 h-3.5" />
          </button>
          <button 
            onClick={() => triggerShortcutAction('Ctrl + Y')}
            title="Redo (Ctrl + Y)"
            className="p-1 hover:bg-slate-100 rounded text-slate-600 hover:text-slate-900"
          >
            <RotateCw className="w-3.5 h-3.5" />
          </button>
        </div>

        {/* Font Group: Bold, Italic, Underline */}
        <div className="flex items-center gap-0.5 pr-2 border-r border-slate-200">
          <button
            onClick={() => {
              toggleFormat('bold');
              triggerShortcutAction('Ctrl + B');
            }}
            title="Bold (Ctrl + B)"
            className={`p-1 rounded text-slate-700 hover:bg-slate-100 ${
              grid[activeCell]?.format?.bold ? 'bg-emerald-100 text-[#107c41] font-bold' : ''
            }`}
          >
            <Bold className="w-3.5 h-3.5" />
          </button>
          <button
            onClick={() => toggleFormat('italic')}
            title="Italic (Ctrl + I)"
            className={`p-1 rounded text-slate-700 hover:bg-slate-100 ${
              grid[activeCell]?.format?.italic ? 'bg-emerald-100 text-[#107c41]' : ''
            }`}
          >
            <Italic className="w-3.5 h-3.5" />
          </button>
          <button
            onClick={() => toggleFormat('underline')}
            title="Underline (Ctrl + U)"
            className={`p-1 rounded text-slate-700 hover:bg-slate-100 ${
              grid[activeCell]?.format?.underline ? 'bg-emerald-100 text-[#107c41]' : ''
            }`}
          >
            <Underline className="w-3.5 h-3.5" />
          </button>
        </div>

        {/* Number Group: Currency, Percentage */}
        <div className="flex items-center gap-0.5 pr-2 border-r border-slate-200">
          <button
            onClick={() => applyNumberFormat('currency')}
            title="Currency Format (Ctrl + Shift + $)"
            className="p-1 hover:bg-slate-100 rounded text-slate-700 flex items-center gap-0.5"
          >
            <DollarSign className="w-3.5 h-3.5" />
          </button>
          <button
            onClick={() => applyNumberFormat('percentage')}
            title="Percent Format (Ctrl + Shift + %)"
            className="p-1 hover:bg-slate-100 rounded text-slate-700 flex items-center gap-0.5"
          >
            <Percent className="w-3.5 h-3.5" />
          </button>
        </div>

        {/* AutoSum Tool */}
        <div className="flex items-center gap-1 pr-2 border-r border-slate-200">
          <button
            onClick={() => {
              triggerAutoSum();
              triggerShortcutAction('Alt + =');
            }}
            title="AutoSum (Alt + =)"
            className="px-2 py-0.5 hover:bg-slate-100 rounded text-slate-700 flex items-center gap-1 text-[11px] font-medium"
          >
            <Sigma className="w-3.5 h-3.5 text-[#107c41]" />
            <span>AutoSum</span>
          </button>
        </div>

        {/* Quick Shortcut simulation bar for mobile / intercepted OS keys */}
        {requiredShortcut && (
          <div className="ml-auto flex items-center gap-1.5">
            <span className="text-[11px] text-slate-500 hidden lg:inline">Shortcut shortcut helper:</span>
            <button
              onClick={() => {
                if (requiredShortcut === 'Ctrl + B') toggleFormat('bold');
                if (requiredShortcut === 'Alt + =') triggerAutoSum();
                if (requiredShortcut === 'Shift + Space') {
                  const p = parseCellAddress(activeCell);
                  if (p) setSelectedRange(DEFAULT_COLS.map((c) => `${c}${p.row}`));
                }
                if (requiredShortcut === 'Ctrl + Space') {
                  const p = parseCellAddress(activeCell);
                  if (p) setSelectedRange(DEFAULT_ROWS.map((r) => `${p.col}${r}`));
                }
                if (requiredShortcut === 'Ctrl + A') {
                  setSelectedRange(Object.keys(grid));
                }
                triggerShortcutAction(requiredShortcut);
              }}
              className="px-2.5 py-1 bg-emerald-600 hover:bg-emerald-700 text-white rounded text-[11px] font-medium flex items-center gap-1 shadow-xs transition-colors"
            >
              <Keyboard className="w-3 h-3" />
              <span>Press {requiredShortcut}</span>
            </button>
          </div>
        )}
      </div>

      {/* 4. Formula Bar & Name Box */}
      <div className="bg-white border-b border-slate-200 px-2 py-1 flex items-center gap-2 text-xs">
        {/* Name Box */}
        <div className="w-20 px-2 py-0.5 border border-slate-300 rounded font-mono font-medium text-slate-800 text-center bg-slate-50 text-[11px]">
          {selectedRange.length > 1
            ? `${selectedRange[0]}:${selectedRange[selectedRange.length - 1]}`
            : activeCell}
        </div>

        {/* fx symbol */}
        <div className="text-slate-400 font-serif italic font-bold px-1 select-none text-sm">
          fx
        </div>

        {/* Formula Input Box */}
        <div className="flex-1 relative">
          <input
            ref={formulaInputRef}
            type="text"
            value={editValue}
            onChange={(e) => setEditValue(e.target.value)}
            onFocus={() => setIsEditing(true)}
            onBlur={() => commitEdit()}
            onKeyDown={(e) => {
              if (e.key === 'Enter') {
                commitEdit();
              }
            }}
            placeholder="Enter formula (e.g. =SUM(B2:B5)) or cell value"
            className="w-full px-2 py-1 border border-slate-300 focus:border-[#107c41] focus:ring-1 focus:ring-[#107c41] rounded font-mono text-slate-900 text-xs focus:outline-none"
          />
        </div>
      </div>

      {/* 5. Spreadsheet Grid */}
      <div className="overflow-x-auto bg-slate-100 max-h-[380px]">
        <table className="border-collapse table-fixed w-full">
          <thead>
            <tr>
              {/* Select all top-left cell */}
              <th className="w-10 bg-slate-200 border-b border-r border-slate-300 text-slate-500 font-normal text-center py-1">
                ◢
              </th>
              {DEFAULT_COLS.map((col) => {
                const isColActive = parseCellAddress(activeCell)?.col === col;
                return (
                  <th
                    key={col}
                    className={`min-w-[110px] py-1 border-b border-r border-slate-300 text-[11px] font-semibold text-center transition-colors ${
                      isColActive ? 'bg-emerald-100 text-[#107c41]' : 'bg-slate-200 text-slate-700'
                    }`}
                  >
                    {col}
                  </th>
                );
              })}
            </tr>
          </thead>
          <tbody>
            {DEFAULT_ROWS.map((row) => {
              const isRowActive = parseCellAddress(activeCell)?.row === row;
              return (
                <tr key={row}>
                  {/* Row header */}
                  <td
                    className={`w-10 border-b border-r border-slate-300 text-[11px] font-semibold text-center py-1 transition-colors ${
                      isRowActive ? 'bg-emerald-100 text-[#107c41]' : 'bg-slate-200 text-slate-700'
                    }`}
                  >
                    {row}
                  </td>

                  {/* Columns */}
                  {DEFAULT_COLS.map((col) => {
                    const cellRef = `${col}${row}`;
                    const cell = grid[cellRef];
                    const isCellActive = activeCell === cellRef;
                    const isCellSelected = selectedRange.includes(cellRef);
                    const isTarget = targetCell === cellRef || highlightCells.includes(cellRef);

                    return (
                      <td
                        key={cellRef}
                        onClick={() => {
                          if (isEditing && activeCell !== cellRef) {
                            commitEdit();
                          }
                          setActiveCell(cellRef);
                          setSelectedRange([cellRef]);
                        }}
                        onDoubleClick={() => {
                          setActiveCell(cellRef);
                          setIsEditing(true);
                          setTimeout(() => cellInputRef.current?.focus(), 50);
                        }}
                        style={{
                          backgroundColor: cell?.format?.bgColor || (isCellSelected && !isCellActive ? '#e0f2fe' : undefined),
                        }}
                        className={`relative px-2 py-1 border-b border-r border-slate-300 min-w-[110px] h-7 text-xs overflow-hidden text-ellipsis whitespace-nowrap cursor-cell transition-colors ${
                          isCellActive
                            ? 'ring-2 ring-[#107c41] bg-white z-10'
                            : isTarget
                            ? 'bg-amber-50/70'
                            : 'bg-white'
                        }`}
                      >
                        {isCellActive && isEditing ? (
                          <input
                            ref={cellInputRef}
                            type="text"
                            value={editValue}
                            onChange={(e) => setEditValue(e.target.value)}
                            onBlur={() => commitEdit()}
                            onKeyDown={(e) => {
                              if (e.key === 'Enter') {
                                commitEdit();
                              }
                            }}
                            autoFocus
                            className="w-full h-full p-0 m-0 border-0 outline-none font-mono text-xs bg-transparent text-slate-900"
                          />
                        ) : (
                          <div
                            className={`w-full h-full flex items-center font-mono ${
                              cell?.format?.bold ? 'font-bold' : ''
                            } ${cell?.format?.italic ? 'italic' : ''} ${
                              cell?.format?.underline ? 'underline' : ''
                            } ${
                              cell?.format?.align === 'center'
                                ? 'justify-center text-center'
                                : cell?.format?.align === 'right' || typeof cell?.value === 'number'
                                ? 'justify-end text-right'
                                : 'justify-start text-left'
                            } ${cell?.format?.textColor || 'text-slate-800'}`}
                          >
                            {formatDisplayValue(cell)}
                          </div>
                        )}

                        {/* Excel bottom-right fill handle indicator on active cell */}
                        {isCellActive && !isEditing && (
                          <div className="absolute -bottom-1 -right-1 w-2 h-2 bg-[#107c41] border border-white z-20 cursor-crosshair" />
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

      {/* 6. Excel Sheet Tabs Bar */}
      <div className="bg-[#f3f4f6] border-t border-slate-200 px-3 py-1 flex items-center justify-between text-xs text-slate-600">
        <div className="flex items-center gap-1">
          <div className="bg-white border-t-2 border-[#107c41] text-[#107c41] font-semibold px-3 py-0.5 rounded-t shadow-2xs text-[11px]">
            Sheet1
          </div>
          <button 
            title="New Sheet" 
            className="w-5 h-5 flex items-center justify-center rounded hover:bg-slate-200 text-slate-500 font-bold"
          >
            +
          </button>
        </div>

        {/* Key Press & Shortcut Feedback readout */}
        <div className="flex items-center gap-2">
          {shortcutFeedback ? (
            <span className="text-[#107c41] font-semibold text-[11px] animate-fade-in">
              ✓ {shortcutFeedback}
            </span>
          ) : lastKeyPressed ? (
            <span className="text-slate-400 text-[11px] font-mono">
              Key: {lastKeyPressed}
            </span>
          ) : (
            <span className="text-slate-400 text-[11px]">Ready</span>
          )}
        </div>
      </div>

      {/* 7. Excel Status Bar (Live AutoCalculate) */}
      <div className="bg-[#107c41] text-white px-3 py-1 flex items-center justify-between text-[11px] font-medium tracking-tight">
        <div className="flex items-center gap-3">
          <span>Ready</span>
          <span className="text-emerald-300">|</span>
          <span>Cell: {activeCell}</span>
        </div>

        {/* Live AutoCalculate Readout */}
        <div className="flex items-center gap-4 text-emerald-50">
          {statusMetrics.hasNumbers && (
            <>
              <span>Average: {statusMetrics.avg.toLocaleString()}</span>
              <span>Count: {statusMetrics.count}</span>
              <span className="font-semibold text-white">
                Sum: {statusMetrics.sum.toLocaleString()}
              </span>
            </>
          )}
          {!statusMetrics.hasNumbers && statusMetrics.count > 0 && (
            <span>Count: {statusMetrics.count}</span>
          )}
          <span className="text-emerald-300">100%</span>
        </div>
      </div>
    </div>
  );
};
