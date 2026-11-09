import React, { useState, useEffect } from 'react';
import { SlideData, GridData, CheckResult } from '../types/excel';
import { ExcelSimulator } from './ExcelSimulator';
import { 
  CheckCircle2, AlertCircle, Eye, EyeOff, 
  ArrowRight, ArrowLeft, Lightbulb, Play, RotateCcw
} from 'lucide-react';

interface SlideViewerProps {
  slide: SlideData;
  isCompleted: boolean;
  onCompleteSlide: (slideId: string) => void;
  onNavigatePrev: () => void;
  onNavigateNext: () => void;
  hasPrev: boolean;
  hasNext: boolean;
}

export const SlideViewer: React.FC<SlideViewerProps> = ({
  slide,
  isCompleted,
  onCompleteSlide,
  onNavigatePrev,
  onNavigateNext,
  hasPrev,
  hasNext,
}) => {
  const [currentGrid, setCurrentGrid] = useState<GridData>(slide.excel.initialGrid);
  const [activeCellRef, setActiveCellRef] = useState<string>(slide.excel.defaultActiveCell || 'A1');
  const [selectedRangeRef, setSelectedRangeRef] = useState<string>('');
  const [lastShortcut, setLastShortcut] = useState<string>('');
  const [checkResult, setCheckResult] = useState<CheckResult | null>(null);
  const [showReveal, setShowReveal] = useState<boolean>(false);
  const [showHint, setShowHint] = useState<boolean>(false);

  // Reset local state on slide change
  useEffect(() => {
    setCurrentGrid(slide.excel.initialGrid);
    setActiveCellRef(slide.excel.defaultActiveCell || 'A1');
    setSelectedRangeRef('');
    setLastShortcut('');
    setCheckResult(null);
    setShowReveal(false);
    setShowHint(false);
  }, [slide.id, slide.excel.initialGrid, slide.excel.defaultActiveCell]);

  // Handle checking user's answer
  const handleCheckAnswer = () => {
    const { solution, practiceQuestion, excel } = slide;
    const targetCell = practiceQuestion.targetCell || excel.defaultActiveCell || 'A1';
    const cell = currentGrid[targetCell];

    // 1. Shortcut practice check
    if (practiceQuestion.practiceType === 'shortcut') {
      const required = practiceQuestion.requiredShortcut?.toLowerCase().replace(/\s+/g, '');
      const pressed = lastShortcut.toLowerCase().replace(/\s+/g, '');
      const isFormatMet = solution.expectedFormat
        ? Object.entries(solution.expectedFormat).every(([k, v]) => cell?.format?.[k as keyof typeof cell.format] === v)
        : false;

      if (pressed === required || isFormatMet) {
        setCheckResult({
          correct: true,
          message: `Excellent! You successfully executed ${practiceQuestion.requiredShortcut}.`,
          detail: solution.explanationText,
        });
        onCompleteSlide(slide.id);
        return;
      } else {
        setCheckResult({
          correct: false,
          message: `Shortcut not detected yet. Please press ${practiceQuestion.requiredShortcut} on your keyboard or click the helper button.`,
        });
        return;
      }
    }

    // 2. Selection practice check
    if (practiceQuestion.practiceType === 'selection') {
      const expected = solution.expectedRange?.toUpperCase();
      const actual = selectedRangeRef.toUpperCase();

      if (actual === expected || actual.includes(expected || '')) {
        setCheckResult({
          correct: true,
          message: `Great job! You highlighted the exact target range ${solution.expectedRange}.`,
          detail: solution.explanationText,
        });
        onCompleteSlide(slide.id);
        return;
      } else {
        setCheckResult({
          correct: false,
          message: `Expected range: ${solution.expectedRange}. Currently selected: ${selectedRangeRef || activeCellRef}.`,
        });
        return;
      }
    }

    // 3. Formula & value check
    if (!cell || (cell.value === '' && !cell.formula)) {
      setCheckResult({
        correct: false,
        message: `Cell ${targetCell} is currently empty. Please enter your formula or answer in cell ${targetCell}.`,
      });
      return;
    }

    let isFormulaMatch = false;
    if (solution.expectedFormula) {
      const enteredFormula = (cell.formula || String(cell.value)).trim().toUpperCase().replace(/\s+/g, '');
      const expectedList = Array.isArray(solution.expectedFormula)
        ? solution.expectedFormula
        : [solution.expectedFormula];

      isFormulaMatch = expectedList.some((exp) => {
        const cleanExp = exp.trim().toUpperCase().replace(/\s+/g, '');
        return enteredFormula === cleanExp;
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
      onCompleteSlide(slide.id);
    } else {
      const enteredStr = cell.formula || String(cell.value);
      setCheckResult({
        correct: false,
        message: `Cell ${targetCell} currently has "${enteredStr}". That does not match the expected result.`,
        detail: 'Tip: Check formula spelling, cell coordinates (e.g. B2:B5), and commas. Click "Reveal Answer" if you need a guide.',
      });
    }
  };

  // Autofill solution into the simulator
  const handleAutofillAnswer = () => {
    if (!slide.solution.autofillAction) return;
    const { targetCell, formula, value, format } = slide.solution.autofillAction;

    const updated: GridData = {
      ...currentGrid,
      [targetCell]: {
        value,
        formula,
        format: format || currentGrid[targetCell]?.format,
      },
    };
    setCurrentGrid(updated);
    setActiveCellRef(targetCell);
    setCheckResult({
      correct: true,
      message: 'Solution applied to worksheet!',
      detail: slide.solution.explanationText,
    });
    onCompleteSlide(slide.id);
  };

  const handleResetSlide = () => {
    setCurrentGrid(slide.excel.initialGrid);
    setCheckResult(null);
    setShowReveal(false);
  };

  return (
    <div className="w-full max-w-5xl mx-auto space-y-6 pb-12">
      {/* =========================================================================
          SECTION 1: TITLE (Strict Slide Structure)
      ========================================================================= */}
      <div className="border-b border-slate-200 pb-4">
        <div className="flex flex-wrap items-center justify-between gap-3 mb-2">
          <div className="flex items-center gap-2 text-xs font-semibold text-slate-500 uppercase tracking-wider">
            <span>{slide.moduleTitle}</span>
            <span aria-hidden="true">·</span>
            <span>Topic {slide.topicIndex}</span>
            {isCompleted && (
              <>
                <span aria-hidden="true">·</span>
                <span className="text-emerald-700 font-semibold flex items-center gap-1">
                  <CheckCircle2 className="w-3.5 h-3.5" /> Completed
                </span>
              </>
            )}
          </div>

          {/* Previous / Next Slide buttons */}
          <div className="flex items-center gap-2">
            <button
              onClick={onNavigatePrev}
              disabled={!hasPrev}
              className={`px-3 py-1.5 text-xs font-medium rounded border flex items-center gap-1 transition-colors ${
                hasPrev
                  ? 'border-slate-300 text-slate-700 hover:bg-slate-100'
                  : 'border-slate-200 text-slate-300 cursor-not-allowed'
              }`}
            >
              <ArrowLeft className="w-3 h-3" />
              <span>Previous</span>
            </button>
            <button
              onClick={onNavigateNext}
              disabled={!hasNext}
              className={`px-3 py-1.5 text-xs font-medium rounded flex items-center gap-1 transition-colors ${
                hasNext
                  ? 'bg-slate-900 text-white hover:bg-slate-800'
                  : 'bg-slate-200 text-slate-400 cursor-not-allowed'
              }`}
            >
              <span>Next</span>
              <ArrowRight className="w-3 h-3" />
            </button>
          </div>
        </div>

        <h1 className="text-2xl sm:text-3xl font-bold tracking-tight text-slate-900">
          {slide.title}
        </h1>
      </div>

      {/* =========================================================================
          SECTION 2: EXPLANATION (Strict Slide Structure)
      ========================================================================= */}
      <div className="bg-white border border-slate-200 rounded-lg p-5 shadow-2xs space-y-4">
        <div className="flex items-center gap-2 text-sm font-bold text-slate-900 uppercase tracking-wide">
          <span className="w-2 h-2 rounded-full bg-[#107c41]"></span>
          <span>1. Explanation</span>
        </div>

        <p className="text-slate-700 text-sm leading-relaxed">
          {slide.explanation.overview}
        </p>

        <div className="space-y-2 pt-1">
          <span className="text-xs font-bold text-slate-800 tracking-wider uppercase">Key Concepts &amp; Rules:</span>
          <ul className="space-y-1.5 text-xs text-slate-600 pl-4 list-disc">
            {slide.explanation.keyPoints.map((pt, i) => (
              <li key={i} className="leading-normal">{pt}</li>
            ))}
          </ul>
        </div>

        {slide.explanation.tips && (
          <div className="bg-emerald-50 border border-emerald-200/80 rounded-md p-3 text-xs text-emerald-900 flex items-start gap-2.5">
            <Lightbulb className="w-4 h-4 text-[#107c41] shrink-0 mt-0.5" />
            <div>
              <span className="font-semibold text-[#107c41]">Pro Tip: </span>
              {slide.explanation.tips}
            </div>
          </div>
        )}
      </div>

      {/* =========================================================================
          SECTION 3: EXAMPLE (Strict Slide Structure)
      ========================================================================= */}
      <div className="bg-white border border-slate-200 rounded-lg p-5 shadow-2xs space-y-3">
        <div className="flex items-center gap-2 text-sm font-bold text-slate-900 uppercase tracking-wide">
          <span className="w-2 h-2 rounded-full bg-blue-600"></span>
          <span>2. Example &amp; Practical Context</span>
        </div>

        <p className="text-slate-700 text-xs sm:text-sm leading-relaxed">
          {slide.example.description}
        </p>

        {slide.example.sampleTable && (
          <div className="overflow-x-auto border border-slate-200 rounded my-2">
            <table className="w-full text-xs text-left border-collapse font-mono">
              <thead className="bg-slate-100 text-slate-700 font-semibold border-b border-slate-200">
                <tr>
                  {slide.example.sampleTable.headers.map((h, i) => (
                    <th key={i} className="px-3 py-1.5 border-r border-slate-200">{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {slide.example.sampleTable.rows.map((row, rIdx) => (
                  <tr key={rIdx} className="border-b border-slate-100">
                    {row.map((cell, cIdx) => (
                      <td key={cIdx} className="px-3 py-1 border-r border-slate-100 text-slate-800">
                        {String(cell)}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}

        <div className="text-xs text-slate-500 font-medium">
          <span className="text-slate-700 font-semibold">Expected Behavior: </span>
          {slide.example.outputDescription}
        </div>
      </div>

      {/* =========================================================================
          SECTION 4: PRACTICE QUESTION (Strict Slide Structure)
      ========================================================================= */}
      <div className="bg-amber-50/50 border border-amber-200 rounded-lg p-5 shadow-2xs space-y-3">
        <div className="flex items-center justify-between">
          <div className="flex items-center gap-2 text-sm font-bold text-amber-900 uppercase tracking-wide">
            <span className="w-2 h-2 rounded-full bg-amber-500"></span>
            <span>3. Practice Question</span>
          </div>

          {slide.practiceQuestion.hint && (
            <button
              onClick={() => setShowHint(!showHint)}
              className="text-xs text-amber-800 hover:text-amber-900 underline font-medium flex items-center gap-1"
            >
              <Lightbulb className="w-3.5 h-3.5 text-amber-600" />
              <span>{showHint ? 'Hide Hint' : 'Need a Hint?'}</span>
            </button>
          )}
        </div>

        <p className="text-slate-900 text-sm font-medium leading-relaxed">
          {slide.practiceQuestion.instruction}
        </p>

        {showHint && slide.practiceQuestion.hint && (
          <div className="bg-white border border-amber-300 rounded p-3 text-xs text-amber-900">
            <span className="font-semibold">Hint: </span>
            {slide.practiceQuestion.hint}
          </div>
        )}
      </div>

      {/* =========================================================================
          SECTION 5: EXCEL (Strict Slide Structure - Interactive Simulator)
      ========================================================================= */}
      <div className="space-y-2">
        <div className="flex items-center justify-between px-1">
          <div className="flex items-center gap-2 text-sm font-bold text-slate-900 uppercase tracking-wide">
            <span className="w-2 h-2 rounded-full bg-[#107c41]"></span>
            <span>4. Excel Interactive Simulator</span>
          </div>

          <button
            onClick={handleResetSlide}
            title="Reset worksheet"
            className="text-xs text-slate-500 hover:text-slate-800 flex items-center gap-1"
          >
            <RotateCcw className="w-3 h-3" />
            <span>Reset Sheet</span>
          </button>
        </div>

        {/* The Live Interactive Spreadsheet Sandbox */}
        <ExcelSimulator
          initialGrid={currentGrid}
          targetCell={slide.practiceQuestion.targetCell}
          targetRange={slide.practiceQuestion.targetRange}
          requiredShortcut={slide.practiceQuestion.requiredShortcut}
          onGridChange={(newGrid) => setCurrentGrid(newGrid)}
          onShortcutTriggered={(shortcut) => setLastShortcut(shortcut)}
          onSelectionChange={(range, active) => {
            setSelectedRangeRef(range);
            setActiveCellRef(active);
          }}
          highlightCells={slide.practiceQuestion.targetCell ? [slide.practiceQuestion.targetCell] : []}
          defaultActiveCell={slide.excel.defaultActiveCell}
        />
      </div>

      {/* =========================================================================
          SECTION 6: CHECK ANSWER (Strict Slide Structure)
      ========================================================================= */}
      <div className="bg-white border border-slate-200 rounded-lg p-5 shadow-2xs space-y-4">
        <div className="flex flex-wrap items-center justify-between gap-4">
          <div className="flex items-center gap-2 text-sm font-bold text-slate-900 uppercase tracking-wide">
            <span className="w-2 h-2 rounded-full bg-emerald-600"></span>
            <span>5. Check Answer</span>
          </div>

          <div className="flex items-center gap-3">
            <button
              onClick={handleCheckAnswer}
              className="px-5 py-2.5 bg-[#107c41] hover:bg-[#0e6e39] text-white rounded-md text-xs font-semibold shadow-xs transition-colors flex items-center gap-1.5"
            >
              <CheckCircle2 className="w-4 h-4" />
              <span>Check Answer</span>
            </button>
          </div>
        </div>

        {/* Live validation feedback banner */}
        {checkResult && (
          <div
            className={`p-4 rounded-md border text-xs space-y-1.5 transition-all ${
              checkResult.correct
                ? 'bg-emerald-50 border-emerald-300 text-emerald-950'
                : 'bg-rose-50 border-rose-300 text-rose-950'
            }`}
          >
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-2 font-bold text-sm">
                {checkResult.correct ? (
                  <>
                    <CheckCircle2 className="w-4 h-4 text-emerald-600" />
                    <span>{checkResult.message}</span>
                  </>
                ) : (
                  <>
                    <AlertCircle className="w-4 h-4 text-rose-600" />
                    <span>{checkResult.message}</span>
                  </>
                )}
              </div>

              {checkResult.correct && hasNext && (
                <button
                  onClick={onNavigateNext}
                  className="px-3 py-1 bg-emerald-700 hover:bg-emerald-800 text-white rounded text-xs font-medium flex items-center gap-1 shrink-0"
                >
                  <span>Next Lesson</span>
                  <ArrowRight className="w-3 h-3" />
                </button>
              )}
            </div>

            {checkResult.detail && (
              <p className="text-slate-600 text-xs pl-6">
                {checkResult.detail}
              </p>
            )}
          </div>
        )}
      </div>

      {/* =========================================================================
          SECTION 7: REVEAL ANSWER (Strict Slide Structure)
      ========================================================================= */}
      <div className="bg-white border border-slate-200 rounded-lg p-5 shadow-2xs space-y-4">
        <div className="flex items-center justify-between">
          <div className="flex items-center gap-2 text-sm font-bold text-slate-900 uppercase tracking-wide">
            <span className="w-2 h-2 rounded-full bg-slate-500"></span>
            <span>6. Reveal Answer</span>
          </div>

          <button
            onClick={() => setShowReveal(!showReveal)}
            className="px-3 py-1.5 border border-slate-300 hover:bg-slate-50 text-slate-700 rounded text-xs font-medium flex items-center gap-1.5 transition-colors"
          >
            {showReveal ? <EyeOff className="w-3.5 h-3.5" /> : <Eye className="w-3.5 h-3.5" />}
            <span>{showReveal ? 'Hide Answer' : 'Reveal Answer'}</span>
          </button>
        </div>

        {showReveal && (
          <div className="border border-slate-200 bg-slate-50 rounded-md p-4 space-y-3 text-xs text-slate-800 animate-fade-in">
            <div className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-200 pb-2">
              <span className="font-bold text-sm text-slate-900 font-mono">
                {slide.solution.revealAnswerText}
              </span>

              {slide.solution.autofillAction && (
                <button
                  onClick={handleAutofillAnswer}
                  className="px-3 py-1 bg-slate-900 hover:bg-slate-800 text-white rounded text-xs font-medium flex items-center gap-1"
                >
                  <Play className="w-3 h-3 text-emerald-400" />
                  <span>Try It / Fill In For Me</span>
                </button>
              )}
            </div>

            <p className="text-slate-700 leading-relaxed">
              {slide.solution.explanationText}
            </p>
          </div>
        )}
      </div>
    </div>
  );
};
