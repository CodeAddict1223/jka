import React, { useState } from 'react';
import { X, Search } from 'lucide-react';

interface CheatSheetModalProps {
  isOpen: boolean;
  onClose: () => void;
}

interface FormulaItem {
  name: string;
  syntax: string;
  description: string;
  example: string;
  category: 'Math' | 'Lookup' | 'Logic' | 'Text' | 'Finance & Dates';
}

const FORMULA_ITEMS: FormulaItem[] = [
  { name: 'SUM', syntax: '=SUM(number1, [number2], ...)', description: 'Adds all numbers in a range together.', example: '=SUM(B2:B10)', category: 'Math' },
  { name: 'AVERAGE', syntax: '=AVERAGE(number1, [number2], ...)', description: 'Calculates the arithmetic mean of numeric cells.', example: '=AVERAGE(C2:C20)', category: 'Math' },
  { name: 'COUNT', syntax: '=COUNT(value1, [value2], ...)', description: 'Counts only the cells containing numbers.', example: '=COUNT(A2:A50)', category: 'Math' },
  { name: 'COUNTA', syntax: '=COUNTA(value1, [value2], ...)', description: 'Counts non-empty cells across all data types.', example: '=COUNTA(A2:A50)', category: 'Math' },
  { name: 'COUNTBLANK', syntax: '=COUNTBLANK(range)', description: 'Counts empty cells to audit missing data.', example: '=COUNTBLANK(D2:D100)', category: 'Math' },
  { name: 'ROUND', syntax: '=ROUND(number, num_digits)', description: 'Rounds number to specified decimal places.', example: '=ROUND(A1, 2)', category: 'Math' },
  { name: 'IF', syntax: '=IF(logical_test, val_if_true, [val_if_false])', description: 'Executes conditional logic based on a test.', example: '=IF(B2>=70, "Pass", "Fail")', category: 'Logic' },
  { name: 'AND', syntax: '=AND(logical1, [logical2], ...)', description: 'Returns TRUE if all conditions evaluate as true.', example: '=AND(A1>50, B1="Approved")', category: 'Logic' },
  { name: 'OR', syntax: '=OR(logical1, [logical2], ...)', description: 'Returns TRUE if any condition evaluates as true.', example: '=OR(A1="Red", A1="Blue")', category: 'Logic' },
  { name: 'IFERROR', syntax: '=IFERROR(value, value_if_error)', description: 'Catches formula errors and returns custom fallback.', example: '=IFERROR(A1/B1, 0)', category: 'Logic' },
  { name: 'COUNTIF', syntax: '=COUNTIF(range, criteria)', description: 'Counts items meeting a specific condition.', example: '=COUNTIF(B2:B50, ">100")', category: 'Math' },
  { name: 'SUMIF', syntax: '=SUMIF(range, criteria, [sum_range])', description: 'Sums values meeting a single condition.', example: '=SUMIF(A2:A50, "North", C2:C50)', category: 'Math' },
  { name: 'SUMIFS', syntax: '=SUMIFS(sum_range, crit_range1, crit1, ...)', description: 'Sums values matching multiple conditions.', example: '=SUMIFS(C2:C50, A2:A50, "East", B2:B50, "Q1")', category: 'Math' },
  { name: 'VLOOKUP', syntax: '=VLOOKUP(lookup, table, col_index, [range])', description: 'Searches down first column and retrieves data to right.', example: '=VLOOKUP(E2, A2:C10, 3, FALSE)', category: 'Lookup' },
  { name: 'XLOOKUP', syntax: '=XLOOKUP(lookup, lookup_arr, return_arr, [if_not_found])', description: 'Modern bidirectional lookup with built-in fallback.', example: '=XLOOKUP(D2, A2:A10, B2:B10, "Missing")', category: 'Lookup' },
  { name: 'INDEX', syntax: '=INDEX(array, row_num, [col_num])', description: 'Returns value at coordinate intersection.', example: '=INDEX(B2:B20, 5)', category: 'Lookup' },
  { name: 'MATCH', syntax: '=MATCH(lookup_val, lookup_arr, [match_type])', description: 'Finds relative position index of an item.', example: '=MATCH("Gold", A2:A10, 0)', category: 'Lookup' },
  { name: 'CONCAT', syntax: '=CONCAT(text1, [text2], ...)', description: 'Combines multiple text strings.', example: '=CONCAT(A2, " ", B2)', category: 'Text' },
  { name: 'TEXTJOIN', syntax: '=TEXTJOIN(delimiter, ignore_empty, text1, ...)', description: 'Joins strings with custom delimiter and skips blanks.', example: '=TEXTJOIN(", ", TRUE, A2:A10)', category: 'Text' },
  { name: 'LEFT / RIGHT', syntax: '=LEFT(text, [chars]) / =RIGHT(text, [chars])', description: 'Extracts characters from beginning or end of text.', example: '=LEFT(A1, 3)', category: 'Text' },
  { name: 'TRIM', syntax: '=TRIM(text)', description: 'Removes extraneous leading/trailing spaces.', example: '=TRIM(A1)', category: 'Text' },
  { name: 'TODAY', syntax: '=TODAY()', description: 'Dynamic current system date serial.', example: '=TODAY()', category: 'Finance & Dates' },
  { name: 'PMT', syntax: '=PMT(rate, nper, pv)', description: 'Calculates periodic loan repayment amount.', example: '=PMT(0.06/12, 360, 300000)', category: 'Finance & Dates' },
  { name: 'UNIQUE', syntax: '=UNIQUE(array)', description: 'Dynamic array: Extracts distinct list of values.', example: '=UNIQUE(A2:A100)', category: 'Lookup' },
  { name: 'SORT', syntax: '=SORT(array, [sort_index], [order])', description: 'Dynamic array: Sorts results ascending or descending.', example: '=SORT(UNIQUE(A2:A100))', category: 'Lookup' },
];

export const CheatSheetModal: React.FC<CheatSheetModalProps> = ({ isOpen, onClose }) => {
  const [filter, setFilter] = useState<string>('All');
  const [search, setSearch] = useState<string>('');

  if (!isOpen) return null;

  const categories = ['All', 'Math', 'Lookup', 'Logic', 'Text', 'Finance & Dates'];

  const filtered = FORMULA_ITEMS.filter((item) => {
    const matchesCat = filter === 'All' || item.category === filter;
    const matchesSearch =
      item.name.toLowerCase().includes(search.toLowerCase()) ||
      item.description.toLowerCase().includes(search.toLowerCase()) ||
      item.syntax.toLowerCase().includes(search.toLowerCase());
    return matchesCat && matchesSearch;
  });

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/70 backdrop-blur-2xs">
      <div className="bg-[#0b1517] rounded-xl shadow-2xl w-full max-w-4xl max-h-[85vh] flex flex-col overflow-hidden border border-[#1b3539] text-[#e2e8f0]">
        {/* Header */}
        <div className="px-6 py-4 border-b border-[#142629] flex items-center justify-between">
          <div>
            <h2 className="text-base font-bold text-white">Excel Formula Cheat Sheet</h2>
            <p className="text-xs text-[#6e8a8d]">Fast syntax guide and descriptions for all core formulas</p>
          </div>
          <button
            onClick={onClose}
            className="p-1 text-[#6e8a8d] hover:text-white rounded-lg hover:bg-[#122326]"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Filter bar */}
        <div className="px-6 py-3 border-b border-[#142629] bg-[#0c181a] flex flex-wrap items-center justify-between gap-3">
          <div className="flex items-center gap-1 overflow-x-auto">
            {categories.map((cat) => (
              <button
                key={cat}
                onClick={() => setFilter(cat)}
                className={`px-3 py-1 text-xs font-medium rounded-md transition-colors ${
                  filter === cat
                    ? 'bg-[#00c274] text-slate-950 font-bold shadow-2xs'
                    : 'text-[#8da3a6] hover:text-white hover:bg-[#122326]'
                }`}
              >
                {cat}
              </button>
            ))}
          </div>

          <div className="relative w-56">
            <Search className="w-3.5 h-3.5 absolute left-2.5 top-2.5 text-[#527175]" />
            <input
              type="text"
              placeholder="Search formula (e.g. VLOOKUP)..."
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              className="w-full pl-8 pr-2 py-1 text-xs bg-[#0e1d20] border border-[#1b3539] text-white placeholder-[#456266] rounded-md focus:outline-none focus:border-emerald-500"
            />
          </div>
        </div>

        {/* Grid content */}
        <div className="flex-1 overflow-y-auto p-6">
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            {filtered.map((item, idx) => (
              <div key={idx} className="border border-[#193236] rounded-lg p-3.5 bg-[#0c181a] hover:border-[#274f55] transition-colors space-y-2">
                <div className="flex items-center justify-between">
                  <span className="font-mono font-bold text-sm text-[#00c274]">{item.name}</span>
                  <span className="text-[10px] font-semibold text-[#527175] uppercase tracking-wide">{item.category}</span>
                </div>
                <div className="font-mono text-xs bg-[#070e10] p-1.5 rounded border border-[#142629] text-emerald-300 break-all">
                  {item.syntax}
                </div>
                <p className="text-xs text-[#a0b5b7] leading-normal">{item.description}</p>
                <div className="text-[11px] text-[#6e8a8d] font-mono">
                  <span className="font-sans text-[#527175] font-medium">Example: </span>
                  {item.example}
                </div>
              </div>
            ))}
          </div>
        </div>

        {/* Footer */}
        <div className="px-6 py-3 border-t border-[#142629] bg-[#0c181a] flex justify-end">
          <button
            onClick={onClose}
            className="px-4 py-1.5 text-xs font-semibold text-[#cbd5e1] bg-[#122326] border border-[#1e3c40] rounded-md hover:bg-[#1a3438]"
          >
            Close Cheat Sheet
          </button>
        </div>
      </div>
    </div>
  );
};
