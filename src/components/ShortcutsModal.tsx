import React, { useState } from 'react';
import { X, Search } from 'lucide-react';

interface ShortcutsModalProps {
  isOpen: boolean;
  onClose: () => void;
}

interface ShortcutItem {
  keys: string;
  action: string;
  category: 'Navigation' | 'Selection' | 'Formatting' | 'Formulas & Editing';
  module: string;
}

const SHORTCUTS_DATA: ShortcutItem[] = [
  { keys: 'Arrow Keys', action: 'Move active cell by one coordinate in any direction', category: 'Navigation', module: 'Module 2' },
  { keys: 'Ctrl + Arrow Keys', action: 'Jump directly to the edge of the data block', category: 'Navigation', module: 'Module 2' },
  { keys: 'Ctrl + Home', action: 'Return immediately to cell A1 from anywhere', category: 'Navigation', module: 'Module 1' },
  { keys: 'Shift + Arrow Keys', action: 'Dynamically expand rectangular cell selection', category: 'Selection', module: 'Module 2' },
  { keys: 'Shift + Space', action: 'Select the entire active row (16,384 columns)', category: 'Selection', module: 'Module 2' },
  { keys: 'Ctrl + Space', action: 'Select the entire active column (1,048,576 rows)', category: 'Selection', module: 'Module 2' },
  { keys: 'Ctrl + A', action: 'Select current data region; second press selects entire sheet', category: 'Selection', module: 'Module 2' },
  { keys: 'Alt + ;', action: 'Select visible cells only (skips hidden rows/columns)', category: 'Selection', module: 'Module 2' },
  { keys: 'F2', action: 'Toggle cell edit mode without overwriting existing content', category: 'Formulas & Editing', module: 'Module 3' },
  { keys: 'F4', action: 'Cycle reference locks: A1 -> $A$1 -> A$1 -> $A1', category: 'Formulas & Editing', module: 'Module 9' },
  { keys: 'Alt + =', action: 'AutoSum: Automatically insert =SUM(...) formula', category: 'Formulas & Editing', module: 'Module 6' },
  { keys: 'Ctrl + D', action: 'Fill Down: Copy top cell contents/formulas downward', category: 'Formulas & Editing', module: 'Module 3' },
  { keys: 'Ctrl + R', action: 'Fill Right: Copy leftmost cell across to the right', category: 'Formulas & Editing', module: 'Module 3' },
  { keys: 'Ctrl + E', action: 'Flash Fill: Automatically parse/split text based on examples', category: 'Formulas & Editing', module: 'Module 5' },
  { keys: 'Ctrl + B', action: 'Toggle Bold text styling', category: 'Formatting', module: 'Module 3' },
  { keys: 'Ctrl + I', action: 'Toggle Italic text styling', category: 'Formatting', module: 'Module 3' },
  { keys: 'Ctrl + U', action: 'Toggle Underline text styling', category: 'Formatting', module: 'Module 3' },
  { keys: 'Ctrl + Shift + $', action: 'Format selected cells as Currency ($#,##0.00)', category: 'Formatting', module: 'Module 3' },
  { keys: 'Ctrl + Shift + %', action: 'Format selected cells as Percentage (0%)', category: 'Formatting', module: 'Module 3' },
  { keys: 'Ctrl + Shift + ~', action: 'Reset selected cells to General number format', category: 'Formatting', module: 'Module 3' },
  { keys: 'Ctrl + F', action: 'Open Find dialog to locate values across sheet', category: 'Navigation', module: 'Module 5' },
  { keys: 'Ctrl + H', action: 'Open Find and Replace dialog', category: 'Formulas & Editing', module: 'Module 5' },
  { keys: 'Ctrl + T', action: 'Convert selected data range into an official Excel Table', category: 'Formulas & Editing', module: 'Module 9' },
];

export const ShortcutsModal: React.FC<ShortcutsModalProps> = ({ isOpen, onClose }) => {
  const [filter, setFilter] = useState<string>('All');
  const [search, setSearch] = useState<string>('');

  if (!isOpen) return null;

  const categories = ['All', 'Navigation', 'Selection', 'Formatting', 'Formulas & Editing'];

  const filtered = SHORTCUTS_DATA.filter((item) => {
    const matchesCat = filter === 'All' || item.category === filter;
    const matchesSearch =
      item.keys.toLowerCase().includes(search.toLowerCase()) ||
      item.action.toLowerCase().includes(search.toLowerCase()) ||
      item.module.toLowerCase().includes(search.toLowerCase());
    return matchesCat && matchesSearch;
  });

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/70 backdrop-blur-2xs">
      <div className="bg-[#0b1517] rounded-xl shadow-2xl w-full max-w-3xl max-h-[85vh] flex flex-col overflow-hidden border border-[#1b3539] text-[#e2e8f0]">
        {/* Header */}
        <div className="px-6 py-4 border-b border-[#142629] flex items-center justify-between">
          <div>
            <h2 className="text-base font-bold text-white">Excel Keyboard Shortcuts Guide</h2>
            <p className="text-xs text-[#6e8a8d]">Curriculum-aligned keystrokes for rapid execution</p>
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

          <div className="relative w-48">
            <Search className="w-3.5 h-3.5 absolute left-2.5 top-2.5 text-[#527175]" />
            <input
              type="text"
              placeholder="Search shortcut..."
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              className="w-full pl-8 pr-2 py-1 text-xs bg-[#0e1d20] border border-[#1b3539] text-white placeholder-[#456266] rounded-md focus:outline-none focus:border-emerald-500"
            />
          </div>
        </div>

        {/* Table content */}
        <div className="flex-1 overflow-y-auto p-6">
          <div className="border border-[#193236] rounded-lg overflow-hidden">
            <table className="w-full text-left text-xs border-collapse font-mono">
              <thead className="bg-[#0c181a] border-b border-[#193236] text-[#8da3a6] font-semibold">
                <tr>
                  <th className="px-4 py-2.5">Key Combination</th>
                  <th className="px-4 py-2.5">Action &amp; Workflow</th>
                  <th className="px-4 py-2.5">Category</th>
                  <th className="px-4 py-2.5">Module</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-[#142629]">
                {filtered.map((item, idx) => (
                  <tr key={idx} className="hover:bg-[#112326]/60 transition-colors">
                    <td className="px-4 py-2.5 font-bold text-emerald-300">
                      <kbd className="bg-[#122326] border border-[#1e3c40] px-2 py-0.5 rounded shadow-2xs">
                        {item.keys}
                      </kbd>
                    </td>
                    <td className="px-4 py-2.5 text-[#cbd5e1] font-sans">{item.action}</td>
                    <td className="px-4 py-2.5 text-[#6e8a8d]">{item.category}</td>
                    <td className="px-4 py-2.5 text-[#527175]">{item.module}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>

        {/* Footer */}
        <div className="px-6 py-3 border-t border-[#142629] bg-[#0c181a] flex justify-end">
          <button
            onClick={onClose}
            className="px-4 py-1.5 text-xs font-semibold text-[#cbd5e1] bg-[#122326] border border-[#1e3c40] rounded-md hover:bg-[#1a3438]"
          >
            Close Guide
          </button>
        </div>
      </div>
    </div>
  );
};
