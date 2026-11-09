import React, { useState } from 'react';
import { X, RefreshCw } from 'lucide-react';
import { ExcelSimulator } from './ExcelSimulator';
import { GridData } from '../types/excel';

interface FreeSandboxModalProps {
  isOpen: boolean;
  onClose: () => void;
}

const DEFAULT_SANDBOX_GRID: GridData = {
  A1: { value: 'Product', format: { bold: true } },
  B1: { value: 'Qty', format: { bold: true } },
  C1: { value: 'Unit Price', format: { bold: true } },
  D1: { value: 'Total Cost', format: { bold: true } },
  A2: { value: 'Laptops' },
  B2: { value: 10 },
  C2: { value: 850, format: { format: 'currency' } },
  D2: { value: 8500, formula: '=B2*C2', format: { format: 'currency' } },
  A3: { value: 'Monitors' },
  B3: { value: 25 },
  C3: { value: 220, format: { format: 'currency' } },
  D3: { value: 5500, formula: '=B3*C3', format: { format: 'currency' } },
  A4: { value: 'Keyboards' },
  B4: { value: 50 },
  C4: { value: 45, format: { format: 'currency' } },
  D4: { value: 2250, formula: '=B4*C4', format: { format: 'currency' } },
  A5: { value: 'Grand Total', format: { bold: true } },
  D5: { value: 16250, formula: '=SUM(D2:D4)', format: { bold: true, format: 'currency' } },
};

export const FreeSandboxModal: React.FC<FreeSandboxModalProps> = ({ isOpen, onClose }) => {
  const [sandboxGrid, setSandboxGrid] = useState<GridData>(DEFAULT_SANDBOX_GRID);

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-3 md:p-6 bg-black/75 backdrop-blur-2xs">
      <div className="bg-[#0b1517] rounded-xl shadow-2xl w-full max-w-5xl max-h-[92vh] flex flex-col overflow-hidden border border-[#1b3539] text-[#e2e8f0]">
        {/* Header */}
        <div className="px-5 py-3 border-b border-[#142629] flex items-center justify-between bg-[#0c181a]">
          <div>
            <h2 className="text-sm font-bold text-white flex items-center gap-2">
              <span className="w-2.5 h-2.5 rounded-full bg-[#00c274]"></span>
              <span>Free Excel Practice Sandbox</span>
            </h2>
            <p className="text-xs text-[#6e8a8d]">Unconstrained spreadsheet workspace to practice any formula or shortcut</p>
          </div>

          <div className="flex items-center gap-2">
            <button
              onClick={() => setSandboxGrid(DEFAULT_SANDBOX_GRID)}
              className="px-2.5 py-1 text-xs font-medium text-[#8da3a6] hover:text-white bg-[#122326] border border-[#1b3539] rounded flex items-center gap-1 hover:bg-[#183135] cursor-pointer"
            >
              <RefreshCw className="w-3 h-3" />
              <span>Reset Data</span>
            </button>
            <button
              onClick={onClose}
              className="p-1 text-[#6e8a8d] hover:text-white rounded-lg hover:bg-[#122326] cursor-pointer"
            >
              <X className="w-5 h-5" />
            </button>
          </div>
        </div>

        {/* Excel Simulator Canvas */}
        <div className="flex-1 overflow-y-auto p-4 bg-[#070e10]">
          <ExcelSimulator
            initialGrid={sandboxGrid}
            onGridChange={(newGrid) => setSandboxGrid(newGrid)}
            defaultActiveCell="A1"
          />
        </div>

        {/* Footer instructions */}
        <div className="px-5 py-2.5 border-t border-[#142629] bg-[#0c181a] text-xs text-[#6e8a8d] flex items-center justify-between">
          <div className="flex items-center gap-4 text-[11px]">
            <span>Try entering: <code className="font-mono text-emerald-300 bg-[#122326] px-1 py-0.5 rounded border border-[#1b3539]">=AVERAGE(B2:B4)</code></span>
            <span>Shortcut: <kbd className="font-mono bg-[#122326] px-1 rounded border border-[#1b3539]">Alt + =</kbd> for AutoSum</span>
          </div>

          <button
            onClick={onClose}
            className="px-4 py-1 text-xs font-semibold text-[#cbd5e1] bg-[#122326] border border-[#1b3539] rounded hover:bg-[#183135] cursor-pointer"
          >
            Done Practicing
          </button>
        </div>
      </div>
    </div>
  );
};
