import React from 'react';
import { Menu, BookOpen, Keyboard, HelpCircle } from 'lucide-react';

interface HeaderProps {
  onOpenMobileMenu: () => void;
  onOpenShortcutsModal: () => void;
  onOpenCheatSheetModal: () => void;
  onOpenSandboxModal: () => void;
  completedCount: number;
  totalCount: number;
}

export const Header: React.FC<HeaderProps> = ({
  onOpenMobileMenu,
  onOpenShortcutsModal,
  onOpenCheatSheetModal,
  onOpenSandboxModal,
  completedCount,
  totalCount,
}) => {
  return (
    <header className="sticky top-0 z-30 h-14 bg-[#0b1517] border-b border-[#142629] px-4 md:px-6 flex items-center justify-between gap-6 text-xs text-[#a0b5b7]">
      {/* Zone 1: Single text element wordmark with mobile menu toggle */}
      <div className="flex items-center gap-3 shrink-0">
        <button
          onClick={onOpenMobileMenu}
          className="p-1.5 text-[#8da3a6] hover:text-white rounded lg:hidden"
          aria-label="Open navigation menu"
        >
          <Menu className="w-5 h-5" />
        </button>

        <a 
          href="/" 
          className="text-base font-extrabold tracking-tight text-white whitespace-nowrap shrink-0 flex items-center gap-2"
        >
          <span className="w-6 h-6 rounded bg-[#00c274] text-slate-950 flex items-center justify-center text-xs font-black">
            X
          </span>
          <span>ExcelMaster</span>
        </a>
      </div>

      {/* Zone 2: 4 concise single-line nav links with subtle hover underlines */}
      <nav className="hidden md:flex items-center gap-6 text-xs font-medium text-[#8da3a6]">
        <button 
          onClick={onOpenMobileMenu}
          className="hover:text-white transition-colors whitespace-nowrap shrink-0 flex items-center gap-1.5 cursor-pointer"
        >
          <BookOpen className="w-3.5 h-3.5 text-emerald-400" />
          <span>35 Modules</span>
        </button>
        <button 
          onClick={onOpenShortcutsModal}
          className="hover:text-white transition-colors whitespace-nowrap shrink-0 flex items-center gap-1.5 cursor-pointer"
        >
          <Keyboard className="w-3.5 h-3.5 text-emerald-400" />
          <span>Shortcuts Guide</span>
        </button>
        <button 
          onClick={onOpenCheatSheetModal}
          className="hover:text-white transition-colors whitespace-nowrap shrink-0 flex items-center gap-1.5 cursor-pointer"
        >
          <HelpCircle className="w-3.5 h-3.5 text-emerald-400" />
          <span>Formula Cheat Sheet</span>
        </button>
      </nav>

      {/* Zone 3: 1 primary action */}
      <div className="flex items-center gap-3 shrink-0">
        <div className="hidden sm:flex items-center gap-1.5 text-xs text-[#6e8a8d] font-medium">
          <span>Progress:</span>
          <span className="font-semibold text-emerald-400 font-mono">
            {completedCount}/{totalCount}
          </span>
        </div>

        <button
          onClick={onOpenSandboxModal}
          className="px-3.5 py-1.5 text-xs font-semibold text-slate-950 bg-[#00c274] hover:bg-[#00d882] rounded-lg transition-colors whitespace-nowrap shrink-0 shadow-xs cursor-pointer"
        >
          Open Sandbox
        </button>
      </div>
    </header>
  );
};
