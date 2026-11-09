import React, { useState, useMemo } from 'react';
import { ModuleData } from '../types/excel';
import { allModules } from '../data/curriculum';
import { 
  CheckCircle2, Circle, ChevronDown, ChevronRight, 
  Search, X
} from 'lucide-react';

interface SidebarProps {
  currentSlideId: string;
  completedSlideIds: string[];
  onSelectSlide: (slideId: string) => void;
  isOpen: boolean;
  onCloseMobile: () => void;
}

export const Sidebar: React.FC<SidebarProps> = ({
  currentSlideId,
  completedSlideIds,
  onSelectSlide,
  isOpen,
  onCloseMobile,
}) => {
  const [searchQuery, setSearchQuery] = useState('');
  const [expandedModuleIds, setExpandedModuleIds] = useState<number[]>([1, 6, 16]);

  const toggleModule = (modId: number) => {
    setExpandedModuleIds((prev) =>
      prev.includes(modId) ? prev.filter((id) => id !== modId) : [...prev, modId]
    );
  };

  // Group modules by Level
  const groupedModules = useMemo(() => {
    const groups: { [level: string]: ModuleData[] } = {};
    allModules.forEach((m) => {
      if (!groups[m.level]) {
        groups[m.level] = [];
      }
      groups[m.level].push(m);
    });
    return groups;
  }, []);

  // Filtered slides when searching
  const filteredModules = useMemo(() => {
    if (!searchQuery.trim()) return allModules;
    const q = searchQuery.toLowerCase().trim();
    return allModules
      .map((m) => ({
        ...m,
        slides: m.slides.filter(
          (s) =>
            s.title.toLowerCase().includes(q) ||
            s.explanation.overview.toLowerCase().includes(q) ||
            s.practiceQuestion.instruction.toLowerCase().includes(q) ||
            (s.practiceQuestion.requiredShortcut && s.practiceQuestion.requiredShortcut.toLowerCase().includes(q))
        ),
      }))
      .filter((m) => m.slides.length > 0);
  }, [searchQuery]);

  return (
    <>
      {/* Mobile Backdrop */}
      {isOpen && (
        <div 
          onClick={onCloseMobile}
          className="fixed inset-0 bg-black/60 z-40 lg:hidden backdrop-blur-2xs"
        />
      )}

      <aside
        className={`fixed top-14 bottom-0 left-0 w-80 bg-[#0b1517] border-r border-[#142629] z-50 flex flex-col transition-transform duration-200 lg:static lg:translate-x-0 ${
          isOpen ? 'translate-x-0' : '-translate-x-full'
        }`}
      >
        {/* Search Header */}
        <div className="p-3 border-b border-[#142629]">
          <div className="flex items-center justify-between mb-2 lg:hidden">
            <span className="font-bold text-sm text-white">Curriculum Modules</span>
            <button 
              onClick={onCloseMobile}
              className="p-1 text-[#8da3a6] hover:text-white rounded"
            >
              <X className="w-4 h-4" />
            </button>
          </div>

          <div className="relative">
            <Search className="w-3.5 h-3.5 absolute left-2.5 top-2.5 text-[#527175]" />
            <input
              type="text"
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              placeholder="Search 35 modules, SUM, VLOOKUP..."
              className="w-full pl-8 pr-7 py-1.5 text-xs bg-[#0e1d20] border border-[#1b3539] rounded-md focus:outline-none focus:border-emerald-500 text-white placeholder-[#456266]"
            />
            {searchQuery && (
              <button
                onClick={() => setSearchQuery('')}
                className="absolute right-2 top-2 text-[#527175] hover:text-white"
              >
                <X className="w-3.5 h-3.5" />
              </button>
            )}
          </div>
        </div>

        {/* Scrollable Curriculum Tree */}
        <div className="flex-1 overflow-y-auto p-2 space-y-4">
          {searchQuery ? (
            // Search Results Mode
            <div className="space-y-3">
              <div className="text-[11px] font-semibold text-[#527175] uppercase tracking-wider px-2">
                Search Results ({filteredModules.reduce((acc, m) => acc + m.slides.length, 0)} topics)
              </div>
              {filteredModules.map((m) => (
                <div key={m.id} className="space-y-1">
                  <div className="text-xs font-bold text-emerald-400 px-2 py-0.5">
                    Module {m.id}: {m.title}
                  </div>
                  <div className="space-y-0.5 pl-2">
                    {m.slides.map((slide) => {
                      const isCurrent = slide.id === currentSlideId;
                      const isDone = completedSlideIds.includes(slide.id);
                      return (
                        <button
                          key={slide.id}
                          onClick={() => {
                            onSelectSlide(slide.id);
                            onCloseMobile();
                          }}
                          className={`w-full text-left px-2.5 py-1.5 rounded text-xs flex items-center justify-between transition-colors ${
                            isCurrent
                              ? 'bg-[#102d29] text-emerald-300 font-semibold border-l-2 border-[#00c274]'
                              : 'text-[#8da3a6] hover:bg-[#112326] hover:text-white'
                          }`}
                        >
                          <span className="truncate pr-2">{slide.title}</span>
                          {isDone ? (
                            <CheckCircle2 className="w-3.5 h-3.5 text-emerald-400 shrink-0" />
                          ) : (
                            <Circle className="w-3.5 h-3.5 text-[#243c3f] shrink-0" />
                          )}
                        </button>
                      );
                    })}
                  </div>
                </div>
              ))}
            </div>
          ) : (
            // Grouped By Level Mode
            Object.entries(groupedModules).map(([levelName, modulesInLevel]) => (
              <div key={levelName} className="space-y-1">
                {/* Level Title Header */}
                <div className="px-2 pt-2 pb-1 text-[10px] font-bold text-[#527175] uppercase tracking-wider">
                  {levelName}
                </div>

                {/* Modules under this level */}
                {modulesInLevel.map((m) => {
                  const isExpanded = expandedModuleIds.includes(m.id);
                  const completedInMod = m.slides.filter((s) => completedSlideIds.includes(s.id)).length;
                  const totalInMod = m.slides.length;
                  const isModComplete = totalInMod > 0 && completedInMod === totalInMod;

                  return (
                    <div key={m.id} className="rounded-md overflow-hidden">
                      {/* Module Header Button */}
                      <button
                        onClick={() => toggleModule(m.id)}
                        className={`w-full text-left px-2.5 py-1.5 flex items-center justify-between rounded text-xs font-semibold transition-colors ${
                          isExpanded ? 'bg-[#112326] text-white' : 'text-[#8da3a6] hover:bg-[#0e1d20]'
                        }`}
                      >
                        <div className="flex items-center gap-1.5 truncate pr-2">
                          {isExpanded ? (
                            <ChevronDown className="w-3.5 h-3.5 text-[#6e8a8d] shrink-0" />
                          ) : (
                            <ChevronRight className="w-3.5 h-3.5 text-[#456266] shrink-0" />
                          )}
                          <span className="truncate">Mod {m.id}: {m.title}</span>
                        </div>

                        <div className="flex items-center gap-1 shrink-0 text-[11px]">
                          {isModComplete ? (
                            <span className="text-emerald-400 font-bold">✓</span>
                          ) : (
                            <span className="text-[#527175] font-mono text-[10px]">
                              {completedInMod}/{totalInMod}
                            </span>
                          )}
                        </div>
                      </button>

                      {/* Subtopics Slides List */}
                      {isExpanded && (
                        <div className="pl-4 pr-1 py-1 space-y-0.5 border-l border-[#193236] ml-3 my-0.5">
                          {m.slides.map((slide) => {
                            const isCurrent = slide.id === currentSlideId;
                            const isDone = completedSlideIds.includes(slide.id);

                            return (
                              <button
                                key={slide.id}
                                onClick={() => {
                                  onSelectSlide(slide.id);
                                  onCloseMobile();
                                }}
                                className={`w-full text-left px-2 py-1.5 rounded text-xs flex items-center justify-between transition-colors ${
                                  isCurrent
                                    ? 'bg-[#102d29] text-emerald-300 font-semibold'
                                    : 'text-[#8da3a6] hover:bg-[#112326] hover:text-white'
                                }`}
                              >
                                <span className="truncate pr-2">{slide.title}</span>
                                {isDone ? (
                                  <CheckCircle2 className="w-3.5 h-3.5 text-emerald-400 shrink-0" />
                                ) : (
                                  <Circle className="w-3.5 h-3.5 text-[#243c3f] shrink-0" />
                                )}
                              </button>
                            );
                          })}
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            ))
          )}
        </div>

        {/* Sidebar Footer Metrics */}
        <div className="p-3 border-t border-[#142629] bg-[#0c181a] text-xs text-[#6e8a8d] flex items-center justify-between">
          <span>Overall Progress</span>
          <span className="font-semibold text-emerald-400 font-mono">
            {completedSlideIds.length} completed
          </span>
        </div>
      </aside>
    </>
  );
};
