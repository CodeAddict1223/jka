import { useState, useEffect } from 'react';
import { allSlides, getSlideById, getSlideNavigation, totalSlidesCount } from './data/curriculum';
import { Header } from './components/Header';
import { Sidebar } from './components/Sidebar';
import { SimpleExcelLessonCard } from './components/SimpleExcelLessonCard';
import { ShortcutsModal } from './components/ShortcutsModal';
import { CheatSheetModal } from './components/CheatSheetModal';
import { FreeSandboxModal } from './components/FreeSandboxModal';

const STORAGE_KEY_PROGRESS = 'excel_master_completed_slides_v1';
const STORAGE_KEY_CURRENT = 'excel_master_current_slide_v1';

export default function App() {
  const [currentSlideId, setCurrentSlideId] = useState<string>(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY_CURRENT);
      if (saved && getSlideById(saved)) return saved;
    } catch {
      // ignore
    }
    return 'm6-sum-function'; // Exactly matches the user's reference screenshot
  });

  const [completedSlideIds, setCompletedSlideIds] = useState<string[]>(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY_PROGRESS);
      if (saved) return JSON.parse(saved);
    } catch {
      // ignore
    }
    return [];
  });

  const [isMobileSidebarOpen, setIsMobileSidebarOpen] = useState(false);
  const [isShortcutsModalOpen, setIsShortcutsModalOpen] = useState(false);
  const [isCheatSheetModalOpen, setIsCheatSheetModalOpen] = useState(false);
  const [isSandboxModalOpen, setIsSandboxModalOpen] = useState(false);

  // Save progress changes
  useEffect(() => {
    try {
      localStorage.setItem(STORAGE_KEY_PROGRESS, JSON.stringify(completedSlideIds));
    } catch {
      // ignore
    }
  }, [completedSlideIds]);

  // Save current slide
  useEffect(() => {
    try {
      localStorage.setItem(STORAGE_KEY_CURRENT, currentSlideId);
    } catch {
      // ignore
    }
  }, [currentSlideId]);

  const currentSlide = getSlideById(currentSlideId) || allSlides[0];
  const { prevSlide, nextSlide, currentIndex } = getSlideNavigation(currentSlideId);

  const handleCompleteSlide = (slideId: string) => {
    if (!completedSlideIds.includes(slideId)) {
      setCompletedSlideIds((prev) => [...prev, slideId]);
    }
  };

  const handleSelectSlide = (slideId: string) => {
    setCurrentSlideId(slideId);
    window.scrollTo({ top: 0, behavior: 'smooth' });
  };

  const handleNavigatePrev = () => {
    if (prevSlide) {
      handleSelectSlide(prevSlide.id);
    }
  };

  const handleNavigateNext = () => {
    if (nextSlide) {
      handleSelectSlide(nextSlide.id);
    }
  };

  const handleResetProgress = () => {
    if (window.confirm('Are you sure you want to reset your course progress?')) {
      setCompletedSlideIds([]);
      setCurrentSlideId(allSlides[0].id);
    }
  };

  return (
    <div className="min-h-screen bg-[#070e10] flex flex-col font-sans text-[#e2e8f0] antialiased selection:bg-emerald-900 selection:text-emerald-100">
      {/* 3-Zone Top Bar Contract */}
      <Header
        onOpenMobileMenu={() => setIsMobileSidebarOpen(true)}
        onOpenShortcutsModal={() => setIsShortcutsModalOpen(true)}
        onOpenCheatSheetModal={() => setIsCheatSheetModalOpen(true)}
        onOpenSandboxModal={() => setIsSandboxModalOpen(true)}
        completedCount={completedSlideIds.length}
        totalCount={totalSlidesCount}
      />

      {/* Main Layout Area */}
      <div className="flex-1 flex max-w-7xl w-full mx-auto">
        {/* Course Syllabus Sidebar */}
        <Sidebar
          currentSlideId={currentSlideId}
          completedSlideIds={completedSlideIds}
          onSelectSlide={handleSelectSlide}
          isOpen={isMobileSidebarOpen}
          onCloseMobile={() => setIsMobileSidebarOpen(false)}
        />

        {/* Lesson Canvas Area */}
        <main className="flex-1 p-3 sm:p-6 md:p-8 overflow-y-auto flex flex-col items-center">
          {/* Breadcrumb & Linear Progress Bar */}
          <div className="w-full max-w-3xl mb-4 flex flex-wrap items-center justify-between gap-4 text-xs">
            <div className="flex items-center gap-2 text-[#6e8a8d]">
              <span className="font-semibold text-emerald-400">Excel Mastery</span>
              <span aria-hidden="true">/</span>
              <span>Lesson {currentIndex + 1} of {totalSlidesCount}</span>
            </div>

            <div className="flex items-center gap-3">
              <div className="w-28 md:w-36 bg-[#122326] h-1.5 rounded-full overflow-hidden border border-[#1b3539]">
                <div
                  className="bg-[#00c274] h-full transition-all duration-300"
                  style={{ width: `${Math.round((completedSlideIds.length / totalSlidesCount) * 100)}%` }}
                />
              </div>
              <span className="font-mono text-[#8da3a6] font-semibold text-[11px]">
                {Math.round((completedSlideIds.length / totalSlidesCount) * 100)}%
              </span>
              {completedSlideIds.length > 0 && (
                <button
                  onClick={handleResetProgress}
                  className="text-[#527175] hover:text-[#8da3a6] underline text-[11px] cursor-pointer"
                >
                  Reset
                </button>
              )}
            </div>
          </div>

          {/* Clean, Simple Excel Lesson Card matching the user's reference image */}
          {currentSlide && (
            <SimpleExcelLessonCard
              slide={currentSlide}
              isCompleted={completedSlideIds.includes(currentSlide.id)}
              onCompleteSlide={handleCompleteSlide}
              onNavigatePrev={handleNavigatePrev}
              onNavigateNext={handleNavigateNext}
              hasPrev={Boolean(prevSlide)}
              hasNext={Boolean(nextSlide)}
            />
          )}
        </main>
      </div>

      {/* Supporting Modals */}
      <ShortcutsModal
        isOpen={isShortcutsModalOpen}
        onClose={() => setIsShortcutsModalOpen(false)}
      />

      <CheatSheetModal
        isOpen={isCheatSheetModalOpen}
        onClose={() => setIsCheatSheetModalOpen(false)}
      />

      <FreeSandboxModal
        isOpen={isSandboxModalOpen}
        onClose={() => setIsSandboxModalOpen(false)}
      />
    </div>
  );
}
