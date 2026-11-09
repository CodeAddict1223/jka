import { ModuleData, SlideData } from '../types/excel';
import { modules1_5 } from './modules/modules1_5';
import { modules6_10 } from './modules/modules6_10';
import { modules11_15 } from './modules/modules11_15';
import { modules16_20 } from './modules/modules16_20';
import { modules21_25 } from './modules/modules21_25';
import { modules26_30 } from './modules/modules26_30';
import { modules31_35 } from './modules/modules31_35';

export const allModules: ModuleData[] = [
  ...modules1_5,
  ...modules6_10,
  ...modules11_15,
  ...modules16_20,
  ...modules21_25,
  ...modules26_30,
  ...modules31_35,
];

// Flat list of all slides in linear order
export const allSlides: SlideData[] = allModules.flatMap((m) => m.slides);

export const totalSlidesCount = allSlides.length;

// Map for instant O(1) lookup
export const slidesMap = new Map<string, SlideData>(
  allSlides.map((slide) => [slide.id, slide])
);

export function getSlideById(id: string): SlideData | undefined {
  return slidesMap.get(id);
}

export function getSlideNavigation(currentId: string): {
  prevSlide: SlideData | null;
  nextSlide: SlideData | null;
  currentIndex: number;
  totalSlides: number;
} {
  const currentIndex = allSlides.findIndex((s) => s.id === currentId);
  const prevSlide = currentIndex > 0 ? allSlides[currentIndex - 1] : null;
  const nextSlide = currentIndex < allSlides.length - 1 ? allSlides[currentIndex + 1] : null;

  return {
    prevSlide,
    nextSlide,
    currentIndex: currentIndex !== -1 ? currentIndex : 0,
    totalSlides: allSlides.length,
  };
}

export function searchCurriculum(query: string): SlideData[] {
  if (!query || query.trim() === '') return [];
  const q = query.toLowerCase().trim();
  return allSlides.filter(
    (slide) =>
      slide.title.toLowerCase().includes(q) ||
      slide.moduleTitle.toLowerCase().includes(q) ||
      slide.explanation.overview.toLowerCase().includes(q) ||
      slide.practiceQuestion.instruction.toLowerCase().includes(q) ||
      (slide.practiceQuestion.requiredShortcut && slide.practiceQuestion.requiredShortcut.toLowerCase().includes(q))
  );
}
