export type CellFormat = {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  align?: 'left' | 'center' | 'right';
  format?: 'general' | 'currency' | 'percentage' | 'number' | 'date';
  bgColor?: string;
  textColor?: string;
  border?: boolean;
};

export type CellData = {
  value: string | number;
  formula?: string;
  format?: CellFormat;
  comment?: string;
};

export type GridData = {
  [cellRef: string]: CellData; // e.g., 'A1', 'B2'
};

export type CheckResult = {
  correct: boolean;
  message: string;
  detail?: string;
};

export type SlidePracticeType = 'formula' | 'shortcut' | 'data_entry' | 'selection' | 'interactive_tool';

export interface SlideData {
  id: string;
  moduleId: number;
  moduleTitle: string;
  topicIndex: number;
  title: string;
  
  // Mandatory slide sections strictly specified by user:
  explanation: {
    overview: string;
    keyPoints: string[];
    tips?: string;
  };
  example: {
    description: string;
    formula?: string;
    sampleTable?: {
      headers: string[];
      rows: (string | number)[][];
    };
    outputDescription: string;
  };
  practiceQuestion: {
    instruction: string;
    hint?: string;
    targetCell?: string;
    targetRange?: string;
    requiredShortcut?: string; // e.g. "Alt + =", "Ctrl + B", "Shift + Space"
    practiceType: SlidePracticeType;
  };
  excel: {
    initialGrid: GridData;
    colCount?: number;
    rowCount?: number;
    highlightCells?: string[];
    defaultActiveCell?: string;
    sheetName?: string;
    hiddenRows?: number[];
    hiddenCols?: string[];
  };
  solution: {
    expectedFormula?: string | string[]; // case-insensitive regex or strings
    expectedValue?: string | number | boolean | (string | number | boolean)[];
    expectedShortcut?: string;
    expectedRange?: string;
    expectedFormat?: Partial<CellFormat>;
    revealAnswerText: string;
    explanationText: string;
    autofillAction?: {
      targetCell: string;
      formula?: string;
      value: string | number;
      format?: CellFormat;
    };
  };
  visualExplanation?: {
    type: 'ribbon' | 'status_bar' | 'freeze_panes' | 'formula_bar' | 'pivot_table' | 'chart' | 'data_validation';
    caption: string;
  };
}

export interface ModuleData {
  id: number;
  title: string;
  level: string; // e.g., "Level 1: Foundations"
  description: string;
  slides: SlideData[];
}

export interface UserProgress {
  completedSlideIds: string[];
  currentSlideId: string;
  lastVisitedTimestamp: number;
}
