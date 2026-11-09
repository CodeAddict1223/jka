import { ModuleData } from '../../types/excel';

export const modules26_30: ModuleData[] = [
  {
    id: 26,
    title: 'Advanced Financial Engineering & Depreciation',
    level: 'Level 6: Dynamic Arrays & Advanced Analytics',
    description: 'Calculate asset depreciation and irregular returns: SLN straight-line, DB declining balance, XNPV, XIRR, IPMT interest portions, and PPMT principal amortization.',
    slides: [
      {
        id: 'm26-sln-depreciation',
        moduleId: 26,
        moduleTitle: 'Module 26: Advanced Financial Engineering & Depreciation',
        topicIndex: 1,
        title: 'Straight-Line Depreciation (SLN)',
        explanation: {
          overview: 'Straight-Line Depreciation (SLN) spreads the depreciable cost of a capital asset evenly across its useful operating life. Depreciable cost equals Initial Cost minus Estimated Salvage Value.',
          keyPoints: [
            'Syntax: =SLN(cost, salvage, life).',
            'Constant Annual Expense: Produces the exact same deduction each year.',
            'Tax & Accounting standard for office equipment and fixtures.',
          ],
          tips: 'Ensure salvage value is zero or an explicitly estimated scrap recovery value.',
        },
        example: {
          description: 'A server costs $50,000, has a $5,000 scrap salvage value, and 5-year life. =SLN(50000, 5000, 5) equals $9,000 annual depreciation.',
          outputDescription: 'Even amortization schedule.',
        },
        practiceQuestion: {
          instruction: 'In cell B4, calculate the annual straight-line depreciation by subtracting Salvage (B2) from Cost (B1) and dividing by Life (B3): =(B1-B2)/B3.',
          hint: 'Select cell B4, enter =(B1-B2)/B3 and press Enter.',
          targetCell: 'B4',
          practiceType: 'formula',
        },
        excel: {
          defaultActiveCell: 'B4',
          initialGrid: {
            A1: { value: 'Asset Cost ($)', format: { bold: true } },
            B1: { value: 60000 },
            A2: { value: 'Salvage Value ($)', format: { bold: true } },
            B2: { value: 6000 },
            A3: { value: 'Useful Life (Years)', format: { bold: true } },
            B3: { value: 6 },
            A4: { value: 'Annual SLN Depreciation', format: { bold: true } },
            B4: { value: '' },
          },
        },
        solution: {
          expectedFormula: ['=(B1-B2)/B3', '=SLN(B1, B2, B3)'],
          expectedValue: 9000,
          revealAnswerText: 'Formula: =(B1-B2)/B3 or =SLN(B1, B2, B3)',
          explanationText: '(60,000 - 6,000) / 6 = 54,000 / 6 = $9,000 annual depreciation.',
          autofillAction: {
            targetCell: 'B4',
            formula: '=(B1-B2)/B3',
            value: 9000,
            format: { format: 'currency', bold: true },
          },
        },
      },
    ],
  },
  {
    id: 27,
    title: 'Accounting & Bank Reconciliations',
    level: 'Level 6: Dynamic Arrays & Advanced Analytics',
    description: 'Reconcile ledgers with bank statements: cleaning ERP narrations, duplicate transaction detection, and automated variance analysis.',
    slides: [
      {
        id: 'm27-reconciliation',
        moduleId: 27,
        moduleTitle: 'Module 27: Accounting & Bank Reconciliations',
        topicIndex: 1,
        title: 'Reconciliation Variance & Discrepancy Flags',
        explanation: {
          overview: 'Monthly bank reconciliations compare cash balances recorded in company general ledgers against actual external bank statements. Variances are isolated using delta formulas and flag rules.',
          keyPoints: [
            'Variance Formula: =Bank_Amount - Ledger_Amount.',
            'Reconciliation Flag: =IF(Variance=0, "Balanced", "Discrepancy").',
            'Automated matching flags timing differences like uncleared checks or unrecorded fees.',
          ],
          tips: 'Conditional formatting can turn rows with non-zero variances soft coral red to accelerate auditing.',
        },
        example: {
          description: 'Bank shows $10,500 and Ledger shows $10,200. Variance of +$300 prompts investigation.',
          outputDescription: 'Pinpoints accounting misalignments.',
        },
        practiceQuestion: {
          instruction: 'In cell C2, calculate the variance between Bank Statement (A2) and General Ledger (B2): =A2-B2.',
          hint: 'Select cell C2, enter =A2-B2 and press Enter.',
          targetCell: 'C2',
          practiceType: 'formula',
        },
        excel: {
          defaultActiveCell: 'C2',
          initialGrid: {
            A1: { value: 'Bank Statement', format: { bold: true } },
            B1: { value: 'General Ledger', format: { bold: true } },
            C1: { value: 'Variance ($)', format: { bold: true } },
            A2: { value: 14500 },
            B2: { value: 14250 },
            C2: { value: '' },
          },
        },
        solution: {
          expectedFormula: '=A2-B2',
          expectedValue: 250,
          revealAnswerText: 'Formula: =A2-B2',
          explanationText: '14,500 - 14,250 = $250 variance to investigate.',
          autofillAction: {
            targetCell: 'C2',
            formula: '=A2-B2',
            value: 250,
            format: { format: 'currency' },
          },
        },
      },
    ],
  },
  {
    id: 28,
    title: 'Conditional Formatting & Visual Highlighting',
    level: 'Level 6: Dynamic Arrays & Advanced Analytics',
    description: 'Transform raw numbers into executive visual dashboards: Cell highlight rules, in-cell Data Bars, Color Scales, Icon Sets, and formula-driven row highlighting.',
    slides: [
      {
        id: 'm28-highlight-rules',
        moduleId: 28,
        moduleTitle: 'Module 28: Conditional Formatting & Visual Highlighting',
        topicIndex: 1,
        title: 'Highlight Cell Rules & Visual Signals',
        explanation: {
          overview: 'Conditional formatting automatically applies font styles, borders, and fills to cells that meet specified rules (e.g. greater than a threshold, containing specific text, or duplicate values).',
          keyPoints: [
            'Highlight Rules: Greater Than, Less Than, Between, Equal To, Text That Contains.',
            'Data Bars: Micro bar charts rendered directly in cell backgrounds.',
            'Color Scales: 2-color or 3-color heatmaps across numerical gradients.',
            'Formula-Driven Rules: Use formulas like =$C2="Overdue" to style entire rows.',
          ],
          tips: 'Keep color intensity soft (pastel tints) to maintain executive readability and avoid visual fatigue.',
        },
        example: {
          description: 'Sales targets: Values >= $50,000 receive soft emerald green fill; below $20,000 receive soft rose red.',
          outputDescription: 'Instant visual scanning of business health.',
        },
        practiceQuestion: {
          instruction: 'Select cell B3 (the highest sales figure) and format it as Bold (Ctrl + B) to signify top performance.',
          hint: 'Select cell B3 and press Ctrl + B.',
          targetCell: 'B3',
          requiredShortcut: 'Ctrl + B',
          practiceType: 'shortcut',
        },
        excel: {
          defaultActiveCell: 'B3',
          initialGrid: {
            A1: { value: 'Rep', format: { bold: true } },
            B1: { value: 'Revenue', format: { bold: true } },
            A2: { value: 'Arthur' },
            B2: { value: 34000 },
            A3: { value: 'Brenda' },
            B3: { value: 78000, format: { bold: false, bgColor: '#dcfce7' } }, // Highlighted top performer
            A4: { value: 'Caleb' },
            B4: { value: 29000 },
          },
        },
        solution: {
          expectedShortcut: 'Ctrl + B',
          expectedFormat: { bold: true },
          revealAnswerText: 'Shortcut: Press Ctrl + B on cell B3',
          explanationText: 'Applies bold text to emphasize Brenda’s $78,000 top performance.',
          autofillAction: {
            targetCell: 'B3',
            value: 78000,
            format: { bold: true, bgColor: '#dcfce7' },
          },
        },
      },
    ],
  },
  {
    id: 29,
    title: 'Dynamic Arrays (Excel 365 Fundamentals)',
    level: 'Level 6: Dynamic Arrays & Advanced Analytics',
    description: 'Harness the Excel calculation engine transformation: UNIQUE distinct lists, SORT, SORTBY, FILTER, and the Spill Operator (#).',
    slides: [
      {
        id: 'm29-unique-sort',
        moduleId: 29,
        moduleTitle: 'Module 29: Dynamic Arrays (Excel 365 Fundamentals)',
        topicIndex: 1,
        title: 'UNIQUE & SORT Dynamic Arrays',
        explanation: {
          overview: 'Dynamic Array formulas return multiple results that automatically "spill" into adjacent empty cells. UNIQUE extracts distinct deduplicated lists. SORT arranges results ascending or descending.',
          keyPoints: [
            '=UNIQUE(array): Extracts one instance of each distinct entry.',
            '=SORT(array, [sort_index], [sort_order]): Sorts arrays dynamically.',
            'Spill Range (#): Reference the entire spilled output using cell# (e.g. D2#).',
            'If a cell in the spill path is blocked, Excel returns a #SPILL! error.',
          ],
          tips: 'Combine functions: =SORT(UNIQUE(A2:A100)) returns a perfectly sorted alphabetical list of unique categories.',
        },
        example: {
          description: 'A list with duplicate regions "North", "South", "North", "East". =UNIQUE(A2:A5) outputs "North", "South", "East".',
          outputDescription: 'Zero VBA scripts required for automated distinct extraction.',
        },
        practiceQuestion: {
          instruction: 'In cell C2, count the number of unique regions using =COUNTA(C2:C4) (where unique regions have spilled).',
          hint: 'Select cell C2 and enter =COUNTA(C2:C4) or count unique items.',
          targetCell: 'D2',
          practiceType: 'formula',
        },
        excel: {
          defaultActiveCell: 'D2',
          initialGrid: {
            A1: { value: 'Raw Region', format: { bold: true } },
            C1: { value: 'Spilled Unique', format: { bold: true } },
            D1: { value: 'Count Unique', format: { bold: true } },
            A2: { value: 'East' },
            C2: { value: 'East' },
            D2: { value: '' },
            A3: { value: 'West' },
            C3: { value: 'West' },
            A4: { value: 'East' },
            C4: { value: 'North' },
            A5: { value: 'North' },
          },
        },
        solution: {
          expectedFormula: ['=COUNTA(C2:C4)', '=COUNT(C2:C4)'],
          expectedValue: 3,
          revealAnswerText: 'Formula: =COUNTA(C2:C4)',
          explanationText: 'Counts the 3 unique extracted regions (East, West, North).',
          autofillAction: {
            targetCell: 'D2',
            formula: '=COUNTA(C2:C4)',
            value: 3,
          },
        },
      },
    ],
  },
  {
    id: 30,
    title: 'Advanced Array Reshaping & Manipulation',
    level: 'Level 6: Dynamic Arrays & Advanced Analytics',
    description: 'Reshape and assemble tables dynamically: VSTACK, HSTACK, TAKE, DROP, CHOOSECOLS, TEXTSPLIT, and TOCOL.',
    slides: [
      {
        id: 'm30-vstack-hstack',
        moduleId: 30,
        moduleTitle: 'Module 30: Advanced Array Reshaping & Manipulation',
        topicIndex: 1,
        title: 'VSTACK & HSTACK Array Stacking',
        explanation: {
          overview: 'VSTACK stacks multiple ranges or tables vertically on top of each other into a single continuous array. HSTACK joins ranges side-by-side horizontally.',
          keyPoints: [
            '=VSTACK(array1, array2, ...): Combines multiple monthly sheets or regional tables into a unified consolidated dataset.',
            '=HSTACK(array1, array2, ...): Joins non-adjacent columns side-by-side.',
            '=TAKE(array, rows): Grabs the first N rows from a stacked array.',
            '=DROP(array, rows): Removes the first N rows (e.g. duplicate headers).',
          ],
          tips: 'Consolidating 12 monthly sheets previously required complex Power Query steps; now =VSTACK(Jan:Dec!A2:D50) does it in one formula.',
        },
        example: {
          description: 'Stack Table 1 (rows 2-3) and Table 2 (rows 4-5) vertically into a master summary.',
          outputDescription: 'Zero copy-paste table consolidation.',
        },
        practiceQuestion: {
          instruction: 'In cell C2, calculate the combined sum of Table 1 (A2:A3) and Table 2 (B2:B3) using =SUM(A2:A3)+SUM(B2:B3) or =SUM(A2:B3).',
          hint: 'Select cell C2, enter =SUM(A2:B3) and press Enter.',
          targetCell: 'C2',
          practiceType: 'formula',
        },
        excel: {
          defaultActiveCell: 'C2',
          initialGrid: {
            A1: { value: 'Batch 1', format: { bold: true } },
            B1: { value: 'Batch 2', format: { bold: true } },
            C1: { value: 'Combined Total', format: { bold: true } },
            A2: { value: 150 },
            B2: { value: 200 },
            C2: { value: '' },
            A3: { value: 250 },
            B3: { value: 300 },
          },
        },
        solution: {
          expectedFormula: ['=SUM(A2:B3)', '=SUM(A2:A3)+SUM(B2:B3)'],
          expectedValue: 900,
          revealAnswerText: 'Formula: =SUM(A2:B3)',
          explanationText: '150 + 250 + 200 + 300 = 900 units consolidated.',
          autofillAction: {
            targetCell: 'C2',
            formula: '=SUM(A2:B3)',
            value: 900,
            format: { bold: true },
          },
        },
      },
    ],
  },
];
