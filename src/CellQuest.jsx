import React, { useReducer, useEffect, useMemo, useRef, useState, useCallback } from 'react';

// ============================================================================
// CELL QUEST — Interactive Excel Fundamentals Quiz
// Workshop: "Demystifying Data: A Beginner's Guide to Microsoft Excel"
// Self-contained single-file React 18+ app (no network calls, no external assets)
// ============================================================================

const DIFFICULTIES = {
  easy: {
    id: 'easy',
    name: 'Sheet Novice',
    shortDesc: 'Relaxed pace, unlimited lives — ideal for warming up.',
    timePerQuestion: 45,
    lives: Infinity,
    basePoints: 10,
    streakBonus: 0,
    badgeColor: '#2E9E5B',
    runLength: 10,
  },
  medium: {
    id: 'medium',
    name: 'Cell Ranger',
    shortDesc: '25s timer, 3 lives, and +5 streak bonus per combo.',
    timePerQuestion: 25,
    lives: 3,
    basePoints: 20,
    streakBonus: 5,
    badgeColor: '#4E8FE0',
    runLength: 10,
  },
  hard: {
    id: 'hard',
    name: 'Formula Master',
    shortDesc: '15s timer, 2 lives, +10 streak bonus. One slip costs a heart.',
    timePerQuestion: 15,
    lives: 2,
    basePoints: 30,
    streakBonus: 10,
    badgeColor: '#FF7000',
    runLength: 9,
  },
};

const TOPICS = ['Workbook & Grid', 'Interface', 'Formulas', 'Functions', 'Workflow'];

const TOPIC_SHORT = {
  'Workbook & Grid': 'GRID',
  Interface: 'INTERFACE',
  Formulas: 'FORMULAS',
  Functions: 'FUNCTIONS',
  Workflow: 'WORKFLOW',
};

const CORRECT_FLAVORS = [
  'Nice, that formula slaps.',
  'Clean cell work!',
  'Grid tamed. Next cell!',
  'Zero #REF! errors here.',
  'Pixel approves this calculation.',
  'Sharp eye on the sheet!',
];

const WRONG_FLAVORS = [
  'Tricky one! Let’s check the sheet.',
  'Close! Even pros hit #VALUE! sometimes.',
  'Good try — here’s the Excel rule:',
  'No worries, Pixel has your back.',
];

// ============================================================================
// QUESTION BANK — 30 Workshop-Aligned Questions (10 Easy, 11 Medium, 9 Hard)
// ============================================================================

const QUESTION_BANK = [
  // ------------------------------ EASY (10) ------------------------------
  {
    id: 'e1',
    difficulty: 'easy',
    type: 'mcq',
    topic: 'Workbook & Grid',
    prompt: 'What is the difference between a workbook and a worksheet in Excel?',
    options: [
      'A workbook is the whole .xlsx file; worksheets are the individual tabs inside it.',
      'A worksheet is the whole .xlsx file; workbooks are the tabs inside it.',
      'A workbook holds only charts, while a worksheet holds only numbers.',
      'They are two different names for the exact same thing.',
    ],
    correctAnswer: 'A workbook is the whole .xlsx file; worksheets are the individual tabs inside it.',
    explanation: 'A workbook is the entire .xlsx file, and it can contain multiple worksheet tabs along the bottom.',
  },
  {
    id: 'e2',
    difficulty: 'easy',
    type: 'mcq',
    topic: 'Workbook & Grid',
    prompt: 'In a well-structured dataset, how are columns and rows organized?',
    options: [
      'Columns run vertically (one variable each); rows run horizontally (one observation each).',
      'Columns run horizontally (one observation each); rows run vertically (one variable each).',
      'Columns hold only text labels; rows hold only formulas.',
      'Columns and rows can be swapped freely without affecting formulas.',
    ],
    correctAnswer: 'Columns run vertically (one variable each); rows run horizontally (one observation each).',
    explanation: 'Columns go up and down and represent variables (like Height or Score), while rows go left to right and represent individual observations.',
  },
  {
    id: 'e3',
    difficulty: 'easy',
    type: 'mcq',
    topic: 'Workbook & Grid',
    prompt: 'How is a cell’s coordinate address written in Excel?',
    options: [
      'Column letter followed by row number (e.g., C4)',
      'Row number followed by column letter (e.g., 4C)',
      'Worksheet name followed by row number (e.g., Sheet4)',
      'Column number followed by row number (e.g., 3,4)',
    ],
    correctAnswer: 'Column letter followed by row number (e.g., C4)',
    explanation: 'Every cell sits at the intersection of a column and a row, named with its column letter first and row number second—like C4.',
  },
  {
    id: 'e4',
    difficulty: 'easy',
    type: 'mcq',
    topic: 'Interface',
    prompt: 'Look at the mini-spreadsheet below. Which cell is currently the active cell?',
    visual: {
      activeCell: 'C2',
      cells: {
        A1: 'ID', B1: 'Group', C1: 'Score', D1: 'Pass',
        A2: '101', B2: 'A', C2: '88', D2: 'Yes',
        A3: '102', B3: 'B', C3: '74', D3: 'Yes',
        A4: '103', B4: 'A', C4: '91', D4: 'Yes',
      },
      formulaBar: '88',
      nameBox: 'C2',
    },
    options: ['C2', 'B2', 'C3', '2C'],
    correctAnswer: 'C2',
    explanation: 'The active cell is always marked by Excel’s bold green border (and its coordinate appears in the Name Box).',
  },
  {
    id: 'e5',
    difficulty: 'easy',
    type: 'cellclick',
    topic: 'Workbook & Grid',
    prompt: 'Click cell C4 on the spreadsheet grid below.',
    visual: {
      mode: 'single',
      cells: {
        A1: 'Item', B1: 'Qty', C1: 'Cost', D1: 'Note',
        A2: 'Vial', B2: '4', C2: '12', D2: 'OK',
        A3: 'Slide', B3: '10', C3: '5', D3: 'OK',
        A4: 'Tray', B4: '2', C4: '18', D4: 'New',
      },
    },
    correctAnswer: 'C4',
    explanation: 'Column C (third column) intersecting Row 4 (fourth row) lands directly on cell C4.',
  },
  {
    id: 'e6',
    difficulty: 'easy',
    type: 'mcq',
    topic: 'Formulas',
    prompt: 'Which character MUST start every formula or calculation in Excel?',
    options: ['=', '+', '#', '$'],
    correctAnswer: '=',
    explanation: 'Typing an equals sign (=) tells Excel to evaluate the cell as a calculation instead of plain text.',
  },
  {
    id: 'e7',
    difficulty: 'easy',
    type: 'mcq',
    topic: 'Formulas',
    prompt: 'Which symbols are used for multiplication and division in Excel formulas?',
    options: [
      '* for multiplication and / for division',
      'x for multiplication and ÷ for division',
      '^ for multiplication and % for division',
      '& for multiplication and \\ for division',
    ],
    correctAnswer: '* for multiplication and / for division',
    explanation: 'Excel uses the asterisk (*) to multiply and the forward slash (/) to divide.',
  },
  {
    id: 'e8',
    difficulty: 'easy',
    type: 'truefalse',
    topic: 'Functions',
    prompt: 'True or False: Excel has a built-in =SUBTRACT() function for subtracting cells.',
    options: ['True', 'False'],
    correctAnswer: 'False',
    explanation: 'Excel has no SUBTRACT function—to subtract, simply use the minus operator, such as =A1-B1.',
  },
  {
    id: 'e9',
    difficulty: 'easy',
    type: 'truefalse',
    topic: 'Interface',
    prompt: 'True or False: The active cell is the cell outlined with a bold green border where your typing goes.',
    options: ['True', 'False'],
    correctAnswer: 'True',
    explanation: 'The bold green border highlights the active cell—any value or formula you type is entered there.',
  },
  {
    id: 'e10',
    difficulty: 'easy',
    type: 'builder',
    topic: 'Formulas',
    prompt: 'Assemble a formula that adds cell A1 and cell B1 together.',
    goalLabel: 'Add A1 and B1',
    tokens: ['=', 'A1', '+', 'B1', '-', '*', 'SUM', '(', ')'],
    correctAnswer: '=A1+B1',
    explanation: 'Every calculation starts with =, followed by the cell references joined by the + operator: =A1+B1.',
  },

  // ------------------------------ MEDIUM (11) ------------------------------
  {
    id: 'm1',
    difficulty: 'medium',
    type: 'mcq',
    topic: 'Interface',
    prompt: 'You see the number 42 in cell D2. Where do you look to find out whether 42 was typed by hand or computed by a formula?',
    options: [
      'The Formula Bar',
      'The Name Box',
      'The Sheet Tab at the bottom',
      'The Column Header',
    ],
    correctAnswer: 'The Formula Bar',
    explanation: 'The Formula Bar reveals the true contents behind a cell’s displayed value—showing either a raw number or the underlying formula.',
  },
  {
    id: 'm2',
    difficulty: 'medium',
    type: 'mcq',
    topic: 'Interface',
    prompt: 'What does the Name Box (to the left of the Formula Bar) display?',
    options: [
      'The coordinate address of the active cell (like B3)',
      'The file name of the .xlsx workbook',
      'The sum of all selected cells',
      'The list of available Excel functions',
    ],
    correctAnswer: 'The coordinate address of the active cell (like B3)',
    explanation: 'The Name Box shows the active cell’s coordinate (or named range) so you always know where you are on the grid.',
  },
  {
    id: 'm3',
    difficulty: 'medium',
    type: 'mcq',
    topic: 'Formulas',
    prompt: 'Which statement accurately describes the relationship between formulas and functions?',
    options: [
      'Every function is a formula, but not every formula uses a function.',
      'Every formula uses a function, but not every function is a formula.',
      'Formulas only work on text, while functions only work on numbers.',
      'Functions do not start with an equals sign (=), whereas formulas do.',
    ],
    correctAnswer: 'Every function is a formula, but not every formula uses a function.',
    explanation: 'A formula is any calculation starting with = (like =A1+B1); a function is a built-in shortcut (like =SUM(A1:B1)) used inside a formula.',
  },
  {
    id: 'm4',
    difficulty: 'medium',
    type: 'mcq',
    topic: 'Formulas',
    prompt: 'In the range reference A1:A5, what does the colon (:) mean?',
    options: [
      '"Through" — it includes A1, A5, and every cell between them.',
      '"And only" — it includes only A1 and A5, skipping A2–A4.',
      '"Divided by" — it divides A1 by A5.',
      '"Compare" — it checks whether A1 equals A5.',
    ],
    correctAnswer: '"Through" — it includes A1, A5, and every cell between them.',
    explanation: 'The colon means "through," defining a continuous rectangular block from the first cell to the last cell.',
  },
  {
    id: 'm5',
    difficulty: 'medium',
    type: 'mcq',
    topic: 'Functions',
    prompt: 'How does =AVERAGE() treat blank cells, text cells, and cells containing 0?',
    options: [
      'It ignores blanks and text, but INCLUDES zeros in the average.',
      'It ignores blanks, text, AND zeros.',
      'It treats blanks as zeros, and ignores text.',
      'It returns an error if any blank cell is inside the range.',
    ],
    correctAnswer: 'It ignores blanks and text, but INCLUDES zeros in the average.',
    explanation: 'AVERAGE skips empty cells and text, but a typed 0 is a valid number and pulls the mean downward.',
  },
  {
    id: 'm6',
    difficulty: 'medium',
    type: 'mcq',
    topic: 'Functions',
    prompt: 'Why are =MIN() and =MAX() especially useful right after importing a dataset?',
    options: [
      'They spot potential outliers and verify that values fall within expected boundaries.',
      'They automatically delete duplicate rows from the worksheet.',
      'They sort the entire table from smallest to largest.',
      'They convert text values into currency formats.',
    ],
    correctAnswer: 'They spot potential outliers and verify that values fall within expected boundaries.',
    explanation: 'Checking MIN and MAX quickly exposes data-entry typos (like an age of 999 or a negative height) and confirms valid boundaries.',
  },
  {
    id: 'm7',
    difficulty: 'medium',
    type: 'mcq',
    topic: 'Workflow',
    prompt: 'What is the recommended best practice for organizing raw data, cleaned data, and summary statistics?',
    options: [
      'Keep them on separate worksheets inside the same workbook.',
      'Overwrite the raw data directly so there is only one sheet.',
      'Save each step as a completely separate .xlsx file on your Desktop.',
      'Put summary statistics in the middle of the raw data rows.',
    ],
    correctAnswer: 'Keep them on separate worksheets inside the same workbook.',
    explanation: 'Keeping Raw, Cleaned, and Summary tabs in one workbook preserves your original data untouched while keeping the project organized.',
  },
  {
    id: 'm8',
    difficulty: 'medium',
    type: 'mcq',
    topic: 'Workflow',
    prompt: 'Why should you use "Format as Table" and apply proper number formats (Currency, Percentage, Date)?',
    options: [
      'Tables enable fast sorting and filtering, and number formats give raw numbers clear context.',
      'Tables lock the workbook with a password so nobody can read it.',
      'Number formats permanently round and erase the underlying decimal digits.',
      'Tables replace all formulas with static text values.',
    ],
    correctAnswer: 'Tables enable fast sorting and filtering, and number formats give raw numbers clear context.',
    explanation: 'Format as Table adds instant sort/filter controls and structured styling, while number formats clarify units without changing the true value.',
  },
  {
    id: 'm9',
    difficulty: 'medium',
    type: 'cellclick',
    topic: 'Workbook & Grid',
    prompt: 'Select EVERY cell covered by the range A1:B3, then click Confirm Selection.',
    visual: {
      mode: 'range',
      cells: {
        A1: '10', B1: '20', C1: '30', D1: '40',
        A2: '15', B2: '25', C2: '35', D2: '45',
        A3: '12', B3: '22', C3: '32', D3: '42',
        A4: '18', B4: '28', C4: '38', D4: '48',
      },
    },
    correctAnswer: ['A1', 'A2', 'A3', 'B1', 'B2', 'B3'],
    explanation: 'A1:B3 spans columns A through B and rows 1 through 3—a 2×3 rectangle of 6 cells: A1, A2, A3, B1, B2, and B3.',
  },
  {
    id: 'm10',
    difficulty: 'medium',
    type: 'builder',
    topic: 'Functions',
    prompt: 'Build a formula to compute the mean of cells B2 through B6.',
    goalLabel: 'Average of B2:B6',
    tokens: ['=', 'AVERAGE', 'SUM', 'MIN', '(', ')', 'B2', ':', 'B6', ','],
    correctAnswer: '=AVERAGE(B2:B6)',
    explanation: 'Use =AVERAGE(B2:B6) with a colon between B2 and B6 to average every numeric cell in that range.',
  },
  {
    id: 'm11',
    difficulty: 'medium',
    type: 'truefalse',
    topic: 'Interface',
    prompt: 'True or False: When you click a cell showing "150", the Formula Bar might show =SUM(B2:B4) instead of 150.',
    options: ['True', 'False'],
    correctAnswer: 'True',
    explanation: 'The cell displays the calculated result (150), while the Formula Bar reveals the formula (=SUM(B2:B4)) that produced it.',
  },

  // ------------------------------ HARD (9) ------------------------------
  {
    id: 'h1',
    difficulty: 'hard',
    type: 'mcq',
    topic: 'Functions',
    prompt: 'Look at the spreadsheet below. What value will =SUM(B2:B4) return?',
    visual: {
      shadedRange: ['B2', 'B3', 'B4'],
      cells: {
        A1: 'Lab', B1: 'Samples', C1: 'Cost', D1: 'Flag',
        A2: 'North', B2: '12', C2: '5', D2: '0',
        A3: 'South', B3: '18', C3: '8', D3: '1',
        A4: 'East', B4: '10', C4: '4', D4: '0',
      },
      formulaBar: '=SUM(B2:B4)',
      nameBox: 'B5',
    },
    options: ['40', '22', '30', '13.3'],
    correctAnswer: '40',
    explanation: 'B2:B4 includes B2 (12), B3 (18), and B4 (10). Their sum is 12 + 18 + 10 = 40.',
  },
  {
    id: 'h2',
    difficulty: 'hard',
    type: 'mcq',
    topic: 'Functions',
    prompt: 'Look at Row 1 below. What will =AVERAGE(A1:C1) return?',
    visual: {
      shadedRange: ['A1', 'B1', 'C1'],
      cells: {
        A1: '10', B1: '0', C1: '20', D1: '30',
        A2: '5', B2: '5', C2: '5', D2: '5',
        A3: '', B3: '', C3: '', D3: '',
        A4: '', B4: '', C4: '', D4: '',
      },
      formulaBar: '=AVERAGE(A1:C1)',
      nameBox: 'D2',
    },
    options: ['10', '15', '30', '20'],
    correctAnswer: '10',
    explanation: 'AVERAGE includes the 0 in B1! So (10 + 0 + 20) / 3 = 10, not 15.',
  },
  {
    id: 'h3',
    difficulty: 'hard',
    type: 'mcq',
    topic: 'Formulas',
    prompt: 'Cell C1 contains =A1+B1, where A1 is 10 and B1 is 5 (so C1 shows 15). If you change A1 to 25, what happens to C1?',
    options: [
      'C1 automatically recalculates and updates to 30.',
      'C1 stays 15 until you retype the formula.',
      'C1 displays #REF! because A1 was modified.',
      'C1 changes to 25 because it copies A1.',
    ],
    correctAnswer: 'C1 automatically recalculates and updates to 30.',
    explanation: 'Dynamic recalculation is Excel’s superpower: formulas reference cells rather than hardcoded values, so outputs update automatically.',
  },
  {
    id: 'h4',
    difficulty: 'hard',
    type: 'builder',
    topic: 'Functions',
    prompt: 'Assemble the formula to calculate the total of cells B2 through B6.',
    goalLabel: 'Total of B2:B6',
    tokens: ['=', 'SUM', 'AVERAGE', 'MIN', 'MAX', '(', ')', 'B2', ':', 'B6', '+', '/'],
    correctAnswer: '=SUM(B2:B6)',
    explanation: 'Use =SUM(B2:B6) to add every value from B2 through B6 cleanly without listing each cell.',
  },
  {
    id: 'h5',
    difficulty: 'hard',
    type: 'order',
    topic: 'Workflow',
    prompt: 'Put the four Excel workshop workflow stages in the correct order from start to finish.',
    items: [
      { id: 'Structure', label: 'Structure', hint: 'Organize tabs, columns (variables), and rows (observations)' },
      { id: 'Entry', label: 'Entry', hint: 'Input clean data, format as Table, apply number formats' },
      { id: 'Action', label: 'Action', hint: 'Write dynamic formulas and Big Four functions' },
      { id: 'Insight', label: 'Insight', hint: 'Spot outliers, summarize findings, and share results' },
    ],
    correctAnswer: ['Structure', 'Entry', 'Action', 'Insight'],
    explanation: 'A reliable spreadsheet follows Structure → Entry → Action → Insight so your analysis is built on clean, organized data.',
  },
  {
    id: 'h6',
    difficulty: 'hard',
    type: 'cellclick',
    topic: 'Functions',
    prompt: 'Using =MAX(B2:D4) to hunt for a data-entry outlier, click the cell containing the maximum value.',
    visual: {
      mode: 'single',
      cells: {
        A1: 'Trial', B1: 'Day 1', C1: 'Day 2', D1: 'Day 3',
        A2: 'T1', B2: '24', C2: '26', D2: '25',
        A3: 'T2', B3: '23', C3: '240', D3: '22',
        A4: 'T3', B4: '25', C4: '24', D4: '26',
      },
    },
    correctAnswer: 'C3',
    explanation: 'Cell C3 holds 240—ten times larger than the surrounding measurements—making it the MAX value and an obvious outlier to inspect.',
  },
  {
    id: 'h7',
    difficulty: 'hard',
    type: 'truefalse',
    topic: 'Formulas',
    prompt: 'True or False: In =SUM(A1:A5), the colon means Excel adds ONLY cells A1 and A5.',
    options: ['True', 'False'],
    correctAnswer: 'False',
    explanation: 'The colon means "through," so A1:A5 adds A1, A2, A3, A4, and A5. (A comma, A1,A5, would select only those two cells.)',
  },
  {
    id: 'h8',
    difficulty: 'hard',
    type: 'truefalse',
    topic: 'Functions',
    prompt: 'True or False: The =AVERAGE() function counts cells containing 0 when computing the average.',
    options: ['True', 'False'],
    correctAnswer: 'True',
    explanation: 'AVERAGE ignores blank cells and text, but 0 is a numeric value and is always included in the denominator.',
  },
  {
    id: 'h9',
    difficulty: 'hard',
    type: 'truefalse',
    topic: 'Functions',
    prompt: 'True or False: Because Excel has a =SUM() function, it also has a built-in =SUBTRACT() function.',
    options: ['True', 'False'],
    correctAnswer: 'False',
    explanation: 'Classic trap! Excel does not have a SUBTRACT function—always use the minus sign (like =B2-C2).',
  },
];

// ============================================================================
// UTILITIES
// ============================================================================

function shuffle(arr) {
  const copy = [...arr];
  for (let i = copy.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [copy[i], copy[j]] = [copy[j], copy[i]];
  }
  return copy;
}

function prepareQuestion(q) {
  if (q.type === 'mcq' && q.options) {
    return { ...q, options: shuffle(q.options) };
  }
  if (q.type === 'builder' && q.tokens) {
    return { ...q, tokens: shuffle(q.tokens) };
  }
  if (q.type === 'order' && q.items) {
    let shuffled = shuffle(q.items);
    // Avoid accidentally starting already in exact correct order
    if (shuffled.every((item, idx) => item.id === q.correctAnswer[idx])) {
      [shuffled[0], shuffled[1]] = [shuffled[1], shuffled[0]];
    }
    return { ...q, items: shuffled };
  }
  return { ...q };
}

function buildDeck(difficulty) {
  const cfg = DIFFICULTIES[difficulty];
  const tierQs = QUESTION_BANK.filter((q) => q.difficulty === difficulty);
  // Guarantee shuffled order while keeping all tier questions up to runLength
  return shuffle(tierQs).slice(0, cfg.runLength).map(prepareQuestion);
}

function getNextTier(tier, delta) {
  const order = ['easy', 'medium', 'hard'];
  const idx = order.indexOf(tier);
  const next = Math.max(0, Math.min(order.length - 1, idx + delta));
  return order[next];
}

function formatAnswerForDisplay(ans) {
  if (Array.isArray(ans)) return ans.join(', ');
  if (ans === null || ans === undefined || ans === '') return '(No answer — timed out)';
  return String(ans);
}

// Render inline formulas like =SUM(B2:B4), =A1+B1, A1:B3 as monospace cell chips
function RichText({ text }) {
  const parts = String(text).split(/(=(?:[A-Z]+)?\([A-Z0-9:,+\-*/]+\)|=[A-Z0-9+\-*/]+|\b[A-D][1-6]:[A-D][1-6]\b)/g);
  return (
    <>
      {parts.map((part, idx) => {
        if (/^=|^[A-D][1-6]:[A-D][1-6]$/.test(part)) {
          return (
            <span
              key={idx}
              className="inline-block font-mono text-[0.92em] px-1.5 py-0.5 mx-0.5 rounded bg-[#FAF7F2] border border-[#E8E1D5] text-[#1E1B16] font-semibold align-baseline"
            >
              {part}
            </span>
          );
        }
        return <React.Fragment key={idx}>{part}</React.Fragment>;
      })}
    </>
  );
}

// ============================================================================
// WEB AUDIO SYNTHESIZER (Muted by default, zero external files)
// ============================================================================

function useSoundEffects(enabled) {
  const ctxRef = useRef(null);

  const getCtx = useCallback(() => {
    if (!enabled || typeof window === 'undefined') return null;
    const AudioCtx = window.AudioContext || window.webkitAudioContext;
    if (!AudioCtx) return null;
    if (!ctxRef.current) ctxRef.current = new AudioCtx();
    if (ctxRef.current.state === 'suspended') ctxRef.current.resume().catch(() => {});
    return ctxRef.current;
  }, [enabled]);

  const playClick = useCallback(() => {
    const ctx = getCtx();
    if (!ctx) return;
    const osc = ctx.createOscillator();
    const gain = ctx.createGain();
    osc.type = 'triangle';
    osc.frequency.setValueAtTime(520, ctx.currentTime);
    osc.frequency.exponentialRampToValueAtTime(320, ctx.currentTime + 0.045);
    gain.gain.setValueAtTime(0.08, ctx.currentTime);
    gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + 0.05);
    osc.connect(gain);
    gain.connect(ctx.destination);
    osc.start();
    osc.stop(ctx.currentTime + 0.055);
  }, [getCtx]);

  const playCorrect = useCallback(() => {
    const ctx = getCtx();
    if (!ctx) return;
    [523.25, 659.25].forEach((freq, i) => {
      const osc = ctx.createOscillator();
      const gain = ctx.createGain();
      osc.type = 'sine';
      const t = ctx.currentTime + i * 0.09;
      osc.frequency.setValueAtTime(freq, t);
      gain.gain.setValueAtTime(0.11, t);
      gain.gain.exponentialRampToValueAtTime(0.001, t + 0.18);
      osc.connect(gain);
      gain.connect(ctx.destination);
      osc.start(t);
      osc.stop(t + 0.19);
    });
  }, [getCtx]);

  const playWrong = useCallback(() => {
    const ctx = getCtx();
    if (!ctx) return;
    const osc = ctx.createOscillator();
    const gain = ctx.createGain();
    osc.type = 'sawtooth';
    osc.frequency.setValueAtTime(210, ctx.currentTime);
    osc.frequency.linearRampToValueAtTime(155, ctx.currentTime + 0.16);
    gain.gain.setValueAtTime(0.07, ctx.currentTime);
    gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + 0.18);
    osc.connect(gain);
    gain.connect(ctx.destination);
    osc.start();
    osc.stop(ctx.currentTime + 0.19);
  }, [getCtx]);

  return { playClick, playCorrect, playWrong };
}

// ============================================================================
// PIXEL THE SPREADSHEET CAT — Inline Pixel-Art SVG Mascot (6 Poses)
// ============================================================================

function PixelCat({ mood = 'waving', streak = 0, urgent = false, medalTier = 'gold', size = 84 }) {
  const showShades = mood === 'celebrating' && streak >= 3;
  const medalColors = {
    gold: { main: '#FFB800', rim: '#D98A00', shine: '#FFF3B0' },
    silver: { main: '#CBD5E1', rim: '#64748B', shine: '#F8FAFC' },
    bronze: { main: '#D97745', rim: '#9A4B22', shine: '#FBD3B6' },
  }[medalTier] || { main: '#FFB800', rim: '#D98A00', shine: '#FFF3B0' };

  const bodyClass =
    mood === 'celebrating'
      ? 'cq-pixel-hop'
      : urgent && mood === 'thinking'
      ? 'cq-pixel-urgent'
      : '';

  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 32 32"
      shapeRendering="crispEdges"
      aria-hidden="true"
      className={`select-none shrink-0 ${bodyClass}`}
    >
      {/* Dramatic flopped pose (Game Over) */}
      {mood === 'dramatic' ? (
        <g>
          {/* Tail curled flat */}
          <rect x="2" y="22" width="4" height="2" fill="#FF7000" />
          <rect x="1" y="20" width="2" height="2" fill="#FF7000" />
          {/* Flopped spreadsheet body */}
          <rect x="6" y="18" width="16" height="9" fill="#FFFDF9" stroke="#1E1B16" strokeWidth="1" />
          <rect x="11" y="18" width="1" height="9" fill="#E8E1D5" />
          <rect x="16" y="18" width="1" height="9" fill="#E8E1D5" />
          <rect x="6" y="22" width="16" height="1" fill="#E8E1D5" />
          {/* Head resting on side */}
          <rect x="19" y="15" width="10" height="10" fill="#FFFDF9" stroke="#1E1B16" strokeWidth="1" />
          {/* Tangerine patch */}
          <rect x="20" y="16" width="4" height="3" fill="#FF7000" />
          {/* Left ear & green active-cell right ear */}
          <rect x="19" y="12" width="3" height="3" fill="#FF7000" stroke="#1E1B16" strokeWidth="1" />
          <rect x="26" y="12" width="3" height="3" fill="#FFFFFF" stroke="#2E9E5B" strokeWidth="1.2" />
          {/* X_X eyes + single pixel tear */}
          <rect x="21" y="19" width="2" height="1" fill="#1E1B16" />
          <rect x="25" y="19" width="2" height="1" fill="#1E1B16" />
          <rect x="26" y="21" width="2" height="2" fill="#4E8FE0" />
          {/* Tiny Retry? sign held up */}
          <rect x="9" y="6" width="11" height="7" fill="#FFFFFF" stroke="#1E1B16" strokeWidth="1" />
          <rect x="14" y="13" width="1" height="5" fill="#1E1B16" />
          <rect x="11" y="8" width="7" height="1" fill="#FF7000" />
          <rect x="11" y="10" width="5" height="1" fill="#1E1B16" />
        </g>
      ) : (
        <g>
          {/* Tail (animated 2-frame flick) */}
          <g className="cq-tail-flick" style={{ transformOrigin: '24px 23px' }}>
            <rect x="24" y="21" width="3" height="2" fill="#FF7000" />
            <rect x="26" y="18" width="2" height="4" fill="#FF7000" />
            <rect x="25" y="16" width="2" height="2" fill="#1E1B16" />
          </g>

          {/* Body (spreadsheet-cell fur pattern) */}
          <rect x="9" y="18" width="14" height="10" fill="#FFFDF9" stroke="#1E1B16" strokeWidth="1" />
          {/* Faint spreadsheet gridlines on fur */}
          <rect x="13" y="19" width="1" height="8" fill="#E8E1D5" />
          <rect x="18" y="19" width="1" height="8" fill="#E8E1D5" />
          <rect x="10" y="22" width="12" height="1" fill="#E8E1D5" />
          <rect x="10" y="25" width="12" height="1" fill="#E8E1D5" />
          {/* Tangerine patch on shoulder */}
          <rect x="10" y="19" width="3" height="3" fill="#FF7000" />
          {/* Paws */}
          <rect x="10" y="28" width="4" height="2" fill="#1E1B16" />
          <rect x="18" y="28" width="4" height="2" fill="#1E1B16" />

          {/* Head */}
          <rect x="8" y="8" width="16" height="11" fill="#FFFDF9" stroke="#1E1B16" strokeWidth="1" />
          {/* Tangerine forehead tab */}
          <rect x="9" y="9" width="5" height="3" fill="#FF7000" />

          {/* Left ear (normal) */}
          {mood === 'sympathetic' ? (
            <rect x="6" y="10" width="3" height="3" fill="#FF7000" stroke="#1E1B16" strokeWidth="1" />
          ) : (
            <rect x="8" y="4" width="4" height="4" fill="#FF7000" stroke="#1E1B16" strokeWidth="1" />
          )}

          {/* Right ear — corner spreadsheet cell with bold Excel green active-cell border! */}
          {mood === 'sympathetic' ? (
            <rect x="23" y="10" width="3" height="3" fill="#FFFFFF" stroke="#2E9E5B" strokeWidth="1.2" />
          ) : (
            <g>
              <rect x="20" y="4" width="4" height="4" fill="#FFFFFF" stroke="#2E9E5B" strokeWidth="1.2" />
              <rect x="23" y="7" width="1.5" height="1.5" fill="#2E9E5B" />
            </g>
          )}

          {/* Eyes */}
          {showShades ? (
            /* Pixel sunglasses for streak >= 3 */
            <g>
              <rect x="9" y="12" width="6" height="3" fill="#1E1B16" />
              <rect x="17" y="12" width="6" height="3" fill="#1E1B16" />
              <rect x="15" y="13" width="2" height="1" fill="#1E1B16" />
              <rect x="10" y="13" width="2" height="1" fill="#4E8FE0" />
              <rect x="18" y="13" width="2" height="1" fill="#4E8FE0" />
            </g>
          ) : mood === 'celebrating' || mood === 'cheering' ? (
            /* Happy ^ ^ eyes */
            <g>
              <rect x="11" y="12" width="2" height="1" fill="#1E1B16" />
              <rect x="10" y="13" width="1" height="1" fill="#1E1B16" />
              <rect x="13" y="13" width="1" height="1" fill="#1E1B16" />
              <rect x="19" y="12" width="2" height="1" fill="#1E1B16" />
              <rect x="18" y="13" width="1" height="1" fill="#1E1B16" />
              <rect x="21" y="13" width="1" height="1" fill="#1E1B16" />
            </g>
          ) : mood === 'sympathetic' ? (
            /* Soft worried eyes */
            <g>
              <rect x="11" y="13" width="2" height="2" fill="#1E1B16" />
              <rect x="19" y="13" width="2" height="2" fill="#1E1B16" />
              <rect x="10" y="12" width="2" height="1" fill="#4E8FE0" />
              <rect x="20" y="12" width="2" height="1" fill="#4E8FE0" />
            </g>
          ) : (
            /* Curious eyes */
            <g>
              <rect x="11" y="12" width="2" height="3" fill="#1E1B16" />
              <rect x="19" y="12" width="2" height="3" fill="#1E1B16" />
              <rect x="11" y="12" width="1" height="1" fill="#FFFFFF" />
              <rect x="19" y="12" width="1" height="1" fill="#FFFFFF" />
            </g>
          )}

          {/* Nose & Mouth */}
          <rect x="15" y="15" width="2" height="1" fill="#FF7000" />
          {mood === 'sympathetic' ? (
            <rect x="14" y="17" width="4" height="1" fill="#1E1B16" />
          ) : (
            <g>
              <rect x="14" y="16" width="1" height="1" fill="#1E1B16" />
              <rect x="17" y="16" width="1" height="1" fill="#1E1B16" />
              <rect x="15" y="17" width="2" height="1" fill="#1E1B16" />
            </g>
          )}

          {/* Pose-specific limbs & props */}
          {mood === 'waving' && (
            <g>
              <rect x="5" y="15" width="3" height="5" fill="#FFFDF9" stroke="#1E1B16" strokeWidth="1" />
              <rect x="5" y="14" width="3" height="2" fill="#FF7000" />
            </g>
          )}

          {mood === 'celebrating' && (
            <g>
              {/* Both paws raised */}
              <rect x="5" y="13" width="3" height="5" fill="#FFFDF9" stroke="#1E1B16" strokeWidth="1" />
              <rect x="24" y="13" width="3" height="5" fill="#FFFDF9" stroke="#1E1B16" strokeWidth="1" />
              {/* Pixel sparkles */}
              <rect x="3" y="6" width="2" height="2" fill="#FF7000" />
              <rect x="27" y="5" width="2" height="2" fill="#2E9E5B" />
              <rect x="15" y="1" width="2" height="2" fill="#4E8FE0" />
            </g>
          )}

          {mood === 'cheering' && (
            <g>
              {/* Raised paw holding pixel medal */}
              <rect x="4" y="14" width="4" height="5" fill="#FFFDF9" stroke="#1E1B16" strokeWidth="1" />
              <rect x="3" y="4" width="2" height="4" fill="#FF4F58" />
              <rect x="5" y="4" width="2" height="4" fill="#4E8FE0" />
              <rect x="2" y="7" width="6" height="6" fill={medalColors.main} stroke={medalColors.rim} strokeWidth="1" />
              <rect x="4" y="9" width="2" height="2" fill={medalColors.shine} />
            </g>
          )}

          {mood === 'thinking' && urgent && (
            /* Nervous exclamation mark when timer <= 5s */
            <g>
              <rect x="27" y="4" width="2" height="5" fill="#FF4F58" />
              <rect x="27" y="10" width="2" height="2" fill="#FF4F58" />
            </g>
          )}
        </g>
      )}
    </svg>
  );
}

// ============================================================================
// PIXEL ICONS (Hearts, Flame Streak, Paw, Medal, Speaker)
// ============================================================================

function PixelHeart({ alive }) {
  return (
    <svg
      width="22"
      height="20"
      viewBox="0 0 11 10"
      shapeRendering="crispEdges"
      aria-hidden="true"
      className={`transition-transform duration-200 ${alive ? 'scale-100' : 'scale-90 opacity-40'}`}
    >
      <path
        d="M2 0h2v1h1v1h1V1h1V0h2v1h1v1h1v3h-1v1H9v1H8v1H7v1H6v1H5V9H4V8H3V7H2V6H1V5H0V2h1V1h1V0z"
        fill={alive ? '#FF4F58' : '#C5BDB0'}
      />
      {alive && <rect x="2" y="2" width="1" height="1" fill="#FFD1D4" />}
    </svg>
  );
}

function PixelFlame({ streak }) {
  const hot = streak >= 3;
  const scale = streak === 0 ? 0.9 : Math.min(1.25, 1 + streak * 0.06);
  return (
    <svg
      width="22"
      height="24"
      viewBox="0 0 12 14"
      shapeRendering="crispEdges"
      aria-hidden="true"
      style={{ transform: `scale(${scale})` }}
      className="transition-transform duration-200"
    >
      <path
        d="M5 0h2v2h1v2h2v2h1v4h-1v2H9v1H3v-1H2v-2H1V6h1V4h1V3h1V2h1V0z"
        fill={streak === 0 ? '#D8D1C5' : hot ? '#FF7000' : '#2E9E5B'}
      />
      {streak > 0 && (
        <path
          d="M5 5h2v2h1v3H7v1H5v-1H4V7h1V5z"
          fill={hot ? '#FFD166' : '#8CE0AB'}
        />
      )}
      {hot && (
        /* Cute pixel face at streak >= 3 */
        <g>
          <rect x="4" y="8" width="1" height="1" fill="#1E1B16" />
          <rect x="7" y="8" width="1" height="1" fill="#1E1B16" />
          <rect x="5" y="10" width="2" height="1" fill="#1E1B16" />
        </g>
      )}
    </svg>
  );
}

function ResultMedalSVG({ tier }) {
  const palette = {
    gold: { ribbonA: '#FF7000', ribbonB: '#2E9E5B', main: '#FFB800', rim: '#C97B00', inner: '#FFE58F', label: 'GOLD' },
    silver: { ribbonA: '#4E8FE0', ribbonB: '#2E9E5B', main: '#CBD5E1', rim: '#64748B', inner: '#F1F5F9', label: 'SILVER' },
    bronze: { ribbonA: '#FF7000', ribbonB: '#4E8FE0', main: '#D97745', rim: '#9A4B22', inner: '#FAD2B8', label: 'BRONZE' },
  }[tier];

  return (
    <svg width="88" height="96" viewBox="0 0 88 96" aria-label={`${palette.label} medal`} role="img">
      {/* Ribbons */}
      <polygon points="24,4 42,4 46,38 30,38" fill={palette.ribbonA} stroke="#1E1B16" strokeWidth="2.5" />
      <polygon points="46,4 64,4 58,38 42,38" fill={palette.ribbonB} stroke="#1E1B16" strokeWidth="2.5" />
      {/* Medal disc */}
      <circle cx="44" cy="60" r="28" fill={palette.main} stroke="#1E1B16" strokeWidth="3" />
      <circle cx="44" cy="60" r="21" fill={palette.inner} stroke={palette.rim} strokeWidth="2" />
      {/* Mini spreadsheet grid emblem inside medal */}
      <rect x="33" y="49" width="22" height="22" rx="2" fill="#FFFFFF" stroke="#1E1B16" strokeWidth="2" />
      <line x1="40" y1="49" x2="40" y2="71" stroke="#E8E1D5" strokeWidth="1.5" />
      <line x1="48" y1="49" x2="48" y2="71" stroke="#E8E1D5" strokeWidth="1.5" />
      <line x1="33" y1="56" x2="55" y2="56" stroke="#E8E1D5" strokeWidth="1.5" />
      <line x1="33" y1="63" x2="55" y2="63" stroke="#E8E1D5" strokeWidth="1.5" />
      <rect x="40" y="56" width="8" height="7" fill="#2E9E5B" fillOpacity="0.25" stroke="#2E9E5B" strokeWidth="2" />
    </svg>
  );
}

// ============================================================================
// MINI SPREADSHEET COMPONENT (Static Visual + Interactive Cell/Range Click)
// ============================================================================

const COLS = ['A', 'B', 'C', 'D'];
const ROWS = [1, 2, 3, 4];

function MiniGrid({
  visual,
  interactive = false,
  selectedCells = [],
  onToggleCell,
  disabled = false,
  correctAnswer = null,
  showResult = false,
}) {
  const [hovered, setHovered] = useState(null);
  const mode = visual?.mode || 'static';
  const cells = visual?.cells || {};
  const shadedSet = new Set(visual?.shadedRange || []);
  const selectedSet = new Set(selectedCells);
  const correctSet = new Set(Array.isArray(correctAnswer) ? correctAnswer : correctAnswer ? [correctAnswer] : []);

  const displayNameBox =
    hovered ||
    (selectedCells.length === 1
      ? selectedCells[0]
      : selectedCells.length > 1
      ? `${selectedCells.length} cells`
      : visual?.nameBox || visual?.activeCell || 'A1');

  const displayFormulaBar =
    visual?.formulaBar ??
    (hovered ? cells[hovered] || '' : selectedCells.length === 1 ? cells[selectedCells[0]] || '' : '');

  return (
    <div className="rounded-xl border-2 border-[#E8E1D5] bg-white overflow-hidden shadow-xs">
      {/* Fake Excel Ribbon Sliver + Name Box + Formula Bar */}
      <div className="flex items-center gap-2 px-3 py-1.5 bg-[#FAF7F2] border-b border-[#E8E1D5] text-xs font-mono">
        <div className="px-2 py-0.5 bg-white border border-[#E8E1D5] rounded text-[#1E1B16] font-bold min-w-[54px] text-center">
          {displayNameBox}
        </div>
        <span className="text-[#8C8477] italic font-serif font-bold select-none">fx</span>
        <div className="flex-1 px-2 py-0.5 bg-white border border-[#E8E1D5] rounded text-[#1E1B16] truncate min-h-[24px]">
          {displayFormulaBar || <span className="text-[#B5AEA2]">—</span>}
        </div>
      </div>

      {/* 4x4 Spreadsheet Grid */}
      <div className="overflow-x-auto">
        <table className="w-full border-collapse table-fixed select-none" role="grid" aria-label="Mini spreadsheet grid">
          <thead>
            <tr className="bg-[#F4EFE6]">
              <th className="w-10 border-b border-r border-[#E8E1D5] py-1.5 text-xs font-mono text-[#8C8477]" aria-label="Corner" />
              {COLS.map((c) => (
                <th
                  key={c}
                  scope="col"
                  className="border-b border-r last:border-r-0 border-[#E8E1D5] py-1.5 text-xs font-mono font-bold text-[#5E574D] text-center"
                >
                  {c}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {ROWS.map((r) => (
              <tr key={r}>
                <th
                  scope="row"
                  className="bg-[#F4EFE6] border-b last:border-b-0 border-r border-[#E8E1D5] py-2 text-xs font-mono font-bold text-[#5E574D] text-center"
                >
                  {r}
                </th>
                {COLS.map((c) => {
                  const coord = `${c}${r}`;
                  const val = cells[coord] ?? '';
                  const isStaticActive = visual?.activeCell === coord;
                  const isShaded = shadedSet.has(coord);
                  const isSelected = selectedSet.has(coord);
                  const isCorrectCell = showResult && correctSet.has(coord);
                  const isWrongPick = showResult && isSelected && !correctSet.has(coord);

                  let cellClasses =
                    'relative w-full min-h-[44px] px-2 py-2 text-center font-mono text-sm transition-colors duration-150 ';

                  if (isWrongPick) {
                    cellClasses += 'bg-[#FFE4E6] text-[#1E1B16] font-bold ring-2 ring-inset ring-[#FF4F58]';
                  } else if (isCorrectCell) {
                    cellClasses += 'bg-[#D6F5E3] text-[#1E1B16] font-bold ring-2 ring-inset ring-[#2E9E5B]';
                  } else if (isSelected) {
                    cellClasses += 'bg-[#D6F5E3] text-[#1E1B16] font-bold ring-2 ring-inset ring-[#2E9E5B]';
                  } else if (isStaticActive) {
                    cellClasses += 'bg-[#E9F9EF] text-[#1E1B16] font-bold ring-[3px] ring-inset ring-[#2E9E5B]';
                  } else if (isShaded) {
                    cellClasses += 'bg-[#E6F1FC] text-[#1E1B16] font-semibold ring-1 ring-inset ring-[#4E8FE0]/50';
                  } else {
                    cellClasses += interactive && !disabled
                      ? 'bg-white text-[#1E1B16] hover:bg-[#FFF4E8]'
                      : 'bg-white text-[#1E1B16]';
                  }

                  return (
                    <td
                      key={coord}
                      className="border-b last:border-b-0 border-r last:border-r-0 border-[#E8E1D5] p-0"
                    >
                      {interactive ? (
                        <button
                          type="button"
                          disabled={disabled}
                          onClick={() => onToggleCell && onToggleCell(coord)}
                          onMouseEnter={() => setHovered(coord)}
                          onMouseLeave={() => setHovered(null)}
                          aria-label={`Cell ${coord}${val ? `, value ${val}` : ''}${isSelected ? ', selected' : ''}`}
                          aria-pressed={isSelected}
                          className={`${cellClasses} cursor-pointer disabled:cursor-default`}
                        >
                          <span>{val}</span>
                          {isSelected && !isWrongPick && (
                            <span className="absolute bottom-0.5 right-0.5 w-2 h-2 bg-[#2E9E5B]" aria-hidden="true" />
                          )}
                          {isWrongPick && (
                            <span className="absolute top-0.5 right-1 text-[11px] text-[#FF4F58] font-bold" aria-hidden="true">✗</span>
                          )}
                          {isCorrectCell && (
                            <span className="absolute top-0.5 right-1 text-[11px] text-[#2E9E5B] font-bold" aria-hidden="true">✓</span>
                          )}
                        </button>
                      ) : (
                        <div className={cellClasses}>
                          <span>{val}</span>
                          {isStaticActive && (
                            <span className="absolute bottom-0 right-0 w-2.5 h-2.5 bg-[#2E9E5B] border border-white" aria-hidden="true" />
                          )}
                        </div>
                      )}
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {interactive && mode === 'range' && !showResult && (
        <div className="px-3 py-1.5 bg-[#FAF7F2] border-t border-[#E8E1D5] flex items-center justify-between text-xs text-[#5E574D]">
          <span>Click cells to toggle them in your range.</span>
          <span className="font-mono font-semibold text-[#1E1B16]">
            {selectedCells.length ? [...selectedCells].sort().join(', ') : 'None selected'}
          </span>
        </div>
      )}
    </div>
  );
}

// ============================================================================
// FORMULA BUILDER COMPONENT
// ============================================================================

function FormulaBuilder({ question, disabled, onSubmit, onTokenClick }) {
  const [tray, setTray] = useState([]);

  useEffect(() => {
    setTray([]);
  }, [question.id]);

  const addToken = (tok, idx) => {
    if (disabled) return;
    onTokenClick();
    setTray((prev) => [...prev, { tok, uid: `${idx}-${Date.now()}-${Math.random()}` }]);
  };

  const removeAt = (pos) => {
    if (disabled) return;
    onTokenClick();
    setTray((prev) => prev.filter((_, i) => i !== pos));
  };

  const clearTray = () => {
    if (disabled) return;
    onTokenClick();
    setTray([]);
  };

  const assembled = tray.map((t) => t.tok).join('');

  return (
    <div className="space-y-4">
      {/* Goal sign */}
      {question.goalLabel && (
        <div className="inline-flex items-center gap-2 px-3 py-1.5 rounded-lg bg-[#EBF3FC] border border-[#4E8FE0]/40 text-sm text-[#1E1B16]">
          <span className="font-mono text-xs font-bold uppercase tracking-wide text-[#2463B5]">GOAL</span>
          <span className="font-semibold">{question.goalLabel}</span>
        </div>
      )}

      {/* Token palette (cell tiles with faint grid texture) */}
      <div>
        <p className="text-xs uppercase tracking-wider font-semibold text-[#6E675C] mb-2">
          Click tokens to assemble your formula:
        </p>
        <div className="flex flex-wrap gap-2">
          {question.tokens.map((tok, idx) => (
            <button
              key={`${tok}-${idx}`}
              type="button"
              disabled={disabled}
              onClick={() => addToken(tok, idx)}
              className="min-w-[48px] min-h-[44px] px-3.5 py-2 rounded-lg font-mono text-base font-bold text-[#1E1B16] bg-white border-2 border-[#E8E1D5] shadow-[0_2px_0_#E8E1D5] hover:border-[#FF7000] hover:bg-[#FFF7EF] active:translate-y-[2px] active:shadow-none transition-all cursor-pointer disabled:opacity-50 disabled:cursor-not-allowed"
            >
              {tok}
            </button>
          ))}
        </div>
      </div>

      {/* Merged-cell formula tray */}
      <div className="rounded-xl border-2 border-[#2E9E5B] bg-[#FAF7F2] p-3">
        <div className="flex items-center gap-2 mb-1.5">
          <span className="font-serif italic font-bold text-sm text-[#2E9E5B]">fx</span>
          <span className="text-xs font-mono text-[#6E675C]">Formula tray (click any token to remove)</span>
        </div>
        <div className="min-h-[52px] rounded-lg bg-white border border-dashed border-[#D3CBBF] px-2.5 py-2 flex flex-wrap items-center gap-1.5">
          {tray.length === 0 ? (
            <span className="font-mono text-sm text-[#A69E92]">Your formula will appear here…</span>
          ) : (
            tray.map((item, pos) => (
              <button
                key={item.uid}
                type="button"
                disabled={disabled}
                onClick={() => removeAt(pos)}
                aria-label={`Remove ${item.tok}`}
                className="inline-flex items-center gap-1 px-2.5 py-1.5 min-h-[36px] rounded-md bg-[#E9F9EF] border border-[#2E9E5B] font-mono text-sm font-bold text-[#1E1B16] hover:bg-[#FFE4E6] hover:border-[#FF4F58] transition-colors cursor-pointer disabled:cursor-default"
              >
                <span>{item.tok}</span>
                {!disabled && <span className="text-xs text-[#6E675C]" aria-hidden="true">×</span>}
              </button>
            ))
          )}
        </div>
        {assembled && (
          <div className="mt-2 text-xs font-mono text-[#5E574D]">
            Preview: <span className="font-bold text-[#1E1B16]">{assembled}</span>
          </div>
        )}
      </div>

      {!disabled && (
        <div className="flex items-center gap-3">
          <button
            type="button"
            onClick={clearTray}
            disabled={tray.length === 0}
            className="min-h-[44px] px-4 py-2 rounded-xl border-2 border-[#E8E1D5] bg-white text-sm font-semibold text-[#5E574D] hover:border-[#1E1B16] hover:text-[#1E1B16] transition-colors cursor-pointer disabled:opacity-40 disabled:cursor-not-allowed"
          >
            Clear
          </button>
          <button
            type="button"
            onClick={() => onSubmit(assembled)}
            disabled={tray.length === 0}
            className="cq-btn-press flex-1 min-h-[46px] px-5 py-2.5 rounded-xl bg-[#FF7000] text-[#1E1B16] font-bold text-base shadow-[0_3px_0_#C25200] hover:bg-[#FF7F1A] transition-all cursor-pointer disabled:opacity-40 disabled:cursor-not-allowed"
          >
            Check Formula
          </button>
        </div>
      )}
    </div>
  );
}

// ============================================================================
// SORT / ORDER COMPONENT (Workflow Stages)
// ============================================================================

function OrderChallenge({ question, disabled, onSubmit, onMoveClick }) {
  const [list, setList] = useState(question.items || []);
  const dragIdxRef = useRef(null);

  useEffect(() => {
    setList(question.items || []);
  }, [question.id, question.items]);

  const move = (idx, dir) => {
    if (disabled) return;
    const target = idx + dir;
    if (target < 0 || target >= list.length) return;
    onMoveClick();
    setList((prev) => {
      const next = [...prev];
      [next[idx], next[target]] = [next[target], next[idx]];
      return next;
    });
  };

  const handleDrop = (targetIdx) => {
    const fromIdx = dragIdxRef.current;
    dragIdxRef.current = null;
    if (disabled || fromIdx === null || fromIdx === targetIdx) return;
    onMoveClick();
    setList((prev) => {
      const next = [...prev];
      const [moved] = next.splice(fromIdx, 1);
      next.splice(targetIdx, 0, moved);
      return next;
    });
  };

  return (
    <div className="space-y-3">
      <p className="text-xs uppercase tracking-wider font-semibold text-[#6E675C]">
        Drag rows or use the arrow buttons to arrange stages 1 → 4:
      </p>
      <div className="space-y-2">
        {list.map((item, idx) => (
          <div
            key={item.id}
            draggable={!disabled}
            onDragStart={() => {
              dragIdxRef.current = idx;
            }}
            onDragOver={(e) => e.preventDefault()}
            onDrop={() => handleDrop(idx)}
            className="flex items-center gap-3 p-3 rounded-xl bg-white border-2 border-[#E8E1D5] shadow-xs"
          >
            <span className="w-8 h-8 rounded-lg bg-[#FAF7F2] border border-[#E8E1D5] font-mono font-bold text-sm text-[#1E1B16] flex items-center justify-center shrink-0">
              {idx + 1}
            </span>
            <div className="flex-1 min-w-0">
              <div className="font-bold text-base text-[#1E1B16]">{item.label}</div>
              <div className="text-xs text-[#6E675C]">{item.hint}</div>
            </div>
            {!disabled && (
              <div className="flex items-center gap-1 shrink-0">
                <button
                  type="button"
                  onClick={() => move(idx, -1)}
                  disabled={idx === 0}
                  aria-label={`Move ${item.label} up`}
                  className="w-11 h-11 rounded-lg border border-[#E8E1D5] bg-[#FAF7F2] text-[#1E1B16] font-bold hover:border-[#FF7000] hover:bg-[#FFF4E8] disabled:opacity-30 cursor-pointer disabled:cursor-not-allowed"
                >
                  ↑
                </button>
                <button
                  type="button"
                  onClick={() => move(idx, 1)}
                  disabled={idx === list.length - 1}
                  aria-label={`Move ${item.label} down`}
                  className="w-11 h-11 rounded-lg border border-[#E8E1D5] bg-[#FAF7F2] text-[#1E1B16] font-bold hover:border-[#FF7000] hover:bg-[#FFF4E8] disabled:opacity-30 cursor-pointer disabled:cursor-not-allowed"
                >
                  ↓
                </button>
              </div>
            )}
          </div>
        ))}
      </div>
      {!disabled && (
        <button
          type="button"
          onClick={() => onSubmit(list.map((x) => x.id))}
          className="cq-btn-press w-full min-h-[46px] px-5 py-2.5 rounded-xl bg-[#FF7000] text-[#1E1B16] font-bold text-base shadow-[0_3px_0_#C25200] hover:bg-[#FF7F1A] transition-all cursor-pointer"
        >
          Check Order
        </button>
      )}
    </div>
  );
}

// ============================================================================
// CONFETTI BURST (Pixel squares in palette colors, >= 70% score)
// ============================================================================

function PixelConfetti() {
  const pieces = useMemo(() => {
    const colors = ['#FF7000', '#2E9E5B', '#4E8FE0', '#FF4F58', '#FFB800'];
    return Array.from({ length: 42 }, (_, i) => ({
      id: i,
      left: `${(i * 2.4 + (i % 7) * 3) % 98}%`,
      delay: `${(i % 10) * 0.14}s`,
      duration: `${2.2 + (i % 5) * 0.45}s`,
      size: 7 + (i % 3) * 3,
      color: colors[i % colors.length],
    }));
  }, []);

  return (
    <div className="pointer-events-none fixed inset-0 overflow-hidden z-40" aria-hidden="true">
      {pieces.map((p) => (
        <span
          key={p.id}
          className="cq-confetti-piece absolute top-[-16px]"
          style={{
            left: p.left,
            width: `${p.size}px`,
            height: `${p.size}px`,
            backgroundColor: p.color,
            animationDelay: p.delay,
            animationDuration: p.duration,
          }}
        />
      ))}
    </div>
  );
}

// ============================================================================
// GAME STATE REDUCER
// ============================================================================

function createInitialState() {
  return {
    screen: 'start', // 'start' | 'quiz' | 'gameover' | 'results'
    difficulty: 'easy',
    adaptive: false,
    currentTier: 'easy',
    consecutiveCorrect: 0,
    consecutiveWrong: 0,
    deck: [],
    askedIds: [],
    index: 0,
    score: 0,
    streak: 0,
    bestStreak: 0,
    lives: Infinity,
    timeLeft: 45,
    answered: null, // { isCorrect, timedOut, userAnswer, pointsEarned, streakBonus, flavorText }
    answers: [],
    topicStats: {},
    scorePopup: null,
  };
}

function evaluateAnswer(question, userAnswer) {
  if (userAnswer === null || userAnswer === undefined) return false;
  if (question.type === 'cellclick') {
    if (Array.isArray(question.correctAnswer)) {
      if (!Array.isArray(userAnswer)) return false;
      const a = [...question.correctAnswer].sort().join(',');
      const b = [...userAnswer].sort().join(',');
      return a === b;
    }
    return String(userAnswer).toUpperCase() === String(question.correctAnswer).toUpperCase();
  }
  if (question.type === 'builder') {
    const clean = (s) => String(s).replace(/\s+/g, '').toUpperCase();
    return clean(userAnswer) === clean(question.correctAnswer);
  }
  if (question.type === 'order') {
    if (!Array.isArray(userAnswer) || !Array.isArray(question.correctAnswer)) return false;
    return (
      userAnswer.length === question.correctAnswer.length &&
      userAnswer.every((v, i) => v === question.correctAnswer[i])
    );
  }
  return String(userAnswer) === String(question.correctAnswer);
}

function gameReducer(state, action) {
  switch (action.type) {
    case 'SELECT_DIFFICULTY':
      return { ...state, difficulty: action.difficulty };

    case 'TOGGLE_ADAPTIVE':
      return { ...state, adaptive: !state.adaptive };

    case 'START_GAME': {
      const diff = action.difficulty || state.difficulty;
      const cfg = DIFFICULTIES[diff];
      const deck = buildDeck(diff);
      return {
        ...createInitialState(),
        screen: 'quiz',
        difficulty: diff,
        adaptive: action.adaptive !== undefined ? action.adaptive : state.adaptive,
        currentTier: diff,
        deck,
        askedIds: deck.length ? [deck[0].id] : [],
        index: 0,
        lives: cfg.lives,
        timeLeft: cfg.timePerQuestion,
      };
    }

    case 'TICK': {
      if (state.screen !== 'quiz' || state.answered) return state;
      if (state.timeLeft <= 1) {
        // Trigger timeout as wrong answer
        return gameReducer(
          { ...state, timeLeft: 0 },
          { type: 'SUBMIT_ANSWER', answer: null, timedOut: true }
        );
      }
      return { ...state, timeLeft: Math.max(0, state.timeLeft - 1) };
    }

    case 'SUBMIT_ANSWER': {
      if (state.screen !== 'quiz' || state.answered) return state;
      const q = state.deck[state.index];
      const cfg = DIFFICULTIES[state.difficulty];
      const timedOut = Boolean(action.timedOut);
      const isCorrect = !timedOut && evaluateAnswer(q, action.answer);

      const newStreak = isCorrect ? state.streak + 1 : 0;
      const bestStreak = Math.max(state.bestStreak, newStreak);
      const streakBonus = isCorrect && newStreak > 1 ? cfg.streakBonus * (newStreak - 1) : 0;
      const pointsEarned = isCorrect ? cfg.basePoints + streakBonus : 0;
      const newLives = isCorrect || !Number.isFinite(state.lives) ? state.lives : Math.max(0, state.lives - 1);

      // Adaptive difficulty tracking
      let nextTier = state.currentTier;
      let consecutiveCorrect = isCorrect ? state.consecutiveCorrect + 1 : 0;
      let consecutiveWrong = !isCorrect ? state.consecutiveWrong + 1 : 0;
      if (state.adaptive) {
        if (consecutiveCorrect >= 2) {
          nextTier = getNextTier(state.currentTier, 1);
          consecutiveCorrect = 0;
        } else if (consecutiveWrong >= 2) {
          nextTier = getNextTier(state.currentTier, -1);
          consecutiveWrong = 0;
        }
      }

      const prevTopic = state.topicStats[q.topic] || { correct: 0, total: 0 };
      const topicStats = {
        ...state.topicStats,
        [q.topic]: {
          correct: prevTopic.correct + (isCorrect ? 1 : 0),
          total: prevTopic.total + 1,
        },
      };

      const flavorPool = isCorrect ? CORRECT_FLAVORS : WRONG_FLAVORS;
      const flavorText = timedOut
        ? 'Time’s up! That timer moves fast.'
        : flavorPool[(state.index + state.score) % flavorPool.length];

      const reviewEntry = {
        question: q,
        userAnswer: action.answer,
        isCorrect,
        timedOut,
        pointsEarned,
      };

      return {
        ...state,
        score: state.score + pointsEarned,
        streak: newStreak,
        bestStreak,
        lives: newLives,
        currentTier: nextTier,
        consecutiveCorrect,
        consecutiveWrong,
        topicStats,
        answered: {
          isCorrect,
          timedOut,
          userAnswer: action.answer,
          pointsEarned,
          streakBonus,
          flavorText,
        },
        answers: [...state.answers, reviewEntry],
        scorePopup: isCorrect ? { id: Date.now(), points: pointsEarned, bonus: streakBonus } : null,
      };
    }

    case 'NEXT_QUESTION': {
      if (state.screen !== 'quiz' || !state.answered) return state;
      const cfg = DIFFICULTIES[state.difficulty];
      if (Number.isFinite(state.lives) && state.lives <= 0) {
        return { ...state, screen: 'gameover', answered: null, scorePopup: null };
      }

      const nextIndex = state.index + 1;
      if (nextIndex >= state.deck.length) {
        return { ...state, screen: 'results', answered: null, scorePopup: null };
      }

      // If adaptive mode shifted tier, try swapping the upcoming slot with an unseen question from currentTier
      let nextDeck = state.deck;
      let nextAskedIds = [...state.askedIds];
      if (state.adaptive && state.deck[nextIndex].difficulty !== state.currentTier) {
        const usedIds = new Set([...state.askedIds, ...state.deck.slice(0, nextIndex).map((x) => x.id)]);
        const candidates = QUESTION_BANK.filter(
          (q) => q.difficulty === state.currentTier && !usedIds.has(q.id)
        );
        if (candidates.length > 0) {
          const picked = prepareQuestion(candidates[Math.floor(Math.random() * candidates.length)]);
          nextDeck = [...state.deck];
          nextDeck[nextIndex] = picked;
        }
      }
      nextAskedIds.push(nextDeck[nextIndex].id);

      return {
        ...state,
        deck: nextDeck,
        askedIds: nextAskedIds,
        index: nextIndex,
        timeLeft: cfg.timePerQuestion,
        answered: null,
        scorePopup: null,
      };
    }

    case 'GO_HOME':
      return { ...createInitialState(), difficulty: state.difficulty, adaptive: state.adaptive };

    default:
      return state;
  }
}

// ============================================================================
// MAIN APPLICATION COMPONENT
// ============================================================================

export default function CellQuest() {
  const [state, dispatch] = useReducer(gameReducer, undefined, createInitialState);
  const [soundEnabled, setSoundEnabled] = useState(false);
  const [selectedCells, setSelectedCells] = useState([]);
  const [expandedReview, setExpandedReview] = useState({});
  const [displayScore, setDisplayScore] = useState(0);
  const [showGridWave, setShowGridWave] = useState(false);

  const { playClick, playCorrect, playWrong } = useSoundEffects(soundEnabled);

  const currentQuestion = state.deck[state.index] || null;
  const diffConfig = DIFFICULTIES[state.difficulty];

  // Reset cell selection whenever question changes
  useEffect(() => {
    setSelectedCells([]);
  }, [state.index, state.screen]);

  // 1-second countdown timer during active question
  useEffect(() => {
    if (state.screen !== 'quiz' || state.answered) return undefined;
    const id = window.setInterval(() => {
      dispatch({ type: 'TICK' });
    }, 1000);
    return () => window.clearInterval(id);
  }, [state.screen, state.answered, state.index]);

  // Play sound on answer resolution
  const lastAnswerCount = useRef(0);
  useEffect(() => {
    if (state.answers.length > lastAnswerCount.current) {
      const latest = state.answers[state.answers.length - 1];
      if (latest.isCorrect) playCorrect();
      else playWrong();
    }
    lastAnswerCount.current = state.answers.length;
  }, [state.answers, playCorrect, playWrong]);

  // Results screen animated score count-up + signature grid wave
  useEffect(() => {
    if (state.screen !== 'results') {
      setDisplayScore(0);
      setShowGridWave(false);
      return undefined;
    }
    setShowGridWave(true);
    const reduced =
      typeof window !== 'undefined' &&
      window.matchMedia &&
      window.matchMedia('(prefers-reduced-motion: reduce)').matches;
    if (reduced || state.score === 0) {
      setDisplayScore(state.score);
      return undefined;
    }
    let frame;
    const start = performance.now();
    const duration = 700;
    const step = (now) => {
      const t = Math.min(1, (now - start) / duration);
      const eased = 1 - Math.pow(1 - t, 3);
      setDisplayScore(Math.round(state.score * eased));
      if (t < 1) frame = requestAnimationFrame(step);
    };
    frame = requestAnimationFrame(step);
    return () => cancelAnimationFrame(frame);
  }, [state.screen, state.score]);

  // Keyboard shortcuts: 1-4 / A-D to answer MCQ/TF, Enter for Next
  useEffect(() => {
    const onKeyDown = (e) => {
      if (state.screen !== 'quiz' || !currentQuestion) return;
      if (e.target && ['INPUT', 'TEXTAREA'].includes(e.target.tagName)) return;

      if (state.answered) {
        if (e.key === 'Enter' || e.key === ' ') {
          e.preventDefault();
          playClick();
          dispatch({ type: 'NEXT_QUESTION' });
        }
        return;
      }

      if (currentQuestion.type === 'mcq' && currentQuestion.options) {
        const keyMap = { '1': 0, '2': 1, '3': 2, '4': 3, a: 0, b: 1, c: 2, d: 3 };
        const idx = keyMap[e.key.toLowerCase()];
        if (idx !== undefined && currentQuestion.options[idx] !== undefined) {
          e.preventDefault();
          dispatch({ type: 'SUBMIT_ANSWER', answer: currentQuestion.options[idx] });
        }
      } else if (currentQuestion.type === 'truefalse') {
        const k = e.key.toLowerCase();
        if (k === '1' || k === 't') {
          e.preventDefault();
          dispatch({ type: 'SUBMIT_ANSWER', answer: 'True' });
        } else if (k === '2' || k === 'f') {
          e.preventDefault();
          dispatch({ type: 'SUBMIT_ANSWER', answer: 'False' });
        }
      }
    };
    window.addEventListener('keydown', onKeyDown);
    return () => window.removeEventListener('keydown', onKeyDown);
  }, [state.screen, state.answered, currentQuestion, playClick]);

  const handleToggleCell = (coord) => {
    if (!currentQuestion || state.answered) return;
    playClick();
    if (currentQuestion.visual?.mode === 'single') {
      setSelectedCells([coord]);
      dispatch({ type: 'SUBMIT_ANSWER', answer: coord });
    } else {
      setSelectedCells((prev) =>
        prev.includes(coord) ? prev.filter((c) => c !== coord) : [...prev, coord]
      );
    }
  };

  // Derived results metrics
  const totalAnswered = state.answers.length;
  const totalCorrect = state.answers.filter((a) => a.isCorrect).length;
  const percentage = totalAnswered > 0 ? Math.round((totalCorrect / totalAnswered) * 100) : 0;
  const medalTier = percentage >= 85 ? 'gold' : percentage >= 60 ? 'silver' : 'bronze';
  const medalTitle =
    medalTier === 'gold'
      ? 'Gold Spreadsheet Master'
      : medalTier === 'silver'
      ? 'Silver Grid Navigator'
      : 'Bronze Cell Explorer';

  // Timer bar calculations
  const timerRatio = diffConfig.timePerQuestion > 0 ? state.timeLeft / diffConfig.timePerQuestion : 0;
  const timerSegments = 12;
  const filledSegments = Math.ceil(timerRatio * timerSegments);
  const timerColor =
    state.timeLeft <= 5 ? '#FF4F58' : timerRatio <= 0.45 ? '#F59E0B' : '#2E9E5B';

  // Mascot mood for Quiz screen
  const quizMood = state.answered
    ? state.answered.isCorrect
      ? 'celebrating'
      : 'sympathetic'
    : 'thinking';

  const mascotSpeech = state.answered
    ? state.answered.flavorText
    : state.timeLeft <= 5
    ? 'Hurry — timer’s almost out!'
    : state.streak >= 3
    ? `${state.streak} in a row! Shades on!`
    : currentQuestion?.type === 'builder'
    ? `Make: ${currentQuestion.goalLabel}`
    : currentQuestion?.type === 'cellclick'
    ? 'Click directly on the grid!'
    : 'Hmm, check the coordinates…';

  return (
    <div className="cq-root min-h-screen text-[#1E1B16] relative overflow-x-hidden">
      <style>{`
        .cq-root {
          background-color: #FAF7F2;
          background-image:
            linear-gradient(to right, rgba(30, 27, 22, 0.04) 1px, transparent 1px),
            linear-gradient(to bottom, rgba(30, 27, 22, 0.04) 1px, transparent 1px);
          background-size: 36px 36px;
          font-family: ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
          font-size: 16px;
          line-height: 1.5;
        }
        .cq-display {
          font-family: ui-rounded, "SF Pro Rounded", "Trebuchet MS", ui-sans-serif, system-ui, sans-serif;
          letter-spacing: -0.02em;
        }
        .cq-root button:focus-visible {
          outline: 3px solid #FF7000;
          outline-offset: 2px;
        }
        .cq-btn-press:active:not(:disabled) {
          transform: translateY(2px);
          box-shadow: 0 1px 0 #C25200 !important;
        }
        @keyframes cqCardEnter {
          from { opacity: 0; transform: translateY(10px); }
          to { opacity: 1; transform: translateY(0); }
        }
        .cq-card-enter {
          animation: cqCardEnter 220ms cubic-bezier(0.22, 1, 0.36, 1) both;
        }
        @keyframes cqShake {
          0%, 100% { transform: translateX(0); }
          20% { transform: translateX(-6px); }
          40% { transform: translateX(6px); }
          60% { transform: translateX(-4px); }
          80% { transform: translateX(4px); }
        }
        .cq-shake {
          animation: cqShake 320ms ease-out;
        }
        @keyframes cqPulseGreen {
          0% { box-shadow: 0 0 0 0 rgba(46, 158, 91, 0.45); }
          100% { box-shadow: 0 0 0 14px rgba(46, 158, 91, 0); }
        }
        .cq-pulse-green {
          animation: cqPulseGreen 450ms ease-out;
        }
        @keyframes cqFloatScore {
          0% { opacity: 0; transform: translateY(6px) scale(0.9); }
          20% { opacity: 1; transform: translateY(-4px) scale(1.06); }
          100% { opacity: 0; transform: translateY(-34px) scale(1); }
        }
        .cq-float-score {
          animation: cqFloatScore 950ms ease-out forwards;
        }
        @keyframes cqTailFlick {
          0%, 45% { transform: rotate(0deg); }
          50%, 95% { transform: rotate(9deg); }
          100% { transform: rotate(0deg); }
        }
        .cq-tail-flick {
          animation: cqTailFlick 2s steps(2, end) infinite;
        }
        @keyframes cqPixelHop {
          0%, 100% { transform: translateY(0); }
          40% { transform: translateY(-7px); }
        }
        .cq-pixel-hop {
          animation: cqPixelHop 360ms ease-out 2;
        }
        @keyframes cqUrgent {
          0%, 100% { transform: translateY(0); }
          50% { transform: translateY(-3px); }
        }
        .cq-pixel-urgent {
          animation: cqUrgent 260ms ease-in-out infinite;
        }
        @keyframes cqConfettiFall {
          0% { transform: translateY(-20px) rotate(0deg); opacity: 1; }
          85% { opacity: 1; }
          100% { transform: translateY(105vh) rotate(520deg); opacity: 0; }
        }
        .cq-confetti-piece {
          animation-name: cqConfettiFall;
          animation-timing-function: linear;
          animation-iteration-count: 2;
        }
        @keyframes cqDotPulse {
          0%, 100% { opacity: 0.25; transform: scale(0.85); }
          50% { opacity: 1; transform: scale(1.1); }
        }
        .cq-dot {
          animation: cqDotPulse 1.8s ease-in-out infinite;
        }
        @keyframes cqGridWave {
          0% { opacity: 0; transform: scale(0.8); }
          35% { opacity: 0.55; transform: scale(1); }
          100% { opacity: 0; transform: scale(1); }
        }
        .cq-wave-cell {
          animation: cqGridWave 900ms ease-out forwards;
        }
        @media (prefers-reduced-motion: reduce) {
          .cq-card-enter, .cq-shake, .cq-pulse-green, .cq-float-score,
          .cq-tail-flick, .cq-pixel-hop, .cq-pixel-urgent, .cq-confetti-piece,
          .cq-dot, .cq-wave-cell {
            animation: none !important;
            transition: none !important;
          }
        }
      `}</style>

      {/* Signature completion background grid wave */}
      {showGridWave && state.screen === 'results' && (
        <div
          className="pointer-events-none fixed inset-0 grid grid-cols-10 grid-rows-8 gap-1 p-2 z-0"
          aria-hidden="true"
        >
          {Array.from({ length: 80 }, (_, i) => {
            const col = i % 10;
            const row = Math.floor(i / 10);
            return (
              <div
                key={i}
                className="cq-wave-cell rounded-xs bg-[#2E9E5B]/25 border border-[#2E9E5B]/30 opacity-0"
                style={{ animationDelay: `${(col + row) * 45}ms` }}
              />
            );
          })}
        </div>
      )}

      {/* Main Centered Container */}
      <div className="relative z-10 max-w-3xl mx-auto px-4 py-6 sm:py-10">
        {/* Global utility bar (Sound toggle + Home link when in game) */}
        <div className="flex items-center justify-between mb-4">
          {state.screen !== 'start' ? (
            <button
              type="button"
              onClick={() => {
                playClick();
                dispatch({ type: 'GO_HOME' });
              }}
              className="min-h-[44px] px-3 py-1.5 rounded-lg text-sm font-semibold text-[#5E574D] hover:text-[#1E1B16] hover:bg-white/80 border border-transparent hover:border-[#E8E1D5] transition-colors cursor-pointer"
            >
              ← Cell Quest
            </button>
          ) : (
            <span className="text-xs font-mono uppercase tracking-wider text-[#8C8477]">
              Demystifying Data · Excel Workshop
            </span>
          )}

          <button
            type="button"
            onClick={() => setSoundEnabled((s) => !s)}
            aria-label={soundEnabled ? 'Mute sound effects' : 'Enable sound effects'}
            aria-pressed={soundEnabled}
            className="min-h-[44px] px-3 py-1.5 rounded-lg bg-white border border-[#E8E1D5] text-xs font-mono font-semibold text-[#5E574D] hover:border-[#1E1B16] hover:text-[#1E1B16] transition-colors cursor-pointer flex items-center gap-1.5"
          >
            <span aria-hidden="true">{soundEnabled ? '🔊' : '🔇'}</span>
            <span>{soundEnabled ? 'Sound On' : 'Sound Off'}</span>
          </button>
        </div>

        {/* ================================================================ */}
        {/* SCREEN 1: START SCREEN                                           */}
        {/* ================================================================ */}
        {state.screen === 'start' && (
          <div className="cq-card-enter space-y-6">
            {/* Hero Title Card with Pixel sitting on top edge */}
            <div className="relative pt-12">
              <div className="absolute top-0 left-1/2 -translate-x-1/2 flex items-end gap-3 z-20">
                <PixelCat mood="waving" size={88} />
                <div className="mb-4 px-3.5 py-2 rounded-xl rounded-bl-xs bg-white border-2 border-[#1E1B16] shadow-xs text-sm font-semibold text-[#1E1B16] whitespace-nowrap">
                  Ready to tame the grid?
                </div>
              </div>

              <div className="bg-white rounded-2xl border-2 border-[#E8E1D5] shadow-sm px-6 pt-12 pb-7 text-center">
                {/* Animated mini-grid icon + CELL QUEST logotype with column-bar L's */}
                <div className="inline-flex items-center gap-2 px-3 py-1 rounded-full bg-[#FAF7F2] border border-[#E8E1D5] text-xs font-mono font-semibold text-[#5E574D] mb-3">
                  <span className="inline-grid grid-cols-2 gap-0.5 w-3.5 h-3.5" aria-hidden="true">
                    <span className="bg-[#2E9E5B] rounded-[1px]" />
                    <span className="bg-[#E8E1D5] rounded-[1px]" />
                    <span className="bg-[#E8E1D5] rounded-[1px]" />
                    <span className="bg-[#FF7000] rounded-[1px]" />
                  </span>
                  <span>INTERACTIVE EXCEL QUIZ</span>
                </div>

                <h1 className="cq-display text-4xl sm:text-5xl font-black text-[#1E1B16] tracking-tight flex items-baseline justify-center gap-1.5">
                  <span>CE</span>
                  <span className="inline-flex items-end gap-1 mx-0.5" aria-label="LL">
                    <span className="inline-block w-3 sm:w-3.5 h-8 sm:h-10 bg-[#2E9E5B] border-2 border-[#1E1B16] rounded-xs" />
                    <span className="inline-block w-3 sm:w-3.5 h-10 sm:h-12 bg-[#FF7000] border-2 border-[#1E1B16] rounded-xs" />
                  </span>
                  <span className="ml-1.5">QUEST</span>
                </h1>

                <p className="mt-2.5 text-base sm:text-lg text-[#5E574D] max-w-md mx-auto">
                  A playful quiz about Excel fundamentals — workbooks, coordinates, formulas, and the Big Four functions.
                </p>
              </div>
            </div>

            {/* Difficulty Cards (mini worksheet tabs) */}
            <div>
              <h2 className="text-xs font-mono uppercase tracking-wider font-bold text-[#6E675C] mb-3">
                Choose your worksheet difficulty
              </h2>
              <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
                {Object.values(DIFFICULTIES).map((d) => {
                  const selected = state.difficulty === d.id;
                  return (
                    <button
                      key={d.id}
                      type="button"
                      onClick={() => {
                        playClick();
                        dispatch({ type: 'SELECT_DIFFICULTY', difficulty: d.id });
                      }}
                      aria-pressed={selected}
                      className={`group text-left rounded-2xl transition-all duration-150 cursor-pointer flex flex-col ${
                        selected
                          ? '-translate-y-1 shadow-md'
                          : 'hover:-translate-y-1 hover:shadow-md'
                      }`}
                    >
                      {/* Worksheet Tab */}
                      <div
                        className={`self-start px-3.5 py-1 rounded-t-xl border-2 border-b-0 font-mono text-xs font-bold transition-colors ${
                          selected
                            ? 'bg-[#FF7000] text-[#1E1B16] border-[#1E1B16]'
                            : 'bg-[#FAF7F2] text-[#5E574D] border-[#E8E1D5] group-hover:bg-[#FF7000] group-hover:text-[#1E1B16] group-hover:border-[#1E1B16]'
                        }`}
                      >
                        {d.name}
                      </div>

                      {/* Card Body */}
                      <div
                        className={`w-full flex-1 p-4 rounded-b-2xl rounded-tr-2xl bg-white border-2 transition-colors flex flex-col justify-between gap-3 ${
                          selected ? 'border-[#1E1B16]' : 'border-[#E8E1D5] group-hover:border-[#1E1B16]'
                        }`}
                      >
                        <div>
                          <div className="flex items-center justify-between text-xs font-mono font-bold text-[#5E574D] mb-2">
                            <span>⏱ {d.timePerQuestion}s</span>
                            <span>
                              {Number.isFinite(d.lives) ? `♥ × ${d.lives}` : '∞ lives'}
                            </span>
                          </div>
                          <p className="text-sm text-[#5E574D] leading-snug">{d.shortDesc}</p>
                        </div>

                        <div className="pt-2 border-t border-[#E8E1D5] flex items-center justify-between text-xs font-mono">
                          <span className="font-bold text-[#1E1B16]">+{d.basePoints} pts</span>
                          <span className="text-[#2E9E5B] font-bold">
                            {d.streakBonus > 0 ? `+${d.streakBonus} streak` : 'No penalty'}
                          </span>
                        </div>
                      </div>
                    </button>
                  );
                })}
              </div>
            </div>

            {/* Adaptive Hardness Stretch Toggle + Start Button */}
            <div className="bg-white rounded-2xl border-2 border-[#E8E1D5] p-4 sm:p-5 flex flex-col sm:flex-row items-start sm:items-center justify-between gap-4">
              <label className="flex items-start gap-3 cursor-pointer select-none">
                <input
                  type="checkbox"
                  checked={state.adaptive}
                  onChange={() => {
                    playClick();
                    dispatch({ type: 'TOGGLE_ADAPTIVE' });
                  }}
                  className="mt-1 w-5 h-5 accent-[#FF7000] rounded cursor-pointer"
                />
                <div>
                  <div className="text-sm font-bold text-[#1E1B16]">
                    Adaptive difficulty <span className="font-mono text-xs text-[#4E8FE0] ml-1">BETA</span>
                  </div>
                  <div className="text-xs text-[#6E675C]">
                    2 right in a row steps up a tier; 2 misses in a row eases back down.
                  </div>
                </div>
              </label>

              <button
                type="button"
                onClick={() => {
                  playClick();
                  dispatch({ type: 'START_GAME' });
                }}
                className="cq-btn-press w-full sm:w-auto min-h-[50px] px-8 py-3 rounded-full bg-[#FF7000] text-[#1E1B16] font-extrabold text-base shadow-[0_3px_0_#C25200] hover:bg-[#FF7F1A] transition-all cursor-pointer shrink-0"
              >
                Start Quest →
              </button>
            </div>

            {/* Bottom pixel progress path preview */}
            <div className="flex items-center justify-center gap-2 pt-1" aria-hidden="true">
              {Array.from({ length: 10 }, (_, i) => (
                <span
                  key={i}
                  className="cq-dot w-2.5 h-2.5 rounded-[2px] bg-[#2E9E5B]"
                  style={{ animationDelay: `${i * 150}ms` }}
                />
              ))}
            </div>
          </div>
        )}

        {/* ================================================================ */}
        {/* SCREEN 2: QUIZ LOOP                                              */}
        {/* ================================================================ */}
        {state.screen === 'quiz' && currentQuestion && (
          <div className="space-y-4">
            {/* Top HUD Bar: Lives | Streak | Score */}
            <div className="bg-white rounded-2xl border-2 border-[#E8E1D5] px-4 py-3 flex items-center justify-between gap-3 relative">
              {/* Lives */}
              <div className="flex items-center gap-1.5" aria-label={`Lives: ${Number.isFinite(state.lives) ? state.lives : 'Unlimited'}`}>
                {Number.isFinite(diffConfig.lives) ? (
                  Array.from({ length: diffConfig.lives }, (_, i) => (
                    <PixelHeart key={i} alive={i < state.lives} />
                  ))
                ) : (
                  <span className="inline-flex items-center gap-1.5 text-sm font-mono font-bold text-[#2E9E5B]">
                    <PixelHeart alive /> ∞
                  </span>
                )}
              </div>

              {/* Streak Meter */}
              <div
                className="flex items-center gap-1.5 px-3 py-1 rounded-full bg-[#FAF7F2] border border-[#E8E1D5]"
                aria-label={`Streak: ${state.streak}`}
              >
                <PixelFlame streak={state.streak} />
                <span className="font-mono font-bold text-sm text-[#1E1B16]">
                  {state.streak}
                </span>
                <span className="hidden sm:inline text-xs text-[#6E675C]">streak</span>
              </div>

              {/* Score + Floating Popup */}
              <div className="relative flex items-center gap-1.5" aria-label={`Score: ${state.score}`}>
                <span className="text-xs font-mono uppercase text-[#8C8477]">PTS</span>
                <span className="cq-display text-2xl font-black text-[#1E1B16] tabular-nums">
                  {state.score}
                </span>
                {state.scorePopup && (
                  <div
                    key={state.scorePopup.id}
                    className="cq-float-score pointer-events-none absolute -top-2 right-0 font-mono font-black text-base text-[#FF7000] whitespace-nowrap"
                  >
                    +{state.scorePopup.points}
                    {state.scorePopup.bonus > 0 ? ` (+${state.scorePopup.bonus}🔥)` : ''}
                  </div>
                )}
              </div>
            </div>

            {/* Shrinking-Cell Timer Bar */}
            <div className="space-y-1">
              <div className="flex items-center justify-between text-xs font-mono">
                <span className="text-[#6E675C]">
                  {diffConfig.name}
                  {state.adaptive ? ` · Adaptive (${state.currentTier.toUpperCase()})` : ''}
                </span>
                <span
                  className={`font-bold tabular-nums ${
                    state.timeLeft <= 5 ? 'text-[#FF4F58]' : 'text-[#1E1B16]'
                  }`}
                >
                  {state.timeLeft}s left
                </span>
              </div>
              <div
                className="grid grid-cols-12 gap-1"
                role="progressbar"
                aria-label="Time remaining"
                aria-valuemin={0}
                aria-valuemax={diffConfig.timePerQuestion}
                aria-valuenow={state.timeLeft}
              >
                {Array.from({ length: timerSegments }, (_, i) => (
                  <div
                    key={i}
                    className="h-2.5 rounded-[3px] transition-colors duration-200 border border-[#E8E1D5]"
                    style={{
                      backgroundColor: i < filledSegments ? timerColor : '#F1ECE2',
                    }}
                  />
                ))}
              </div>
            </div>

            {/* Progress Map: Cell-shaped trail (hidden as full cells on tiny screens -> slim bar) */}
            <div aria-label={`Question ${state.index + 1} of ${state.deck.length}`}>
              <div className="hidden sm:flex items-center gap-1.5">
                {state.deck.map((q, i) => {
                  const past = state.answers[i];
                  const isCurrent = i === state.index;
                  return (
                    <div
                      key={q.id}
                      className={`flex-1 h-8 rounded-md font-mono text-xs font-bold flex items-center justify-center transition-all ${
                        isCurrent
                          ? 'bg-[#E9F9EF] text-[#1E1B16] border-[2.5px] border-[#2E9E5B] shadow-xs'
                          : past
                          ? past.isCorrect
                            ? 'bg-[#2E9E5B] text-[#1E1B16] border border-[#2E9E5B]'
                            : 'bg-[#FFE4E6] text-[#1E1B16] border border-[#FF4F58]'
                          : 'bg-white text-[#A69E92] border border-[#E8E1D5]'
                      }`}
                    >
                      {past ? (past.isCorrect ? '✓' : '✗') : `${String.fromCharCode(65 + (i % 26))}${i + 1}`}
                    </div>
                  );
                })}
              </div>
              <div className="sm:hidden flex items-center gap-2 text-xs font-mono text-[#6E675C]">
                <span>Q{state.index + 1}/{state.deck.length}</span>
                <div className="flex-1 h-2 bg-[#E8E1D5] rounded-full overflow-hidden">
                  <div
                    className="h-full bg-[#2E9E5B] transition-all duration-300"
                    style={{ width: `${((state.index + 1) / state.deck.length) * 100}%` }}
                  />
                </div>
              </div>
            </div>

            {/* Question Card with Pixel peeking over top-left */}
            <div className="relative pt-8">
              {/* Mascot + Speech Bubble */}
              <div className="flex items-end gap-2.5 pl-3 -mb-3 relative z-20">
                <PixelCat
                  mood={quizMood}
                  streak={state.streak}
                  urgent={!state.answered && state.timeLeft <= 5}
                  size={68}
                />
                <div className="mb-3 px-3 py-1.5 rounded-xl rounded-bl-xs bg-white border-2 border-[#1E1B16] text-xs sm:text-sm font-semibold text-[#1E1B16] shadow-xs max-w-[260px] sm:max-w-sm">
                  {mascotSpeech}
                </div>
              </div>

              <div
                key={currentQuestion.id}
                className={`cq-card-enter bg-white rounded-2xl border-2 border-[#E8E1D5] shadow-sm p-5 sm:p-6 space-y-5 ${
                  state.answered
                    ? state.answered.isCorrect
                      ? 'cq-pulse-green border-[#2E9E5B]'
                      : 'cq-shake border-[#FF4F58]'
                    : ''
                }`}
              >
                {/* Topic Chip + Question Counter */}
                <div className="flex items-center justify-between gap-2">
                  <span className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-md bg-[#EBF3FC] border border-[#4E8FE0]/40 font-mono text-xs font-bold text-[#1E1B16]">
                    {TOPIC_SHORT[currentQuestion.topic] || currentQuestion.topic}
                  </span>
                  <span className="font-mono text-xs text-[#8C8477]">
                    Question {state.index + 1} of {state.deck.length}
                  </span>
                </div>

                {/* Prompt */}
                <h2 className="cq-display text-xl sm:text-2xl font-extrabold text-[#1E1B16] leading-snug">
                  <RichText text={currentQuestion.prompt} />
                </h2>

                {/* Optional Mini-Grid Visual (Static or Interactive) */}
                {currentQuestion.visual && (
                  <MiniGrid
                    visual={currentQuestion.visual}
                    interactive={currentQuestion.type === 'cellclick'}
                    selectedCells={selectedCells}
                    onToggleCell={handleToggleCell}
                    disabled={Boolean(state.answered)}
                    correctAnswer={currentQuestion.correctAnswer}
                    showResult={Boolean(state.answered)}
                  />
                )}

                {/* Answer Area by Question Type */}
                {currentQuestion.type === 'mcq' && (
                  <div className="space-y-2.5">
                    {currentQuestion.options.map((opt, idx) => {
                      const isPicked = state.answered && state.answered.userAnswer === opt;
                      const isCorrectOpt = opt === currentQuestion.correctAnswer;
                      let btnStyle =
                        'bg-white border-[#E8E1D5] text-[#1E1B16] hover:border-[#FF7000] hover:bg-[#FFF7EF]';
                      if (state.answered) {
                        if (isCorrectOpt) {
                          btnStyle = 'bg-[#D6F5E3] border-[#2E9E5B] text-[#1E1B16] font-bold';
                        } else if (isPicked) {
                          btnStyle = 'bg-[#FFE4E6] border-[#FF4F58] text-[#1E1B16]';
                        } else {
                          btnStyle = 'bg-[#FAF7F2] border-[#E8E1D5] text-[#8C8477] opacity-70';
                        }
                      }
                      return (
                        <button
                          key={opt}
                          type="button"
                          disabled={Boolean(state.answered)}
                          onClick={() => dispatch({ type: 'SUBMIT_ANSWER', answer: opt })}
                          className={`w-full min-h-[52px] px-4 py-3 rounded-xl border-2 text-left flex items-center gap-3 transition-all cursor-pointer disabled:cursor-default ${btnStyle}`}
                        >
                          <span className="w-7 h-7 rounded-md bg-[#FAF7F2] border border-[#E8E1D5] font-mono text-xs font-bold text-[#1E1B16] flex items-center justify-center shrink-0">
                            {idx + 1}
                          </span>
                          <span className="flex-1 text-base">
                            <RichText text={opt} />
                          </span>
                          {state.answered && isCorrectOpt && (
                            <span className="font-bold text-[#2E9E5B]" aria-label="Correct">✓</span>
                          )}
                          {state.answered && isPicked && !isCorrectOpt && (
                            <span className="font-bold text-[#FF4F58]" aria-label="Incorrect">✗</span>
                          )}
                        </button>
                      );
                    })}
                    {!state.answered && (
                      <p className="text-xs font-mono text-[#8C8477] pt-1">
                        Tip: press keys 1–4 or A–D to answer.
                      </p>
                    )}
                  </div>
                )}

                {currentQuestion.type === 'truefalse' && (
                  <div className="grid grid-cols-2 gap-3">
                    {['True', 'False'].map((opt, idx) => {
                      const isPicked = state.answered && state.answered.userAnswer === opt;
                      const isCorrectOpt = opt === currentQuestion.correctAnswer;
                      let btnStyle =
                        'bg-white border-[#E8E1D5] text-[#1E1B16] hover:border-[#FF7000] hover:bg-[#FFF7EF]';
                      if (state.answered) {
                        if (isCorrectOpt) btnStyle = 'bg-[#D6F5E3] border-[#2E9E5B] text-[#1E1B16]';
                        else if (isPicked) btnStyle = 'bg-[#FFE4E6] border-[#FF4F58] text-[#1E1B16]';
                        else btnStyle = 'bg-[#FAF7F2] border-[#E8E1D5] text-[#8C8477] opacity-70';
                      }
                      return (
                        <button
                          key={opt}
                          type="button"
                          disabled={Boolean(state.answered)}
                          onClick={() => dispatch({ type: 'SUBMIT_ANSWER', answer: opt })}
                          className={`min-h-[68px] px-4 py-4 rounded-xl border-2 font-bold text-lg flex flex-col items-center justify-center gap-1 transition-all cursor-pointer disabled:cursor-default ${btnStyle}`}
                        >
                          <span className="text-xl" aria-hidden="true">
                            {opt === 'True' ? '✓' : '✗'}
                          </span>
                          <span>{opt}</span>
                          <span className="text-[11px] font-mono text-[#8C8477]">
                            Key {idx + 1} / {opt[0]}
                          </span>
                        </button>
                      );
                    })}
                  </div>
                )}

                {currentQuestion.type === 'cellclick' &&
                  currentQuestion.visual?.mode === 'range' &&
                  !state.answered && (
                    <div className="flex items-center gap-3">
                      <button
                        type="button"
                        onClick={() => {
                          playClick();
                          setSelectedCells([]);
                        }}
                        disabled={selectedCells.length === 0}
                        className="min-h-[44px] px-4 py-2 rounded-xl border-2 border-[#E8E1D5] bg-white text-sm font-semibold text-[#5E574D] hover:border-[#1E1B16] hover:text-[#1E1B16] cursor-pointer disabled:opacity-40 disabled:cursor-not-allowed"
                      >
                        Clear
                      </button>
                      <button
                        type="button"
                        onClick={() => dispatch({ type: 'SUBMIT_ANSWER', answer: selectedCells })}
                        disabled={selectedCells.length === 0}
                        className="cq-btn-press flex-1 min-h-[46px] px-5 py-2.5 rounded-xl bg-[#FF7000] text-[#1E1B16] font-bold text-base shadow-[0_3px_0_#C25200] hover:bg-[#FF7F1A] cursor-pointer disabled:opacity-40 disabled:cursor-not-allowed"
                      >
                        Confirm Selection ({selectedCells.length} {selectedCells.length === 1 ? 'cell' : 'cells'})
                      </button>
                    </div>
                  )}

                {currentQuestion.type === 'builder' && (
                  <FormulaBuilder
                    question={currentQuestion}
                    disabled={Boolean(state.answered)}
                    onTokenClick={playClick}
                    onSubmit={(formula) => dispatch({ type: 'SUBMIT_ANSWER', answer: formula })}
                  />
                )}

                {currentQuestion.type === 'order' && (
                  <OrderChallenge
                    question={currentQuestion}
                    disabled={Boolean(state.answered)}
                    onMoveClick={playClick}
                    onSubmit={(ordered) => dispatch({ type: 'SUBMIT_ANSWER', answer: ordered })}
                  />
                )}

                {/* Feedback Panel */}
                {state.answered && (
                  <div
                    role="status"
                    aria-live="polite"
                    className={`cq-card-enter rounded-xl p-4 border-2 space-y-3 ${
                      state.answered.isCorrect
                        ? 'bg-[#E9F9EF] border-[#2E9E5B]'
                        : 'bg-[#FFF0F1] border-[#FF4F58]'
                    }`}
                  >
                    <div className="flex items-center justify-between gap-2">
                      <div className="font-extrabold text-base text-[#1E1B16] flex items-center gap-2">
                        <span
                          className={`inline-flex items-center justify-center w-6 h-6 rounded-full font-bold text-sm ${
                            state.answered.isCorrect
                              ? 'bg-[#2E9E5B] text-[#1E1B16]'
                              : 'bg-[#FF4F58] text-white'
                          }`}
                          aria-hidden="true"
                        >
                          {state.answered.isCorrect ? '✓' : '✗'}
                        </span>
                        <span>
                          {state.answered.isCorrect
                            ? 'Correct!'
                            : state.answered.timedOut
                            ? 'Time ran out!'
                            : 'Not quite!'}
                        </span>
                      </div>
                      {state.answered.pointsEarned > 0 && (
                        <span className="font-mono font-black text-sm px-2.5 py-0.5 rounded-full bg-white border border-[#2E9E5B] text-[#1E1B16]">
                          +{state.answered.pointsEarned} pts
                        </span>
                      )}
                    </div>

                    {!state.answered.isCorrect && (
                      <div className="text-sm text-[#1E1B16]">
                        <span className="font-semibold">Correct answer: </span>
                        <span className="font-mono font-bold bg-white px-1.5 py-0.5 rounded border border-[#E8E1D5]">
                          {formatAnswerForDisplay(currentQuestion.correctAnswer)}
                        </span>
                      </div>
                    )}

                    <p className="text-sm sm:text-base text-[#1E1B16]">
                      <RichText text={currentQuestion.explanation} />
                    </p>

                    <button
                      type="button"
                      onClick={() => {
                        playClick();
                        dispatch({ type: 'NEXT_QUESTION' });
                      }}
                      className="cq-btn-press w-full min-h-[48px] px-5 py-3 rounded-full bg-[#FF7000] text-[#1E1B16] font-extrabold text-base shadow-[0_3px_0_#C25200] hover:bg-[#FF7F1A] transition-all cursor-pointer"
                    >
                      {Number.isFinite(state.lives) && state.lives <= 0
                        ? 'See Summary →'
                        : state.index + 1 >= state.deck.length
                        ? 'View Results →'
                        : 'Next Question → (Enter)'}
                    </button>
                  </div>
                )}
              </div>
            </div>
          </div>
        )}

        {/* ================================================================ */}
        {/* SCREEN 3: GAME OVER (Lives exhausted)                            */}
        {/* ================================================================ */}
        {state.screen === 'gameover' && (
          <div className="cq-card-enter max-w-lg mx-auto pt-10">
            <div className="flex justify-center -mb-4 relative z-10">
              <PixelCat mood="dramatic" size={104} />
            </div>
            <div
              className="bg-white rounded-2xl border-2 border-[#1E1B16] shadow-md p-7 text-center space-y-4"
              style={{ transform: 'rotate(-1deg)' }}
            >
              <span className="inline-block px-3 py-1 rounded-full bg-[#FFE4E6] text-[#1E1B16] font-mono text-xs font-bold">
                #REF! · NO HEARTS LEFT
              </span>
              <h1 className="cq-display text-3xl sm:text-4xl font-black text-[#1E1B16]">
                Out of lives!
              </h1>
              <p className="text-base text-[#5E574D]">
                You scored <span className="font-mono font-bold text-[#1E1B16]">{state.score} pts</span> across{' '}
                {totalAnswered} {totalAnswered === 1 ? 'question' : 'questions'} ({totalCorrect} correct).
              </p>

              <div className="pt-2 space-y-3">
                <button
                  type="button"
                  onClick={() => {
                    playClick();
                    dispatch({ type: 'START_GAME', difficulty: state.difficulty });
                  }}
                  className="cq-btn-press w-full min-h-[50px] px-6 py-3 rounded-full bg-[#FF7000] text-[#1E1B16] font-extrabold text-base shadow-[0_3px_0_#C25200] hover:bg-[#FF7F1A] cursor-pointer"
                >
                  Retry {diffConfig.name} →
                </button>
                <button
                  type="button"
                  onClick={() => {
                    playClick();
                    dispatch({ type: 'GO_HOME' });
                  }}
                  className="min-h-[44px] px-4 py-2 text-sm font-semibold text-[#5E574D] hover:text-[#1E1B16] underline underline-offset-4 cursor-pointer"
                >
                  Change difficulty
                </button>
              </div>

              <p className="text-xs font-mono text-[#8C8477] pt-2">
                Every spreadsheet pro has hit #REF! at least once.
              </p>
            </div>
          </div>
        )}

        {/* ================================================================ */}
        {/* SCREEN 4: RESULTS SCREEN                                         */}
        {/* ================================================================ */}
        {state.screen === 'results' && (
          <div className="cq-card-enter space-y-6">
            {percentage >= 70 && <PixelConfetti />}

            {/* Top Hero Card: Pixel + Medal + Animated Score */}
            <div className="bg-white rounded-2xl border-2 border-[#E8E1D5] shadow-sm p-6 sm:p-8 text-center space-y-4">
              <div className="flex items-center justify-center gap-4">
                <PixelCat mood="cheering" medalTier={medalTier} size={96} />
                <ResultMedalSVG tier={medalTier} />
              </div>

              <div>
                <span className="inline-block px-3 py-1 rounded-full bg-[#FAF7F2] border border-[#E8E1D5] font-mono text-xs font-bold uppercase text-[#5E574D]">
                  {diffConfig.name} · {medalTitle}
                </span>
                <div className="mt-2 cq-display text-5xl sm:text-6xl font-black text-[#1E1B16] tabular-nums">
                  {displayScore} <span className="text-2xl text-[#FF7000]">pts</span>
                </div>
                <p className="mt-1 text-base text-[#5E574D]">
                  <span className="font-bold text-[#1E1B16]">{totalCorrect} / {totalAnswered}</span> correct ({percentage}%) · Best streak:{' '}
                  <span className="font-mono font-bold text-[#1E1B16]">{state.bestStreak}</span>
                </p>
              </div>

              <div className="flex flex-col sm:flex-row items-center justify-center gap-3 pt-2">
                <button
                  type="button"
                  onClick={() => {
                    playClick();
                    dispatch({ type: 'START_GAME', difficulty: state.difficulty });
                  }}
                  className="cq-btn-press w-full sm:w-auto min-h-[48px] px-8 py-3 rounded-full bg-[#FF7000] text-[#1E1B16] font-extrabold text-base shadow-[0_3px_0_#C25200] hover:bg-[#FF7F1A] cursor-pointer"
                >
                  Play Again →
                </button>
                <button
                  type="button"
                  onClick={() => {
                    playClick();
                    dispatch({ type: 'GO_HOME' });
                  }}
                  className="min-h-[44px] px-4 py-2 text-sm font-semibold text-[#5E574D] hover:text-[#1E1B16] underline underline-offset-4 cursor-pointer"
                >
                  Change difficulty
                </button>
              </div>
            </div>

            {/* Per-Topic Breakdown (In-Cell Spreadsheet Data Bars) */}
            <div className="bg-white rounded-2xl border-2 border-[#E8E1D5] p-5 sm:p-6 space-y-3">
              <h2 className="cq-display text-lg font-extrabold text-[#1E1B16]">
                Topic Breakdown
              </h2>
              <div className="space-y-2.5">
                {TOPICS.map((topic) => {
                  const stat = state.topicStats[topic];
                  if (!stat || stat.total === 0) return null;
                  const pct = Math.round((stat.correct / stat.total) * 100);
                  const strong = pct >= 70;
                  return (
                    <div key={topic} className="grid grid-cols-12 items-center gap-3">
                      <div className="col-span-5 sm:col-span-4 flex items-center gap-2 min-w-0">
                        <span className="px-2 py-0.5 rounded bg-[#EBF3FC] border border-[#4E8FE0]/30 font-mono text-[11px] font-bold text-[#1E1B16] shrink-0">
                          {TOPIC_SHORT[topic]}
                        </span>
                        <span className="text-xs text-[#5E574D] truncate hidden sm:inline">{topic}</span>
                      </div>
                      <div className="col-span-5 sm:col-span-6 h-7 rounded-md bg-[#FAF7F2] border border-[#E8E1D5] overflow-hidden relative">
                        <div
                          className="h-full transition-all duration-500"
                          style={{
                            width: `${pct}%`,
                            backgroundColor: strong ? '#2E9E5B' : '#F59E0B',
                            opacity: 0.35,
                          }}
                        />
                        <div
                          className="absolute inset-y-0 left-0 w-1"
                          style={{ backgroundColor: strong ? '#2E9E5B' : '#F59E0B' }}
                        />
                        <span className="absolute inset-0 flex items-center px-2.5 font-mono text-xs font-bold text-[#1E1B16]">
                          {pct}%
                        </span>
                      </div>
                      <div className="col-span-2 text-right font-mono text-xs font-bold text-[#5E574D]">
                        {stat.correct}/{stat.total}
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>

            {/* Review Answers (Zebra-striped worksheet rows, expandable) */}
            <div className="bg-white rounded-2xl border-2 border-[#E8E1D5] overflow-hidden">
              <div className="px-5 py-4 border-b border-[#E8E1D5] flex items-center justify-between bg-[#FAF7F2]">
                <h2 className="cq-display text-lg font-extrabold text-[#1E1B16]">
                  Review Answers
                </h2>
                <button
                  type="button"
                  onClick={() => {
                    const allOpen = state.answers.every((_, idx) => expandedReview[idx]);
                    const next = {};
                    if (!allOpen) state.answers.forEach((_, idx) => (next[idx] = true));
                    setExpandedReview(next);
                  }}
                  className="min-h-[36px] px-3 py-1 rounded-lg border border-[#E8E1D5] bg-white text-xs font-mono font-bold text-[#5E574D] hover:border-[#1E1B16] hover:text-[#1E1B16] cursor-pointer"
                >
                  {state.answers.every((_, idx) => expandedReview[idx]) ? 'Collapse all' : 'Expand all'}
                </button>
              </div>

              <div className="divide-y divide-[#E8E1D5]">
                {state.answers.map((entry, idx) => {
                  const open = Boolean(expandedReview[idx]);
                  return (
                    <div
                      key={`${entry.question.id}-${idx}`}
                      className={idx % 2 === 0 ? 'bg-white' : 'bg-[#FAF7F2]/60'}
                    >
                      <button
                        type="button"
                        onClick={() =>
                          setExpandedReview((prev) => ({ ...prev, [idx]: !prev[idx] }))
                        }
                        aria-expanded={open}
                        className="w-full min-h-[52px] px-5 py-3 text-left flex items-center gap-3 hover:bg-[#FFF7EF] transition-colors cursor-pointer"
                      >
                        <span className="font-mono text-xs text-[#8C8477] w-6 shrink-0">
                          {idx + 1}.
                        </span>
                        <span
                          className={`w-6 h-6 rounded-md font-bold text-xs flex items-center justify-center shrink-0 ${
                            entry.isCorrect
                              ? 'bg-[#D6F5E3] text-[#1E1B16] border border-[#2E9E5B]'
                              : 'bg-[#FFE4E6] text-[#1E1B16] border border-[#FF4F58]'
                          }`}
                          aria-label={entry.isCorrect ? 'Correct' : 'Incorrect'}
                        >
                          {entry.isCorrect ? '✓' : '✗'}
                        </span>
                        <span className="flex-1 text-sm font-semibold text-[#1E1B16] truncate">
                          {entry.question.prompt}
                        </span>
                        <span className="font-mono text-xs text-[#8C8477] shrink-0">
                          {open ? '▲' : '▼'}
                        </span>
                      </button>

                      {open && (
                        <div className="px-5 pb-4 pt-1 pl-14 space-y-1.5 text-sm border-t border-[#E8E1D5]/60">
                          <div>
                            <span className="font-semibold text-[#5E574D]">Your answer: </span>
                            <span className="font-mono text-[#1E1B16]">
                              {formatAnswerForDisplay(entry.userAnswer)}
                            </span>
                          </div>
                          <div>
                            <span className="font-semibold text-[#5E574D]">Correct answer: </span>
                            <span className="font-mono font-bold text-[#1E1B16]">
                              {formatAnswerForDisplay(entry.question.correctAnswer)}
                            </span>
                          </div>
                          <p className="text-[#5E574D] pt-1">
                            <RichText text={entry.question.explanation} />
                          </p>
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
