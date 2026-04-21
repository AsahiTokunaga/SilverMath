export const SUM_MIN = 1;
export const SUM_MAX = 60;
export const TERM_MIN = -30;
export const TERM_MAX = 30;
export const TERM_COUNT_MIN = 7;
export const TERM_COUNT_MAX = 10;
export const WORKSHEET_HEADER_TERM_RANGE = `ー${Math.abs(TERM_MIN)}～${TERM_MAX}`;

export function randInt(min: number, max: number): number {
  if (max < min) {
    throw new Error("randInt requires max to be greater than or equal to min.");
  }

  return Math.floor(Math.random() * (max - min + 1)) + min;
}

export function randIntExceptZero(min: number, max: number): number {
  if (min === 0 && max === 0) {
    throw new Error("randIntExceptZero does not support a range containing only 0.");
  }

  let value = 0;

  do {
    value = randInt(min, max);
  } while (value === 0);

  return value;
}

export function toZenkaku(str: string): string {
  return str.replace(/[A-Za-z0-9=+\- ]/g, (char) => {
    if (char === " ") return "　";
    return String.fromCharCode(char.charCodeAt(0) + 0xfee0);
  });
}

function formatExpression(terms: number[]): string {
  return terms
    .map((term, index) => {
      if (index === 0) return term.toString();
      return term >= 0 ? `+${term}` : `${term}`;
    })
    .join("");
}

function generateWorksheetExpression(): string {
  for (let attempt = 0; attempt < 1000; attempt++) {
    const termCount = randInt(TERM_COUNT_MIN, TERM_COUNT_MAX);
    const terms: number[] = [randInt(1, TERM_MAX)];

    for (let index = 1; index < termCount - 1; index++) {
      terms.push(randIntExceptZero(TERM_MIN, TERM_MAX));
    }

    const targetSum = randInt(SUM_MIN, SUM_MAX);
    const currentSum = terms.reduce((sum, term) => sum + term, 0);
    const lastTerm = targetSum - currentSum;

    if (lastTerm === 0) continue;
    if (lastTerm < TERM_MIN || lastTerm > TERM_MAX) continue;

    terms.push(lastTerm);
    return formatExpression(terms);
  }

  throw new Error("Failed to generate a worksheet expression.");
}

export function generateWorksheetExpressions(questionCount: number): string[] {
  return Array.from({ length: questionCount }, () => generateWorksheetExpression());
}

export function formatTodayJst(date = new Date()): string {
  return new Intl.DateTimeFormat("ja-JP", {
    timeZone: "Asia/Tokyo",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).format(date);
}

export function formatTodayJstForFile(date = new Date()): string {
  return formatTodayJst(date).replace(/\//g, "-");
}

export function buildWorksheetFileName(params: {
  questionCount: number;
  creatorName: string;
  solverNumber: string;
  todayJst: string;
}): string {
  const safeCreatorName = params.creatorName.trim().replace(/[\\/:*?"<>|]/g, "_") || "未設定";
  const safeSolverNumber = params.solverNumber.trim().replace(/[\\/:*?"<>|]/g, "_") || "0";

  return `脳トレ用計算問題_${params.questionCount}問_${safeCreatorName}_No.${safeSolverNumber}_${params.todayJst}.docx`;
}