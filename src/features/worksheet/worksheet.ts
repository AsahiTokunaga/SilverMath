import {
  EMPTY_TERM,
  EQUATION_LENGTH_DEFAULT,
  EQUATION_LENGTH_MAX,
  EQUATION_LENGTH_MIN,
  FIRST_INDEX,
  INCLUSIVE_RANGE_STEP,
  MAX_GENERATION_ATTEMPTS,
  MIN_POSITIVE_TERM,
  NO_REMAINING_TERMS,
  SINGLE_REMAINING_TERM,
  STARTING_SUM,
  SUM_MAX,
  SUM_MIN,
  TERM_MAX,
  TERM_MIN,
  ZENKAKU_OFFSET,
} from "@/App";

export function randInt(min: number, max: number): number {
  if (max < min) {
    throw new Error("randInt requires max to be greater than or equal to min.");
  }

  return Math.floor(Math.random() * (max - min + INCLUSIVE_RANGE_STEP)) + min;
}

export function randIntExceptZero(min: number, max: number): number {
  if (min === EMPTY_TERM && max === EMPTY_TERM) {
    throw new Error(
      "randIntExceptZero does not support a range containing only 0.",
    );
  }

  let value = EMPTY_TERM;

  do {
    value = randInt(min, max);
  } while (value === EMPTY_TERM);

  return value;
}

export function toZenkaku(str: string): string {
  return str.replace(/[A-Za-z0-9=+\- ]/g, (char) => {
    if (char === " ") return "　";
    return String.fromCharCode(char.charCodeAt(0) + ZENKAKU_OFFSET);
  });
}

function formatExpression(terms: number[]): string {
  return terms
    .map((term, index) => {
      if (index === FIRST_INDEX) return term.toString();
      return term >= STARTING_SUM ? `+${term}` : `${term}`;
    })
    .join("");
}

function getNextTermMinimum(currentSum: number): number {
  if (currentSum === STARTING_SUM) {
    return MIN_POSITIVE_TERM;
  }

  return Math.max(TERM_MIN, -currentSum);
}

function shuffleNumbers(values: number[]): number[] {
  for (
    let index = values.length - SINGLE_REMAINING_TERM;
    index > FIRST_INDEX;
    index--
  ) {
    const swapIndex = randInt(FIRST_INDEX, index);
    [values[index], values[swapIndex]] = [values[swapIndex], values[index]];
  }

  return values;
}

function canCompleteExpression(
  currentSum: number,
  remainingTerms: number,
  targetSum: number,
  memo: Map<string, boolean>,
): boolean {
  const cacheKey = `${currentSum}:${remainingTerms}`;
  const cached = memo.get(cacheKey);

  if (cached !== undefined) {
    return cached;
  }

  let result = false;

  if (remainingTerms === NO_REMAINING_TERMS) {
    result = currentSum === targetSum;
  } else if (remainingTerms === SINGLE_REMAINING_TERM) {
    const lastTerm = targetSum - currentSum;
    result =
      lastTerm !== EMPTY_TERM && lastTerm >= TERM_MIN && lastTerm <= TERM_MAX;
  } else {
    const minimumTerm = getNextTermMinimum(currentSum);

    for (let term = minimumTerm; term <= TERM_MAX; term++) {
      if (term === EMPTY_TERM) continue;

      if (
        canCompleteExpression(
          currentSum + term,
          remainingTerms - SINGLE_REMAINING_TERM,
          targetSum,
          memo,
        )
      ) {
        result = true;
        break;
      }
    }
  }

  memo.set(cacheKey, result);
  return result;
}

function buildExpressionTerms(
  currentSum: number,
  remainingTerms: number,
  targetSum: number,
  memo: Map<string, boolean>,
): number[] | null {
  if (remainingTerms === NO_REMAINING_TERMS) {
    return currentSum === targetSum ? [] : null;
  }

  if (remainingTerms === SINGLE_REMAINING_TERM) {
    const lastTerm = targetSum - currentSum;

    if (lastTerm === EMPTY_TERM || lastTerm < TERM_MIN || lastTerm > TERM_MAX) {
      return null;
    }

    return [lastTerm];
  }

  const candidateTerms: number[] = [];
  const minimumTerm = getNextTermMinimum(currentSum);

  for (let term = minimumTerm; term <= TERM_MAX; term++) {
    if (term !== EMPTY_TERM) {
      candidateTerms.push(term);
    }
  }

  shuffleNumbers(candidateTerms);

  for (const term of candidateTerms) {
    const nextSum = currentSum + term;

    if (
      !canCompleteExpression(
        nextSum,
        remainingTerms - SINGLE_REMAINING_TERM,
        targetSum,
        memo,
      )
    ) {
      continue;
    }

    const restTerms = buildExpressionTerms(
      nextSum,
      remainingTerms - SINGLE_REMAINING_TERM,
      targetSum,
      memo,
    );

    if (restTerms !== null) {
      return [term, ...restTerms];
    }
  }

  return null;
}

function validateEquationLength(equationLength: number): void {
  if (!Number.isInteger(equationLength)) {
    throw new Error("equationLength must be an integer.");
  }

  if (
    equationLength < EQUATION_LENGTH_MIN ||
    equationLength > EQUATION_LENGTH_MAX
  ) {
    throw new Error("equationLength is out of supported range.");
  }
}

function generateWorksheetExpression(equationLength: number): string {
  validateEquationLength(equationLength);

  for (
    let attempt = FIRST_INDEX;
    attempt < MAX_GENERATION_ATTEMPTS;
    attempt++
  ) {
    const targetSum = randInt(SUM_MIN, SUM_MAX);
    const firstTerm = randInt(MIN_POSITIVE_TERM, TERM_MAX);
    const memo = new Map<string, boolean>();
    const remainingTerms = equationLength - SINGLE_REMAINING_TERM;

    if (!canCompleteExpression(firstTerm, remainingTerms, targetSum, memo)) {
      continue;
    }

    const restTerms = buildExpressionTerms(
      firstTerm,
      remainingTerms,
      targetSum,
      memo,
    );

    if (restTerms === null) {
      continue;
    }

    return formatExpression([firstTerm, ...restTerms]);
  }

  throw new Error("Failed to generate a worksheet expression.");
}

export function generateWorksheetExpressions(
  questionCount: number,
  equationLength = EQUATION_LENGTH_DEFAULT,
): string[] {
  return Array.from({ length: questionCount }, () =>
    generateWorksheetExpression(equationLength),
  );
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
  const safeCreatorName =
    params.creatorName.trim().replace(/[\\/:*?"<>|]/g, "_") || "未設定";
  const safeSolverNumber =
    params.solverNumber.trim().replace(/[\\/:*?"<>|]/g, "_") || "0";

  return `脳トレ用計算問題_${params.questionCount}問_${safeCreatorName}_No.${safeSolverNumber}_${params.todayJst}.docx`;
}
