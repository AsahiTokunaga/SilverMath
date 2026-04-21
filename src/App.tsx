import { useRef, useState } from "react";
import { Packer } from "docx";
import "@/App.css";
import { downloadBlob } from "@/utils/download";

export const QUESTION_COUNT_DEFAULT = 20;
export const QUESTION_COUNT_MIN = 5;
export const QUESTION_COUNT_MAX = 40;
export const EQUATION_LENGTH_DEFAULT = 7;
export const EQUATION_LENGTH_MIN = 5;
export const EQUATION_LENGTH_MAX = 10;
export const SUM_MIN = 1;
export const SUM_MAX = 60;
export const TERM_MIN = -30;
export const TERM_MAX = 30;
export const MAX_GENERATION_ATTEMPTS = 1000;
export const FIRST_INDEX = 0;
export const EMPTY_TERM = 0;
export const STARTING_SUM = 0;
export const NO_REMAINING_TERMS = 0;
export const SINGLE_REMAINING_TERM = 1;
export const MIN_POSITIVE_TERM = 1;
export const INCLUSIVE_RANGE_STEP = 1;
export const ZENKAKU_OFFSET = 0xfee0;
export const WORKSHEET_HEADER_TERM_RANGE = `ー${Math.abs(TERM_MIN)}～${TERM_MAX}`;
export const A4_PAGE_WIDTH_TWIPS = 11906;
export const A4_PAGE_HEIGHT_TWIPS = 16838;
export const AVAILABLE_HEIGHT_PT = 729;
export const AVAILABLE_WIDTH_PT = 523;
export const QUESTIONS_PER_PAGE = 10;
export const PAGE_HEIGHT_SAFETY_MARGIN_PT = 20;
export const MIN_FONT_SIZE_PT = 8;
export const PROBLEM_LINE_PADDING_CHARS = 2;
export const PROBLEM_VERTICAL_SPAN = 2;
export const TWIPS_PER_POINT = 20;
export const PROBLEM_TEXT_SCALE = 1.8;
export const HEADER_TEXT_SIZE_HALF_POINTS = 48;
export const FOOTER_TEXT_SIZE_HALF_POINTS = 24;
export const FULL_PERCENT = 100;
export const FOOTER_INFO_CELL_WIDTH_PERCENT = 85;
export const FOOTER_PAGE_CELL_WIDTH_PERCENT = 15;
export const PAGE_HEADER_MARGIN_TWIPS = 400;
export const PAGE_FOOTER_MARGIN_TWIPS = 400;
export const PAGE_EDGE_MARGIN_TWIPS = 720;
export const BORDER_NONE_SIZE = 0;
export const BORDER_NONE_COLOR = "FFFFFF";
export const FOCUS_DELAY_MS = 0;
export const FULL_PAGE_END_PARAGRAPH_RESERVE_PT = 16;
export const BODY_FONT_FAMILY = "MS Gothic";

function App() {
  const [questionCount, setQuestionCount] = useState<number>(
    QUESTION_COUNT_DEFAULT,
  );
  const [equationLength, setEquationLength] = useState<number>(
    EQUATION_LENGTH_DEFAULT,
  );
  const [creatorName, setCreatorName] = useState<string>("");
  const [solverNumber, setSolverNumber] = useState<string>("");
  const creatorNameInputRef = useRef<HTMLInputElement>(null);
  const solverNumberInputRef = useRef<HTMLInputElement>(null);

  const focusInput = (inputRef: React.RefObject<HTMLInputElement | null>) => {
    window.setTimeout(() => {
      inputRef.current?.focus();
    }, FOCUS_DELAY_MS);
  };

  const handleCreate = async () => {
    const normalizedCreatorName = creatorName.trim();
    const normalizedSolverNumber = solverNumber.trim();

    if (normalizedCreatorName === "" || normalizedSolverNumber === "") {
      alert("作成者、番号を入力してください。");
      if (normalizedCreatorName === "") {
        focusInput(creatorNameInputRef);
        return;
      }

      focusInput(solverNumberInputRef);
      return;
    }

    try {
      const {
        buildWorksheetFileName,
        createWorksheetDocument,
        formatTodayJst,
        formatTodayJstForFile,
        generateWorksheetExpressions,
      } = await import("@/features/worksheet");

      const problems = generateWorksheetExpressions(
        questionCount,
        equationLength,
      );
      const now = new Date();
      const todayJst = formatTodayJst(now);
      const todayJstForFile = formatTodayJstForFile(now);

      const doc = createWorksheetDocument({
        problemExpressions: problems,
        questionCount,
        creatorName: normalizedCreatorName,
        solverNumber: normalizedSolverNumber,
        todayJst,
      });

      const blob = await Packer.toBlob(doc);
      downloadBlob(
        blob,
        buildWorksheetFileName({
          questionCount,
          creatorName: normalizedCreatorName,
          solverNumber: normalizedSolverNumber,
          todayJst: todayJstForFile,
        }),
      );
    } catch (error) {
      console.error(error);
      alert("問題の作成に失敗しました。もう一度お試しください。");
      focusInput(solverNumberInputRef);
    }
  };

  return (
    <div className="app-container">
      <div className="card">
        <h1 className="title">算数問題作成 (A4縦)</h1>

        <div className="input-group">
          <label className="label">作成者</label>
          <input
            type="text"
            className="text-input"
            value={creatorName}
            onChange={(e) => setCreatorName(e.target.value)}
            ref={creatorNameInputRef}
            placeholder="名前"
            required
          />
        </div>
        <div className="input-group">
          <label className="label">番号</label>
          <input
            type="number"
            className="text-input"
            value={solverNumber}
            onChange={(e) => setSolverNumber(e.target.value)}
            ref={solverNumberInputRef}
            placeholder="番号（半角数字で入力してください）"
            required
          />
        </div>

        <div className="input-group">
          <label className="label">
            作成する問題数:{" "}
            <span className="highlight-text">{questionCount}</span>
          </label>
          <input
            type="range"
            min={QUESTION_COUNT_MIN}
            max={QUESTION_COUNT_MAX}
            value={questionCount}
            onChange={(e) => setQuestionCount(Number(e.target.value))}
            className="slider"
          />
          <div className="slider-labels">
            <span>{QUESTION_COUNT_MIN}問</span>
            <span>{QUESTION_COUNT_MAX}問</span>
          </div>
        </div>

        <div className="input-group">
          <label className="label">
            式の長さ（項数）:{" "}
            <span className="highlight-text">{equationLength}</span>
          </label>
          <input
            type="range"
            min={EQUATION_LENGTH_MIN}
            max={EQUATION_LENGTH_MAX}
            value={equationLength}
            onChange={(e) => setEquationLength(Number(e.target.value))}
            className="slider"
          />
          <div className="slider-labels">
            <span>{EQUATION_LENGTH_MIN}項</span>
            <span>{EQUATION_LENGTH_MAX}項</span>
          </div>
        </div>

        <button onClick={handleCreate} className="create-button">
          作成
        </button>
      </div>
    </div>
  );
}

export default App;
