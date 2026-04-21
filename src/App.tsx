import { useRef, useState } from "react";
import { Packer } from "docx";
import "./App.css";
import {
  buildWorksheetFileName,
  calculateWorksheetFontSize,
  createWorksheetDocument,
  formatTodayJst,
  formatTodayJstForFile,
  generateWorksheetExpressions,
} from "./features/worksheet";
import { downloadBlob } from "./utils/download";

function App() {
  const [questionCount, setQuestionCount] = useState<number>(20);
  const [creatorName, setCreatorName] = useState<string>("");
  const [solverNumber, setSolverNumber] = useState<string>("");
  const creatorNameInputRef = useRef<HTMLInputElement>(null);
  const solverNumberInputRef = useRef<HTMLInputElement>(null);

  const focusInput = (inputRef: React.RefObject<HTMLInputElement | null>) => {
    window.setTimeout(() => {
      inputRef.current?.focus();
    }, 0);
  };

  const handleCreate = async () => {
    if (creatorName === "" || solverNumber === "") {
      alert("作成者と番号を入力してください。");
      if (creatorName === "") {
        focusInput(creatorNameInputRef);
        return;
      }

      focusInput(solverNumberInputRef);
      return;
    }

    let problems: string[];

    try {
      problems = generateWorksheetExpressions(questionCount);
    } catch (error) {
      console.error(error);
      alert("問題の作成に失敗しました。もう一度お試しください。");
      focusInput(creatorNameInputRef);
      return;
    }

    const fontSizePt = calculateWorksheetFontSize(problems, questionCount);
    const now = new Date();
    const todayJst = formatTodayJst(now);
    const todayJstForFile = formatTodayJstForFile(now);

    const doc = createWorksheetDocument({
      problemExpressions: problems,
      questionCount,
      creatorName,
      solverNumber,
      todayJst,
      fontSizePt,
    });

    const blob = await Packer.toBlob(doc);
    downloadBlob(
      blob,
      buildWorksheetFileName({
        questionCount,
        creatorName,
        solverNumber,
        todayJst: todayJstForFile,
      }),
    );
  };

  return (
    <div className="app-container">
      <div className="card">
        <h1 className="title">算数問題作成 (A4横)</h1>

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
            min="5"
            max="40"
            value={questionCount}
            onChange={(e) => setQuestionCount(parseInt(e.target.value))}
            className="slider"
          />
          <div className="slider-labels">
            <span>5問</span>
            <span>40問</span>
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
