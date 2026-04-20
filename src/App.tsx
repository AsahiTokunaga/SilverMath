import { useState } from "react";
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  PageOrientation,
  Header,
  Footer,
  AlignmentType,
  PageNumber,
} from "docx";
import "./App.css";

function App() {
  const [questionCount, setQuestionCount] = useState<number>(20);

  const toZenkaku = (str: string) => {
    return str.replace(/[A-Za-z0-9=+-\s]/g, (char) => {
      if (char === ' ') return '　';
      return String.fromCharCode(char.charCodeAt(0) + 0xfee0);
    });
  };

  const handleCreate = async () => {
    const problems = [];
    for (let i = 0; i < questionCount; i++) {
      const termCount = Math.floor(Math.random() * 4) + 7; // 7 to 10
      let eq = "";
      for (let j = 0; j < termCount; j++) {
        let n = Math.floor(Math.random() * 101) - 50;
        while (n === 0) {
          n = Math.floor(Math.random() * 101) - 50;
        }
        if (j === 0) {
          eq += n.toString();
        } else {
          eq += n > 0 ? `+${n}` : `${n}`; // n is already negative, so toString() includes '-'
        }
      }
      eq += "=";
      problems.push(toZenkaku(eq));
    }

    // A4 Landscape available height is approx 520 points (with 12.7mm margins)
    // We adjust font size to occupy the available height
    // Word's default line spacing is usually 1.15x the font size
    const availableHeightPt = 480;
    let fontSizePt = Math.floor(availableHeightPt / questionCount);

    // Ensure it's bounded (don't exceed width length of max 31 characters)
    // width ~ 770 points.
    const maxFontByWidth = Math.floor(770 / 32); // ~24
    if (fontSizePt > maxFontByWidth) fontSizePt = maxFontByWidth;
    if (fontSizePt < 8) fontSizePt = 8; // min size

    const doc = new Document({
      sections: [
        {
          properties: {
            page: {
              size: {
                orientation: PageOrientation.LANDSCAPE,
              },
              margin: {
                header: 400,
                footer: 400,
                top: 720, // 0.5 inch
                bottom: 720,
                left: 720,
                right: 720,
              },
            },
          },
          headers: {
            default: new Header({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: `脳トレ用　計算問題　－５０～５０　${toZenkaku(questionCount.toString())}問`,
                      size: 48, // 24pt (48 half-points)
                    }),
                  ],
                }),
              ],
            }),
          },
          footers: {
            default: new Footer({
              children: [
                new Paragraph({
                  alignment: AlignmentType.RIGHT,
                  children: [
                    new TextRun({
                      children: [PageNumber.CURRENT],
                      size: 24, // 12pt
                    }),
                  ],
                }),
              ],
            }),
          },
          children: problems.map((prob) => {
            return new Paragraph({
              children: [
                new TextRun({
                  text: prob,
                  size: fontSizePt * 2, // size is in half-points
                }),
              ],
            });
          }),
        },
      ],
    });

    const blob = await Packer.toBlob(doc);
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `Math_Problems_${questionCount}.docx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  return (
    <div className="app-container">
      <div className="card">
        <h1 className="title">算数問題作成 (A4横)</h1>

        <div className="input-group">
          <label className="label">
            作成する問題数: <span className="highlight-text">{questionCount}</span>
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

        <button
          onClick={handleCreate}
          className="create-button"
        >
          作成
        </button>
      </div>
    </div>
  );
}

export default App;
