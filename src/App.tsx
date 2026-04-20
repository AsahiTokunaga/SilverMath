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
  TableCell,
  TableRow,
  Table,
  BorderStyle,
  HeightRule,
  WidthType,
  VerticalAlign,
} from "docx";
import "./App.css";

function App() {
  const [questionCount, setQuestionCount] = useState<number>(20);
  const [creatorName, setCreatorName] = useState<string>("");
  const [solverName, setSolverName] = useState<string>("");
  const [solverNumber, setSolverNumber] = useState<string>("");

  const toZenkaku = (str: string) => {
    return str.replace(/[A-Za-z0-9=+-\s]/g, (char) => {
      if (char === " ") return "　";
      return String.fromCharCode(char.charCodeAt(0) + 0xfee0);
    });
  };

  const handleCreate = async () => {
    if (creatorName === "" || solverName === "" || solverNumber === "") {
      alert("作成者、解答者、番号を入力してください。");
      return;
    }
    const problems = [];
    for (let i = 0; i < questionCount; i++) {
      const termCount = Math.floor(Math.random() * 4) + 7;
      let eq = "";
      for (let j = 0; j < termCount; j++) {
        let n = Math.floor(Math.random() * 101) - 50;
        while (n === 0) {
          n = Math.floor(Math.random() * 101) - 50;
        }
        if (j === 0) {
          eq += n.toString();
        } else {
          eq += n > 0 ? `+${n}` : `${n}`;
        }
      }
      eq += "=";
      problems.push(toZenkaku(eq));
    }

    const availableHeightPt = 480;
    let fontSizePt = Math.floor(availableHeightPt / questionCount);

    const maxCalcLength = Math.max(...problems.map((prob) => prob.length - 1));
    const maxLineLength = maxCalcLength + 2;

    const maxFontByWidth = Math.floor(770 / maxLineLength);
    if (fontSizePt > maxFontByWidth) fontSizePt = maxFontByWidth;
    if (fontSizePt < 8) fontSizePt = 8;

    const todayJST = new Date().toLocaleDateString("ja-JP", {
      timeZone: "Asia/Tokyo",
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
    });

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
                top: 720,
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
                      text: `脳トレ用　計算問題　－５０～５０　${toZenkaku(questionCount.toString())}問　　（　　/${toZenkaku(questionCount.toString())}）`,
                      size: 48,
                    }),
                  ],
                }),
              ],
            }),
          },
          footers: {
            default: new Footer({
              children: [
                new Table({
                  borders: {
                    top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                    bottom: {
                      style: BorderStyle.NONE,
                      size: 0,
                      color: "FFFFFF",
                    },
                    left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                    right: {
                      style: BorderStyle.NONE,
                      size: 0,
                      color: "FFFFFF",
                    },
                    insideHorizontal: {
                      style: BorderStyle.NONE,
                      size: 0,
                      color: "FFFFFF",
                    },
                    insideVertical: {
                      style: BorderStyle.NONE,
                      size: 0,
                      color: "FFFFFF",
                    },
                  },
                  width: {
                    size: 100,
                    type: WidthType.PERCENTAGE,
                  },
                  rows: [
                    new TableRow({
                      children: [
                        new TableCell({
                          width: {
                            size: 85,
                            type: WidthType.PERCENTAGE,
                          },
                          verticalAlign: VerticalAlign.BOTTOM,
                          borders: {
                            top: {
                              style: BorderStyle.NONE,
                              size: 0,
                              color: "FFFFFF",
                            },
                            bottom: {
                              style: BorderStyle.NONE,
                              size: 0,
                              color: "FFFFFF",
                            },
                            left: {
                              style: BorderStyle.NONE,
                              size: 0,
                              color: "FFFFFF",
                            },
                            right: {
                              style: BorderStyle.NONE,
                              size: 0,
                              color: "FFFFFF",
                            },
                          },
                          children: [
                            new Paragraph({
                              children: [
                                new TextRun({
                                  text: `作成者: ${creatorName}　　解答者: ${solverName}　　番号: ${solverNumber}　　作成日: ${todayJST}`,
                                  size: 24,
                                }),
                              ],
                            }),
                          ],
                        }),

                        new TableCell({
                          width: {
                            size: 15,
                            type: WidthType.PERCENTAGE,
                          },
                          verticalAlign: VerticalAlign.BOTTOM,
                          borders: {
                            top: {
                              style: BorderStyle.NONE,
                              size: 0,
                              color: "FFFFFF",
                            },
                            bottom: {
                              style: BorderStyle.NONE,
                              size: 0,
                              color: "FFFFFF",
                            },
                            left: {
                              style: BorderStyle.NONE,
                              size: 0,
                              color: "FFFFFF",
                            },
                            right: {
                              style: BorderStyle.NONE,
                              size: 0,
                              color: "FFFFFF",
                            },
                          },
                          children: [
                            new Paragraph({
                              alignment: AlignmentType.RIGHT,
                              children: [
                                new TextRun({
                                  children: [PageNumber.CURRENT],
                                  size: 24,
                                }),
                              ],
                            }),
                          ],
                        }),
                      ],
                    }),
                  ],
                }),
              ],
            }),
          },
          children: problems.map((prob) => {
            const calcText = prob.slice(0, -1);

            return new Table({
              borders: {
                top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                insideHorizontal: {
                  style: BorderStyle.NONE,
                  size: 0,
                  color: "FFFFFF",
                },
                insideVertical: {
                  style: BorderStyle.NONE,
                  size: 0,
                  color: "FFFFFF",
                },
              },
              rows: [
                new TableRow({
                  height: {
                    value: fontSizePt * 2 * 20, // 1pt = 20px
                    rule: HeightRule.EXACT,
                  },
                  children: [
                    new TableCell({
                      children: [
                        new Paragraph({
                          alignment: AlignmentType.DISTRIBUTE,
                          children: [
                            new TextRun({
                              text: calcText,
                              size: fontSizePt * 1.8,
                            }),
                          ],
                        }),
                      ],
                    }),
                    new TableCell({
                      children: [
                        new Paragraph({
                          alignment: AlignmentType.RIGHT,
                          children: [
                            new TextRun({
                              text: " ＝",
                              size: fontSizePt * 1.8,
                            }),
                          ],
                        }),
                      ],
                    }),
                  ],
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
    a.download = `脳トレ用計算問題_${questionCount}問_${creatorName}_${solverName}_No.${solverNumber}.docx`;
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
          <label className="label">作成者</label>
          <input
            type="text"
            className="text-input"
            value={creatorName}
            onChange={(e) => setCreatorName(e.target.value)}
            placeholder="名前"
            required
          />
        </div>
        <div className="input-group">
          <label className="label">解答者</label>
          <input
            type="text"
            className="text-input"
            value={solverName}
            onChange={(e) => setSolverName(e.target.value)}
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
