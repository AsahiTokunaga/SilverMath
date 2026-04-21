import {
  AlignmentType,
  BorderStyle,
  Document,
  Footer,
  Header,
  HeightRule,
  PageNumber,
  PageOrientation,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  VerticalAlign,
  WidthType,
} from "docx";
import { toZenkaku, WORKSHEET_HEADER_TERM_RANGE } from "./worksheet";

const AVAILABLE_HEIGHT_PT = 480;
const AVAILABLE_WIDTH_PT = 770;
const QUESTIONS_PER_PAGE = 10;
const PAGE_HEIGHT_SAFETY_MARGIN_PT = 20;
const MIN_FONT_SIZE_PT = 8;

function createNoBorder() {
  return { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
}

function createTableBorders(includeInsideBorders = true) {
  const border = createNoBorder();

  if (includeInsideBorders) {
    return {
      top: border,
      bottom: border,
      left: border,
      right: border,
      insideHorizontal: border,
      insideVertical: border,
    };
  }

  return {
    top: border,
    bottom: border,
    left: border,
    right: border,
  };
}

export function calculateWorksheetFontSize(problemExpressions: string[], questionCount: number): number {
  const questionsOnPage = Math.min(questionCount, QUESTIONS_PER_PAGE);
  const usableHeightPt = AVAILABLE_HEIGHT_PT - PAGE_HEIGHT_SAFETY_MARGIN_PT;

  let fontSizePt = Math.floor(usableHeightPt / (questionsOnPage * 2));
  const maxCalcLength = Math.max(...problemExpressions.map((expression) => expression.length));
  const maxLineLength = maxCalcLength + 2;
  const maxFontByWidth = Math.floor(AVAILABLE_WIDTH_PT / maxLineLength);

  if (fontSizePt > maxFontByWidth) fontSizePt = maxFontByWidth;
  if (fontSizePt < MIN_FONT_SIZE_PT) fontSizePt = MIN_FONT_SIZE_PT;

  return fontSizePt;
}

function createProblemTable(problemExpression: string, fontSizePt: number): Table {
  return new Table({
    borders: createTableBorders(),
    rows: [
      new TableRow({
        height: {
          value: fontSizePt * 2 * 20,
          rule: HeightRule.EXACT,
        },
        children: [
          new TableCell({
            children: [
              new Paragraph({
                alignment: AlignmentType.DISTRIBUTE,
                children: [
                  new TextRun({
                    text: toZenkaku(problemExpression),
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
}

function createHeader(questionCount: number): Header {
  return new Header({
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({
            text: `脳トレ用　計算問題　${WORKSHEET_HEADER_TERM_RANGE}　${questionCount.toString()}問　　（　　/${questionCount.toString()}）`,
            size: 48,
          }),
        ],
      }),
    ],
  });
}

function createFooter(creatorName: string, solverNumber: string, todayJst: string): Footer {
  return new Footer({
    children: [
      new Table({
        borders: createTableBorders(),
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
                borders: createTableBorders(false),
                children: [
                  new Paragraph({
                    children: [
                      new TextRun({
                        text: `作成者: ${creatorName}　　番号: ${solverNumber}　　作成日: ${todayJst}　　解答日: (          /      /      )`,
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
                borders: createTableBorders(false),
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
  });
}

function chunkQuestions(problemExpressions: string[]): string[][] {
  const chunks: string[][] = [];

  for (let index = 0; index < problemExpressions.length; index += QUESTIONS_PER_PAGE) {
    chunks.push(problemExpressions.slice(index, index + QUESTIONS_PER_PAGE));
  }

  return chunks;
}

export function createWorksheetDocument(params: {
  problemExpressions: string[];
  questionCount: number;
  creatorName: string;
  solverNumber: string;
  todayJst: string;
  fontSizePt: number;
}): Document {
  const pageChunks = chunkQuestions(params.problemExpressions);

  return new Document({
    sections: [
      ...pageChunks.map((chunk) => ({
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
          default: createHeader(params.questionCount),
        },
        footers: {
          default: createFooter(params.creatorName, params.solverNumber, params.todayJst),
        },
        children: chunk.map((problemExpression) => createProblemTable(problemExpression, params.fontSizePt)),
      })),
    ],
  });
}