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
import {
  AVAILABLE_HEIGHT_PT,
  AVAILABLE_WIDTH_PT,
  A4_PAGE_HEIGHT_TWIPS,
  A4_PAGE_WIDTH_TWIPS,
  BORDER_NONE_COLOR,
  BORDER_NONE_SIZE,
  FOOTER_INFO_CELL_WIDTH_PERCENT,
  FOOTER_PAGE_CELL_WIDTH_PERCENT,
  FOOTER_TEXT_SIZE_HALF_POINTS,
  FULL_PERCENT,
  HEADER_TEXT_SIZE_HALF_POINTS,
  MIN_FONT_SIZE_PT,
  PAGE_EDGE_MARGIN_TWIPS,
  PAGE_FOOTER_MARGIN_TWIPS,
  PAGE_HEIGHT_SAFETY_MARGIN_PT,
  PAGE_HEADER_MARGIN_TWIPS,
  PROBLEM_LINE_PADDING_CHARS,
  PROBLEM_TEXT_SCALE,
  PROBLEM_VERTICAL_SPAN,
  FIRST_INDEX,
  QUESTIONS_PER_PAGE,
  TWIPS_PER_POINT,
} from "@/App";
import { toZenkaku } from "@/features/worksheet/worksheet";

const FULL_PAGE_END_PARAGRAPH_RESERVE_PT = 16;

function createNoBorder() {
  return {
    style: BorderStyle.NONE,
    size: BORDER_NONE_SIZE,
    color: BORDER_NONE_COLOR,
  };
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

function calculateWorksheetRowHeight(questionCount: number): number {
  const questionsOnPage = Math.max(
    1,
    Math.min(questionCount, QUESTIONS_PER_PAGE),
  );
  const fullPageReservePt =
    questionsOnPage === QUESTIONS_PER_PAGE
      ? FULL_PAGE_END_PARAGRAPH_RESERVE_PT
      : 0;
  const usableHeightPt =
    AVAILABLE_HEIGHT_PT - PAGE_HEIGHT_SAFETY_MARGIN_PT - fullPageReservePt;

  return Math.floor(usableHeightPt / questionsOnPage);
}

export function calculateWorksheetFontSize(
  problemExpressions: string[],
  questionCount: number,
): number {
  const rowHeightPt = calculateWorksheetRowHeight(questionCount);

  let fontSizePt = Math.floor(rowHeightPt / PROBLEM_VERTICAL_SPAN);
  const maxCalcLength = Math.max(
    ...problemExpressions.map((expression) => expression.length),
  );
  const maxLineLength = maxCalcLength + PROBLEM_LINE_PADDING_CHARS;
  const maxFontByWidth = Math.floor(AVAILABLE_WIDTH_PT / maxLineLength);

  if (fontSizePt > maxFontByWidth) fontSizePt = maxFontByWidth;
  if (fontSizePt < MIN_FONT_SIZE_PT) fontSizePt = MIN_FONT_SIZE_PT;

  return fontSizePt;
}

function createProblemTable(
  problemExpression: string,
  fontSizePt: number,
  rowHeightPt: number,
): Table {
  return new Table({
    borders: createTableBorders(),
    rows: [
      new TableRow({
        height: {
          value: rowHeightPt * TWIPS_PER_POINT,
          rule: HeightRule.EXACT,
        },
        children: [
          new TableCell({
            verticalAlign: VerticalAlign.CENTER,
            children: [
              new Paragraph({
                alignment: AlignmentType.DISTRIBUTE,
                children: [
                  new TextRun({
                    text: toZenkaku(problemExpression),
                    size: fontSizePt * PROBLEM_TEXT_SCALE,
                  }),
                ],
              }),
            ],
          }),
          new TableCell({
            verticalAlign: VerticalAlign.CENTER,
            children: [
              new Paragraph({
                alignment: AlignmentType.RIGHT,
                children: [
                  new TextRun({
                    text: " ＝",
                    size: fontSizePt * PROBLEM_TEXT_SCALE,
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
            text: `脳トレ用　計算問題　${questionCount.toString()}問　　（　　/${questionCount.toString()}）`,
            size: HEADER_TEXT_SIZE_HALF_POINTS,
          }),
        ],
      }),
    ],
  });
}

function createFooter(
  creatorName: string,
  solverNumber: string,
  todayJst: string,
): Footer {
  return new Footer({
    children: [
      new Table({
        borders: createTableBorders(),
        width: {
          size: FULL_PERCENT,
          type: WidthType.PERCENTAGE,
        },
        rows: [
          new TableRow({
            children: [
              new TableCell({
                width: {
                  size: FOOTER_INFO_CELL_WIDTH_PERCENT,
                  type: WidthType.PERCENTAGE,
                },
                verticalAlign: VerticalAlign.BOTTOM,
                borders: createTableBorders(false),
                children: [
                  new Paragraph({
                    children: [
                      new TextRun({
                        text: `作成者: ${creatorName}　　番号: ${solverNumber}　　作成日: ${todayJst}　　解答日: (          /      /      )`,
                        size: FOOTER_TEXT_SIZE_HALF_POINTS,
                      }),
                    ],
                  }),
                ],
              }),
              new TableCell({
                width: {
                  size: FOOTER_PAGE_CELL_WIDTH_PERCENT,
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
                        size: FOOTER_TEXT_SIZE_HALF_POINTS,
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

  for (
    let index = FIRST_INDEX;
    index < problemExpressions.length;
    index += QUESTIONS_PER_PAGE
  ) {
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
}): Document {
  const pageChunks = chunkQuestions(params.problemExpressions);
  const fontSizePt = calculateWorksheetFontSize(
    params.problemExpressions,
    params.questionCount,
  );

  return new Document({
    sections: [
      ...pageChunks.map((chunk) => {
        const rowHeightPt = calculateWorksheetRowHeight(chunk.length);

        return {
          properties: {
            page: {
              size: {
                width: A4_PAGE_WIDTH_TWIPS,
                height: A4_PAGE_HEIGHT_TWIPS,
                orientation: PageOrientation.PORTRAIT,
              },
              margin: {
                header: PAGE_HEADER_MARGIN_TWIPS,
                footer: PAGE_FOOTER_MARGIN_TWIPS,
                top: PAGE_EDGE_MARGIN_TWIPS,
                bottom: PAGE_EDGE_MARGIN_TWIPS,
                left: PAGE_EDGE_MARGIN_TWIPS,
                right: PAGE_EDGE_MARGIN_TWIPS,
              },
            },
          },
          headers: {
            default: createHeader(params.questionCount),
          },
          footers: {
            default: createFooter(
              params.creatorName,
              params.solverNumber,
              params.todayJst,
            ),
          },
          children: chunk.map((problemExpression) =>
            createProblemTable(problemExpression, fontSizePt, rowHeightPt),
          ),
        };
      }),
    ],
  });
}
