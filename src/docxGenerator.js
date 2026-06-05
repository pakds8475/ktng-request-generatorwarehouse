import {
  AlignmentType,
  BorderStyle,
  Document,
  Packer,
  PageOrientation,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  VerticalAlign,
  VerticalMergeType,
  WidthType,
} from "docx";

function safeValue(value) {
  return value ?? "";
}

export function formatDateDots(value) {
  if (!value) return "";

  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "";

  const day = String(date.getDate()).padStart(2, "0");
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const year = date.getFullYear();

  return `${day}.${month}.${year}`;
}

export function formatDateLong(value) {
  if (!value) return "«__» ________ 2026 г.";

  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "«__» ________ 2026 г.";

  const months = [
    "января",
    "февраля",
    "марта",
    "апреля",
    "мая",
    "июня",
    "июля",
    "августа",
    "сентября",
    "октября",
    "ноября",
    "декабря",
  ];

  const day = String(date.getDate()).padStart(2, "0");
  const month = months[date.getMonth()];
  const year = date.getFullYear();

  return `«${day}» ${month} ${year} г.`;
}

function paragraph(text, options = {}) {
  return new Paragraph({
    alignment: options.alignment ?? AlignmentType.LEFT,
    spacing: options.spacing,
    children: [
      new TextRun({
        text: safeValue(text),
        bold: options.bold ?? false,
        size: options.size ?? 24,
        font: options.font ?? "Times New Roman",
      }),
    ],
  });
}

function cell(text, options = {}) {
  return new TableCell({
    verticalAlign: VerticalAlign.CENTER,
    verticalMerge: options.verticalMerge,
    width: options.width
      ? { size: options.width, type: WidthType.DXA }
      : undefined,
    margins: {
      top: 70,
      bottom: 70,
      left: 70,
      right: 70,
    },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
      left: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
      right: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
    },
    children: [
      new Paragraph({
        alignment: options.alignment ?? AlignmentType.CENTER,
        children: [
          new TextRun({
            text: safeValue(text),
            bold: options.bold ?? false,
            size: options.size ?? 20,
            font: "Times New Roman",
          }),
        ],
      }),
    ],
  });
}

function goodsTableRows(data) {
  const goods =
    data.goods?.filter((item) => item.name || item.weight || item.quantity) ?? [];

  const normalizedGoods =
    goods.length > 0 ? goods : [{ name: "", weight: "", quantity: "" }];

  const dateTime = `${formatDateDots(data.arrivalDate)}${data.timeRange ? ` ${data.timeRange}` : ""}`.trim();

  return normalizedGoods.map((item, index) => {
    const isFirst = index === 0;

    return new TableRow({
      children: [
        cell(isFirst ? "1" : "", {
          width: 700,
          alignment: AlignmentType.CENTER,
          verticalMerge: isFirst
            ? VerticalMergeType.RESTART
            : VerticalMergeType.CONTINUE,
        }),
        cell(isFirst ? dateTime : "", {
          width: 1800,
          alignment: AlignmentType.CENTER,
          verticalMerge: isFirst
            ? VerticalMergeType.RESTART
            : VerticalMergeType.CONTINUE,
        }),
        cell(isFirst ? safeValue(data.vehicleNumber) : "", {
          width: 1300,
          alignment: AlignmentType.CENTER,
          verticalMerge: isFirst
            ? VerticalMergeType.RESTART
            : VerticalMergeType.CONTINUE,
        }),
        cell(isFirst ? safeValue(data.driverName).replaceAll("\n", " ") : "", {
          width: 2100,
          alignment: AlignmentType.CENTER,
          verticalMerge: isFirst
            ? VerticalMergeType.RESTART
            : VerticalMergeType.CONTINUE,
        }),
        cell(safeValue(item.name), {
          width: 4300,
          alignment: AlignmentType.LEFT,
        }),
        cell(safeValue(item.weight), {
          width: 900,
          alignment: AlignmentType.CENTER,
        }),
        cell(safeValue(item.quantity), {
          width: 900,
          alignment: AlignmentType.CENTER,
        }),
      ],
    });
  });
}

export async function downloadDocx(data) {
  const title =
    data.applicationType === "acceptance"
      ? "Заявка на приемку"
      : "Заявка на выдачу товара";

  const doc = new Document({
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: 700,
              right: 700,
              bottom: 700,
              left: 700,
            },
            size: {
              orientation: PageOrientation.LANDSCAPE,
            },
          },
        },
        children: [
          paragraph(`№ от ${formatDateLong(data.docDate)}`, {
            size: 24,
            spacing: { after: 150 },
          }),

          paragraph(title, {
            bold: true,
            size: 28,
            alignment: AlignmentType.CENTER,
            spacing: { after: 150 },
          }),

          paragraph(
            `г. ${safeValue(data.city) || "________"} ${formatDateLong(data.docDate)}`,
            {
              size: 24,
              alignment: AlignmentType.CENTER,
              spacing: { after: 180 },
            }
          ),

          paragraph(`ХРАНИТЕЛЬ-ОПЕРАТОР: ${safeValue(data.operator)}`, {
            size: 24,
            spacing: { after: 100 },
          }),

          paragraph(`ПОКЛАЖЕДАТЕЛЬ-ЗАКАЗЧИК: ${safeValue(data.customer)}`, {
            size: 24,
            spacing: { after: 180 },
          }),

          new Table({
            width: {
              size: 100,
              type: WidthType.PERCENTAGE,
            },
            rows: [
              new TableRow({
                children: [
                  cell("№", {
                    width: 700,
                    bold: true,
                  }),
                  cell("Дата/время", {
                    width: 1800,
                    bold: true,
                  }),
                  cell("№ а/м", {
                    width: 1300,
                    bold: true,
                  }),
                  cell("ФИО водителя", {
                    width: 2100,
                    bold: true,
                  }),
                  cell("Наименование", {
                    width: 4300,
                    bold: true,
                  }),
                  cell("Вес", {
                    width: 900,
                    bold: true,
                  }),
                  cell("Кол-во", {
                    width: 900,
                    bold: true,
                  }),
                ],
              }),
              ...goodsTableRows(data),
            ],
          }),

          paragraph(`Иная информация о грузе: ${safeValue(data.cargoInfo)}`, {
            size: 24,
            spacing: { before: 180, after: 260 },
          }),

          new Table({
            width: {
              size: 100,
              type: WidthType.PERCENTAGE,
            },
            borders: {
              top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
              bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
              left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
              right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
              insideHorizontal: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
              insideVertical: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
            },
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    borders: {
                      top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                      bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                      left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                      right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                    },
                    width: { size: 50, type: WidthType.PERCENTAGE },
                    children: [
                      paragraph("От Хранителя-Оператора:", {
                        size: 24,
                        spacing: { after: 120 },
                      }),
                      paragraph("______________________", {
                        size: 24,
                      }),
                      paragraph(safeValue(data.operator), {
                        size: 24,
                      }),
                      paragraph("М.П.", {
                        size: 24,
                      }),
                    ],
                  }),
                  new TableCell({
                    borders: {
                      top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                      bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                      left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                      right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                    },
                    width: { size: 50, type: WidthType.PERCENTAGE },
                    children: [
                      paragraph("От Поклажедателя-Заказчика:", {
                        size: 24,
                        spacing: { after: 40 },
                      }),
                      paragraph(safeValue(data.signatoryTitle), {
                        size: 24,
                      }),
                      paragraph(safeValue(data.customer), {
                        size: 24,
                      }),
                      paragraph("______________________", {
                        size: 24,
                      }),
                      paragraph("М.П.", {
                        size: 24,
                      }),
                    ],
                  }),
                ],
              }),
            ],
          }),
        ],
      },
    ],
  });

  const blob = await Packer.toBlob(doc);
  const url = URL.createObjectURL(blob);

  const link = document.createElement("a");
  link.href = url;
  link.download = `${title}_${safeValue(data.city) || "Без_города"}.docx`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);

  URL.revokeObjectURL(url);
}
