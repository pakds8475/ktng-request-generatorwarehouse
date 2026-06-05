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

function parseDate(value) {
  if (!value) return null;

  const parts = String(value).split("-");
  if (parts.length !== 3) return null;

  const year = Number(parts[0]);
  const month = Number(parts[1]);
  const day = Number(parts[2]);

  if (!year || !month || !day) return null;

  return {
    year,
    month,
    day,
  };
}

export function formatDateDots(value) {
  const parsed = parseDate(value);

  if (!parsed) return "";

  const day = String(parsed.day).padStart(2, "0");
  const month = String(parsed.month).padStart(2, "0");

  return `${day}.${month}.${parsed.year}`;
}

export function formatDateLong(value) {
  const parsed = parseDate(value);

  if (!parsed) return "«__» ________ 2026 г.";

  const months = [
    "Января",
    "Февраля",
    "Марта",
    "Апреля",
    "Мая",
    "Июня",
    "Июля",
    "Августа",
    "Сентября",
    "Октября",
    "Ноября",
    "Декабря",
  ];

  const day = String(parsed.day).padStart(2, "0");
  const monthName = months[parsed.month - 1];

  return `«${day}» ${monthName} ${parsed.year} г.`;
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
        font: "Times New Roman",
      }),
    ],
  });
}

function tableText(text, options = {}) {
  return new Paragraph({
    alignment: options.alignment ?? AlignmentType.CENTER,
    spacing: {
      before: 0,
      after: 0,
    },
    children: [
      new TextRun({
        text: safeValue(text),
        bold: options.bold ?? false,
        size: options.size ?? 18,
        font: "Times New Roman",
      }),
    ],
  });
}

function cell(text, options = {}) {
  return new TableCell({
    verticalAlign: VerticalAlign.CENTER,
    verticalMerge: options.verticalMerge,
    columnSpan: options.columnSpan,
    width: options.width
      ? {
          size: options.width,
          type: WidthType.DXA,
        }
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
      tableText(text, {
        alignment: options.alignment ?? AlignmentType.CENTER,
        bold: options.bold ?? false,
        size: options.size ?? 18,
      }),
    ],
  });
}

function borderlessCell(children, options = {}) {
  return new TableCell({
    width: options.width
      ? {
          size: options.width,
          type: WidthType.PERCENTAGE,
        }
      : undefined,
    margins: {
      top: 80,
      bottom: 80,
      left: 80,
      right: 80,
    },
    borders: {
      top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
      bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
      left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
      right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    },
    children,
  });
}

function buildDateTime(data) {
  const date = formatDateDots(data.arrivalDate);
  const time = safeValue(data.timeRange);

  return [date, time].filter(Boolean).join(" ");
}

function createGoodsRows(data) {
  const goods =
    data.goods?.filter((item) => item.name || item.weight || item.quantity) ?? [];

  const normalizedGoods =
    goods.length > 0 ? goods : [{ name: "", weight: "", quantity: "" }];

  const dateTime = buildDateTime(data);
  const vehicleNumber = safeValue(data.vehicleNumber);
  const driverName = safeValue(data.driverName).replaceAll("\n", " ");

  return normalizedGoods.map((item, index) => {
    const isFirst = index === 0;

    return new TableRow({
      children: [
        cell(isFirst ? "1" : "", {
          width: 600,
          verticalMerge: isFirst
            ? VerticalMergeType.RESTART
            : VerticalMergeType.CONTINUE,
        }),

        cell(isFirst ? dateTime : "", {
          width: 1800,
          verticalMerge: isFirst
            ? VerticalMergeType.RESTART
            : VerticalMergeType.CONTINUE,
        }),

        cell(isFirst ? vehicleNumber : "", {
          width: 1300,
          verticalMerge: isFirst
            ? VerticalMergeType.RESTART
            : VerticalMergeType.CONTINUE,
        }),

        cell(isFirst ? driverName : "", {
          width: 2300,
          verticalMerge: isFirst
            ? VerticalMergeType.RESTART
            : VerticalMergeType.CONTINUE,
        }),

cell(safeValue(item.name), {
  width: 5200,
  alignment: AlignmentType.CENTER,
}),

        cell(safeValue(item.weight), {
          width: 900,
        }),

        cell(safeValue(item.quantity), {
          width: 1000,
        }),
      ],
    });
  });
}

function createGoodsTable(data) {
  return new Table({
    width: {
      size: 100,
      type: WidthType.PERCENTAGE,
    },
    rows: [
      new TableRow({
        children: [
          cell("№\nп/п", {
            width: 600,
            bold: true,
            verticalMerge: VerticalMergeType.RESTART,
          }),

          cell("Дата/время\nприхода а/м", {
            width: 1800,
            bold: true,
            verticalMerge: VerticalMergeType.RESTART,
          }),

          cell("№ а/м,\nконтейнера", {
            width: 1300,
            bold: true,
            verticalMerge: VerticalMergeType.RESTART,
          }),

          cell("ФИО\nводителя", {
            width: 2300,
            bold: true,
            verticalMerge: VerticalMergeType.RESTART,
          }),

          cell("Сведения о грузе", {
            width: 7100,
            bold: true,
            columnSpan: 3,
          }),
        ],
      }),

      new TableRow({
        children: [
          cell("", {
            width: 600,
            verticalMerge: VerticalMergeType.CONTINUE,
          }),

          cell("", {
            width: 1800,
            verticalMerge: VerticalMergeType.CONTINUE,
          }),

          cell("", {
            width: 1300,
            verticalMerge: VerticalMergeType.CONTINUE,
          }),

          cell("", {
            width: 2300,
            verticalMerge: VerticalMergeType.CONTINUE,
          }),

          cell("Наименование", {
            width: 5200,
            bold: true,
          }),

          cell("Вес (кг)", {
            width: 900,
            bold: true,
          }),

          cell("Кол-во,\nединиц", {
            width: 1000,
            bold: true,
          }),
        ],
      }),

      ...createGoodsRows(data),
    ],
  });
}

function createSignatureTable(data) {
  return new Table({
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
          borderlessCell(
            [
              paragraph("От Хранителя-Оператора:", {
                size: 24,
                spacing: { after: 160 },
              }),
              paragraph("______________________", {
                size: 24,
              }),
              paragraph(safeValue(data.operator), {
                size: 24,
              }),
              paragraph("________________________", {
                size: 24,
              }),
              paragraph("М.П.", {
                size: 24,
              }),
            ],
            {
              width: 50,
            }
          ),

          borderlessCell(
            [
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
              paragraph("____________________________", {
                size: 24,
              }),
              paragraph("М.П.", {
                size: 24,
              }),
            ],
            {
              width: 50,
            }
          ),
        ],
      }),
    ],
  });
}

export async function downloadDocx(data) {
  const title =
    data.applicationType === "acceptance"
      ? "Заявка на приемку"
      : "Заявка на выдачу товара";

  const city = safeValue(data.city) || "________";
  const operator = safeValue(data.operator);
  const customer = safeValue(data.customer);

  const doc = new Document({
    sections: [
      {
        properties: {
          page: {
            size: {
              orientation: PageOrientation.LANDSCAPE,
            },
            margin: {
              top: 600,
              right: 600,
              bottom: 600,
              left: 600,
            },
          },
        },
        children: [
          paragraph(`№ от ${formatDateLong(data.docDate)}`, {
            size: 24,
            spacing: { after: 120 },
          }),

          paragraph(title, {
            bold: true,
            size: 28,
            alignment: AlignmentType.CENTER,
            spacing: { after: 120 },
          }),

          paragraph(`г. ${city} ${formatDateLong(data.docDate)}`, {
            size: 24,
            alignment: AlignmentType.CENTER,
            spacing: { after: 160 },
          }),

          paragraph(`ХРАНИТЕЛЬ-ОПЕРАТОР: ${operator}`, {
            bold: true,
            size: 24,
            spacing: { after: 80 },
          }),

          paragraph(`ПОКЛАЖЕДАТЕЛЬ-ЗАКАЗЧИК: ${customer}`, {
            bold: true,
            size: 24,
            spacing: { after: 160 },
          }),

          createGoodsTable(data),

          paragraph(`Иная информация о грузе: ${safeValue(data.cargoInfo)}`, {
            size: 24,
            spacing: { before: 180, after: 260 },
          }),

          createSignatureTable(data),
        ],
      },
    ],
  });

  const blob = await Packer.toBlob(doc);
  const url = URL.createObjectURL(blob);

  const fileNameCity = city.replaceAll(" ", "_");
  const fileNameDate = formatDateDots(data.docDate).replaceAll(".", "-");
  const fileName = `${title}_${fileNameCity}_${fileNameDate || "без_даты"}.docx`;

  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);

  URL.revokeObjectURL(url);
}
