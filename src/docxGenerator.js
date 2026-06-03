import {
  AlignmentType,
  BorderStyle,
  Document,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  VerticalAlign,
  VerticalMergeType,
  WidthType,
} from "docx";

const MONTHS_RU = [
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

const noBorder = {
  top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
};

const tableBorders = {
  top: { style: BorderStyle.SINGLE, size: 8, color: "000000" },
  bottom: { style: BorderStyle.SINGLE, size: 8, color: "000000" },
  left: { style: BorderStyle.SINGLE, size: 8, color: "000000" },
  right: { style: BorderStyle.SINGLE, size: 8, color: "000000" },
  insideHorizontal: { style: BorderStyle.SINGLE, size: 8, color: "000000" },
  insideVertical: { style: BorderStyle.SINGLE, size: 8, color: "000000" },
};

function parseDate(dateString) {
  const [year, month, day] = dateString.split("-").map(Number);
  return new Date(year, month - 1, day);
}

export function formatDateDots(dateString) {
  const date = parseDate(dateString);
  const day = String(date.getDate()).padStart(2, "0");
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const year = date.getFullYear();
  return `${day}.${month}.${year}`;
}

export function formatDateLong(dateString) {
  const date = parseDate(dateString);
  const day = String(date.getDate()).padStart(2, "0");
  const month = MONTHS_RU[date.getMonth()];
  const year = date.getFullYear();
  return `«${day}» ${month} ${year} г.`;
}

function run(text, options = {}) {
  return new TextRun({
    text: String(text ?? ""),
    font: "Times New Roman",
    size: options.size ?? 24,
    bold: options.bold ?? false,
    italics: options.italics ?? false,
  });
}

function paragraph(text, options = {}) {
  const lines = String(text ?? "").split("\n");
  const children = [];

  lines.forEach((line, index) => {
    if (index > 0) {
      children.push(new TextRun({ break: 1 }));
    }
    children.push(run(line, options));
  });

  return new Paragraph({
    children,
    alignment: options.alignment ?? AlignmentType.LEFT,
    spacing: options.spacing ?? { after: 80 },
  });
}

function cell(text, options = {}) {
  const align = options.align ?? AlignmentType.CENTER;
  return new TableCell({
    width: options.width ? { size: options.width, type: WidthType.DXA } : undefined,
    columnSpan: options.columnSpan,
    verticalMerge: options.verticalMerge,
    verticalAlign: options.verticalAlign ?? VerticalAlign.CENTER,
    margins: {
      top: 80,
      bottom: 80,
      left: 80,
      right: 80,
    },
    children: [
      paragraph(text, {
        alignment: align,
        size: options.size ?? 20,
        bold: options.bold ?? false,
        spacing: { after: 0 },
      }),
    ],
  });
}

function signatureCell(text) {
  return new TableCell({
    borders: noBorder,
    margins: { top: 120, bottom: 120, left: 0, right: 0 },
    children: text.split("\n").map((line) =>
      paragraph(line, {
        spacing: { after: 60 },
        size: 22,
      })
    ),
  });
}

function createCargoTable(data) {
  const goods = data.goods.filter((item) => item.name || item.weight || item.quantity);
  const safeGoods = goods.length ? goods : [{ name: "", weight: "", quantity: "" }];
  const dateTime = `${formatDateDots(data.arrivalDate)}\n${data.timeRange}`;

  const rows = [
    new TableRow({
      tableHeader: true,
      children: [
        cell("№\nп/п", { width: 700, verticalMerge: VerticalMergeType.RESTART, bold: true }),
        cell("Дата/время\nприхода а/м", { width: 1700, verticalMerge: VerticalMergeType.RESTART, bold: true }),
        cell("№ а/м,\nконтейнера", { width: 1600, verticalMerge: VerticalMergeType.RESTART, bold: true }),
        cell("ФИО водителя", { width: 1900, verticalMerge: VerticalMergeType.RESTART, bold: true }),
        cell("Сведения о грузе", { columnSpan: 3, bold: true }),
      ],
    }),
    new TableRow({
      tableHeader: true,
      children: [
        cell("", { verticalMerge: VerticalMergeType.CONTINUE }),
        cell("", { verticalMerge: VerticalMergeType.CONTINUE }),
        cell("", { verticalMerge: VerticalMergeType.CONTINUE }),
        cell("", { verticalMerge: VerticalMergeType.CONTINUE }),
        cell("Наименование", { width: 3100, bold: true }),
        cell("Вес (кг)", { width: 900, bold: true }),
        cell("Кол-во,\nединиц", { width: 1100, bold: true }),
      ],
    }),
  ];

  safeGoods.forEach((item, index) => {
    const isFirst = index === 0;
    rows.push(
      new TableRow({
        children: [
          cell(isFirst ? "1" : "", {
            verticalMerge: isFirst ? VerticalMergeType.RESTART : VerticalMergeType.CONTINUE,
            width: 700,
          }),
          cell(isFirst ? dateTime : "", {
            verticalMerge: isFirst ? VerticalMergeType.RESTART : VerticalMergeType.CONTINUE,
            width: 1700,
          }),
          cell(isFirst ? data.vehicleNumber : "", {
            verticalMerge: isFirst ? VerticalMergeType.RESTART : VerticalMergeType.CONTINUE,
            width: 1600,
          }),
          cell(isFirst ? data.driverName : "", {
            verticalMerge: isFirst ? VerticalMergeType.RESTART : VerticalMergeType.CONTINUE,
            width: 1900,
          }),
          cell(item.name, { align: AlignmentType.LEFT, width: 3100 }),
          cell(item.weight, { width: 900 }),
          cell(item.quantity, { width: 1100 }),
        ],
      })
    );
  });

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: tableBorders,
    rows,
  });
}

export function buildDocx(data) {
  const title = data.applicationType === "issue" ? "Заявка на выдачу товара" : "Заявка на приемку";
  const longDate = formatDateLong(data.docDate);

  const doc = new Document({
    creator: "KT&G Request Generator",
    title,
    description: "Автоматически созданная заявка",
    styles: {
      default: {
        document: {
          run: { font: "Times New Roman", size: 24 },
          paragraph: { spacing: { after: 80 } },
        },
      },
    },
    sections: [
      {
        properties: {
          page: {
            margin: { top: 720, right: 720, bottom: 720, left: 720 },
          },
        },
        children: [
          paragraph(`№ от ${longDate}`, { size: 22 }),
          paragraph("", { spacing: { after: 120 } }),
          paragraph(title, { alignment: AlignmentType.CENTER, bold: true, size: 28, spacing: { after: 120 } }),
          paragraph(`г. ${data.city} ${longDate}`, { alignment: AlignmentType.CENTER, size: 24, spacing: { after: 180 } }),
          paragraph(`ХРАНИТЕЛЬ-ОПЕРАТОР : ${data.operator}`, { size: 22, spacing: { after: 80 } }),
          paragraph(`ПОКЛАЖЕДАТЕЛЬ-ЗАКАЗЧИК: ${data.customer}`, { size: 22, spacing: { after: 160 } }),
          createCargoTable(data),
          paragraph("", { spacing: { after: 160 } }),
          paragraph(`Иная информация о грузе:${data.cargoInfo ? ` ${data.cargoInfo}` : ""}`, { size: 22, spacing: { after: 180 } }),
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            borders: noBorder,
            rows: [
              new TableRow({
                children: [
                  signatureCell(`От Хранителя-Оператора:\n______________________\n${data.operator}\n________________________\nМ.П.`),
                  signatureCell(`От Поклажедателя-Заказчика:\n${data.signatoryTitle}\n${data.customer}\n___________________\nМ.П.`),
                ],
              }),
            ],
          }),
        ],
      },
    ],
  });

  return doc;
}

export async function downloadDocx(data) {
  const doc = buildDocx(data);
  const blob = await Packer.toBlob(doc);
  const fileNameType = data.applicationType === "issue" ? "выдача" : "приемка";
  const fileName = `Заявка_${fileNameType}_${data.city}_${formatDateDots(data.docDate)}.docx`;

  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(link.href);
}
