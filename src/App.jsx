import { useMemo, useState } from "react";
import { downloadDocx, formatDateDots, formatDateLong } from "./docxGenerator.js";

const allowedCities = [
  "Андижан",
  "Самарканд",
  "Ташкент",
  "Фергана",
  "Наманган",
  "Бухара",
];

const sortedCities = [...allowedCities].sort((a, b) => a.localeCompare(b, "ru"));
const cityOperators = {
  "Ташкент": "ООО «RUZ TRADER»",
  "Андижан": "ООО «SIMBA»",
  "Самарканд": "ООО «SIMBA»",
  "Фергана": "ООО «SIMBA»",
  "Наманган": "ООО «SIMBA»",
  "Бухара": "ООО «SIMBA»",
};
const timeOptions = [
  "08:00",
  "08:30",
  "09:00",
  "09:30",
  "10:00",
  "10:30",
  "11:00",
  "11:30",
  "12:00",
  "12:30",
  "13:00",
  "13:30",
  "14:00",
  "14:30",
  "15:00",
  "15:30",
  "16:00",
  "16:30",
  "17:00",
  "17:30",
  "18:00",
  "18:30",
  "19:00",
  "19:30",
  "20:00",
];

const allowedCargoItems = [
  "Кастомная витрина (с освещением)",
  "Кастомная витрина (без освещения)",
  "Монтажные материалы (без освещения)",
  "Торговое оборудование 6C9T (верхняя часть)",
  "Торговое оборудование 6C9T (нижняя часть)",
  "Торговое оборудование 10x6 DG (б/у) рабочий",
  "Торговое оборудование 6x10 METAWERK б/у рабочий",
  "Торговое оборудование 12C10T (верхняя часть) б/у рабочий",
  "Торговое оборудование 12C10T (нижняя часть) б/у рабочий",
  "Торговое оборудование 10C6T б/у рабочий",
  "Торговое оборудование 6x10 DG Брак",
  "Торговое оборудование 6C9T (верхняя часть) б/у рабочий",
  "Торговое оборудование 6C9T (нижняя часть) б/у рабочий",
  "Торговое оборудование 10x10 METAWERK б/у рабочий",
  "Торговое оборудование 10x2 METAWERK б/у рабочий",
  "Торговое оборудование 6C9T (верхняя часть)",
  "Торговое оборудование 6C9T (нижняя часть)",
  "Торговое оборудование 10C9T (верхняя часть) б/у рабочий",
  "Торговое оборудование 10C9T (нижняя часть) б/у рабочий",
  "Кастомная витрина (с освещением) б/у рабочий",
  "Кастомная витрина (без освещения) 3 ячейки (0,5)",
  "Торговое оборудование 10x10 MW б/у рабочий",
  "Кастомная витрина (без освещения) б/у рабочий",
  "Торговое оборудование 2x10 METAWERK б/у рабочий",
  "Торговое оборудование 12C10T (верхняя часть)",
  "Торговое оборудование 12C10T (нижняя часть)",
  "Торговое оборудование 6C10T",
  "Торговое оборудование 10C6T",
  "Торговое оборудование 10C6T (б/у)",
  "Торговое оборудование 10C9T (верхняя часть)",
  "Торговое оборудование 10C9T (нижняя часть)",
  "Торговое оборудование 6x10 METAWERK (б/у) рабочий",
  "Торговое оборудование 6x10 METAWERK (б/у) Брак",
  "Торговое оборудование 10x10 METAWERK (б/у) рабочий",
  "Кастомная витрина (без освещения) (б/у)",
  "Кастомная витрина (с освещения) (б/у)",
  "Торговое оборудование 10C9T (верхняя часть) Б/У",
  "Торговое оборудование 10C9T (нижняя часть) Б/У",
  "Торговое оборудование 10C9T (верхняя часть) Брак",
  "Торговое оборудование 10C9T (нижняя часть) Брак",
  "Торговое оборудование 12C10T (верхняя часть) брак",
  "Торговое оборудование 12C10T (нижняя часть) брак",
  "Торговое оборудование 6C9T (верхняя часть) брак",
  "Торговое оборудование 6C9T (нижняя часть) брак",
  "Торговое оборудование 10x10 METAWERK брак",
  "Торговое оборудование 2x10 METAWERK б/у",
  "Торговое оборудование 6x10 DG б/у рабочий",
  "Лайтбокс (комплект)",
  "Монтажные материалы (с освещением)",
  "Повреждённые акцизные марки",
  "Пластмассовые ценники",
];

const getCargoCondition = (cargoName) => {
  const text = cargoName.toLowerCase();

  if (text.includes("брак")) {
    return "defect";
  }

  if (text.includes("б/у") || text.includes("бу") || text.includes("б.у")) {
    return "used";
  }

  return "new";
};

const normalizeCargoKey = (value) =>
  value
    .trim()
    .replace(/\s+/g, " ")
    .replace(/[Сс]/g, "C")
    .replace(/[Хх]/g, "x")
    .toLowerCase();

const uniqueCargoItems = allowedCargoItems.filter((cargo, index, array) => {
  const currentKey = normalizeCargoKey(cargo);

  return array.findIndex((item) => normalizeCargoKey(item) === currentKey) === index;
});

const getCargoGroup = (cargoName) => {
  const text = normalizeCargoKey(cargoName);

  if (text.includes("12c10t")) {
    return "XL — 12C10T";
  }

  if (text.includes("10x10") || text.includes("10c9t")) {
    return "L — 10x10 / 10C9T";
  }

  if (text.includes("6x10") || text.includes("6c9t") || text.includes("6c9t")) {
    return "M — 6x10 / 6C9T / 6C9T";
  }

  if (text.includes("10x6") || text.includes("10c6t")) {
    return "C — 10x6 / 10C6T";
  }

  if (text.includes("2x10") || text.includes("2x10")) {
    return "S — 2x10 / 10x2";
  }

  if (
    text.includes("витрина") ||
    text.includes("монтажные") ||
    text.includes("лайтбокс") ||
    text.includes("ценники") ||
    text.includes("марки")
  ) {
    return "Прочее — витрины / материалы / POSM";
  }

  return "Другое";
};

const cargoGroupOrder = [
  "XL — 12C10T",
  "L — 10x10 / 10C9T",
  "M — 6x10 / 6C9T / 6C9T",
  "C — 10x6 / 10C6T",
  "S — 2x10 / 10x2",
  "Прочее — витрины / материалы / POSM",
  "Другое",
];

const sortedCargoItems = [...uniqueCargoItems].sort((a, b) => {
  const groupA = getCargoGroup(a);
  const groupB = getCargoGroup(b);

  const groupIndexA = cargoGroupOrder.indexOf(groupA);
  const groupIndexB = cargoGroupOrder.indexOf(groupB);

  if (groupIndexA !== groupIndexB) {
    return groupIndexA - groupIndexB;
  }

  return a.localeCompare(b, "ru");
});

const groupCargoItems = (items) =>
  cargoGroupOrder
    .map((groupName) => ({
      groupName,
      items: items.filter((item) => getCargoGroup(item) === groupName),
    }))
    .filter((group) => group.items.length > 0);

const initialForm = {
  applicationType: "issue",
  city: "",
  docDate: "",
  arrivalDate: "",
  timeRange: "",
  timeStart: "",
timeEnd: "",
  vehicleNumber: "",
  driverName: "",
  operator: "ООО «SIMBA»",
  customer: "ИП OOO «KTNG GLOBAL TAS»",
  signatoryTitle: "Специалист по Логистике",
  cargoInfo: "",
  goods: [{ name: "", weight: "", quantity: "" }],
};

const emptyForm = {
  ...initialForm,
};

function Field({ label, children, hint }) {
  return (
    <label className="field">
      <span>{label}</span>
      {children}
      {hint ? <small>{hint}</small> : null}
    </label>
  );
}

function App() {
  const [form, setForm] = useState(initialForm);
  const [isGenerating, setIsGenerating] = useState(false);
  const [cargoFilter, setCargoFilter] = useState("new");

  const filteredCargoItems = sortedCargoItems.filter((cargo) => {
    if (cargoFilter === "all") {
      return true;
    }

    return getCargoCondition(cargo) === cargoFilter;
  });

  const groupedCargoItems = groupCargoItems(filteredCargoItems);

  const title =
    form.applicationType === "issue"
      ? "Заявка на выдачу товара"
      : "Заявка на приемку";

  const hasRequiredFields =
    form.city &&
    form.docDate &&
    form.arrivalDate &&
    form.timeRange &&
    form.vehicleNumber &&
    form.driverName;

  const previewGoods = useMemo(
    () => form.goods.filter((item) => item.name || item.weight || item.quantity),
    [form.goods]
  );

  function updateField(name, value) {
    setForm((prev) => ({ ...prev, [name]: value }));
  }

  function updateGoods(index, name, value) {
    setForm((prev) => ({
      ...prev,
      goods: prev.goods.map((item, itemIndex) =>
        itemIndex === index ? { ...item, [name]: value } : item
      ),
    }));
  }

  function addGoodsRow() {
    setForm((prev) => ({
      ...prev,
      goods: [...prev.goods, { name: "", weight: "", quantity: "" }],
    }));
  }

  function removeGoodsRow(index) {
    setForm((prev) => ({
      ...prev,
      goods:
        prev.goods.length === 1
          ? prev.goods
          : prev.goods.filter((_, itemIndex) => itemIndex !== index),
    }));
  }

  async function handleDownload() {
    setIsGenerating(true);

    try {
      await downloadDocx(form);
    } finally {
      setIsGenerating(false);
    }
  }

  return (
    <main className="page">
      <section className="hero">
        <div>
          <p className="eyebrow">KT&G Logistics Tool</p>
          <h1>Генератор заявок на выдачу и приемку</h1>
          <p className="heroText">
            Заполните данные автомобиля, водителя и груза. Сайт автоматически
            соберет DOCX-заявку в формате ваших документов.
          </p>
        </div>

        <div className="heroCard">
          <span>Формат</span>
          <strong>DOCX</strong>
          <small>готово для отправки / печати</small>
        </div>
      </section>

      <section className="toolbar">
        <button
          type="button"
          className="secondary"
          onClick={() => setForm(emptyForm)}
        >
          Очистить форму
        </button>
      </section>

      <section className="workspace">
        <form className="panel" onSubmit={(event) => event.preventDefault()}>
          <div className="sectionTitle">
            <h2>Данные заявки</h2>
            <p>Поля ниже попадут в готовый документ.</p>
          </div>

          <div className="requestChoiceBlock">
            <div className="field">
              <span>Тип заявки</span>

              <div className="applicationTypeButtons">
                <button
                  type="button"
                  className={
                    form.applicationType === "issue"
                      ? "issueButton active"
                      : "issueButton"
                  }
                  onClick={() => updateField("applicationType", "issue")}
                >
                  Выдача товара
                </button>

                <button
                  type="button"
                  className={
                    form.applicationType === "acceptance"
                      ? "acceptanceButton active"
                      : "acceptanceButton"
                  }
                  onClick={() => updateField("applicationType", "acceptance")}
                >
                  Приемка
                </button>
              </div>
            </div>

            <div className="field">
              <span>Город</span>

              <div className="cityButtons">
                {sortedCities.map((city) => (
                 <button
  type="button"
  key={city}
  className={form.city === city ? "cityButton active" : "cityButton"}
  onClick={() => {
    updateField("city", city);
    updateField("operator", cityOperators[city] || "");
  }}
>
  {city}
</button>
                ))}
              </div>
            </div>
          </div>

          <div className="grid two">
            <Field label="Дата документа">
              <input
                type="date"
                value={form.docDate}
                onChange={(event) => updateField("docDate", event.target.value)}
              />
            </Field>

           <Field label="Время прихода автомобиля">
  <div className="timeRangeInputs">
    <select
      value={form.timeStart || ""}
      onChange={(event) => {
        const newStart = event.target.value;
        const end = form.timeEnd || "";

        updateField("timeStart", newStart);
        updateField("timeRange", newStart && end ? `${newStart}-${end}` : "");
      }}
    >
      <option value="">С</option>
      {timeOptions.map((time) => (
        <option key={time} value={time}>
          {time}
        </option>
      ))}
    </select>

    <span>—</span>

    <select
      value={form.timeEnd || ""}
      onChange={(event) => {
        const newEnd = event.target.value;
        const start = form.timeStart || "";

        updateField("timeEnd", newEnd);
        updateField("timeRange", start && newEnd ? `${start}-${newEnd}` : "");
      }}
    >
      <option value="">До</option>
      {timeOptions.map((time) => (
        <option key={time} value={time}>
          {time}
        </option>
      ))}
    </select>
  </div>
</Field>


<Field label="Время прихода автомобиля">
  <div className="timeRangeInputs">
    <select
      value={form.timeStart || ""}
      onChange={(event) => {
        const newStart = event.target.value;
        const end = form.timeEnd || "";

        updateField("timeStart", newStart);
        updateField("timeRange", newStart && end ? `${newStart}-${end}` : "");
      }}
    >
      <option value="">С</option>
      {timeOptions.map((time) => (
        <option key={time} value={time}>
          {time}
        </option>
      ))}
    </select>

    <span>—</span>

    <select
      value={form.timeEnd || ""}
      onChange={(event) => {
        const newEnd = event.target.value;
        const start = form.timeStart || "";

        updateField("timeEnd", newEnd);
        updateField("timeRange", start && newEnd ? `${start}-${newEnd}` : "");
      }}
    >
      <option value="">До</option>
      {timeOptions.map((time) => (
        <option key={time} value={time}>
          {time}
        </option>
      ))}
    </select>
  </div>
</Field>

            <Field label="№ автомобиля / контейнера">
              <input
                value={form.vehicleNumber}
                onChange={(event) => updateField("vehicleNumber", event.target.value)}
                placeholder="40 E 366 KB"
              />
            </Field>
          </div>

          <Field
            label="ФИО водителя"
            hint="Можно писать в одну строку или переносить строки как в шаблоне."
          >
            <textarea
              value={form.driverName}
              onChange={(event) => updateField("driverName", event.target.value)}
              rows={3}
              placeholder="ФИО водителя"
            />
          </Field>

          <div className="grid two">
            <Field label="Хранитель-оператор">
              <input
                value={form.operator}
                onChange={(event) => updateField("operator", event.target.value)}
              />
            </Field>

            <Field label="Поклажедатель-заказчик">
              <input
                value={form.customer}
                onChange={(event) => updateField("customer", event.target.value)}
              />
            </Field>
          </div>

          <Field label="Должность подписанта со стороны заказчика">
            <input
              value={form.signatoryTitle}
              onChange={(event) => updateField("signatoryTitle", event.target.value)}
            />
          </Field>

          <div className="sectionTitle goodsTitle">
            <h2>Сведения о грузе</h2>
            <button type="button" onClick={addGoodsRow}>
              + Добавить строку
            </button>
          </div>

          <div className="goodsTable">
            <div className="goodsHeader">
              <span>Наименование</span>
              <span>Вес (кг)</span>
              <span>Кол-во</span>
              <span></span>
            </div>

            <div className="cargoFilters">
   

              <button
                type="button"
                className={cargoFilter === "new" ? "activeFilter" : ""}
                onClick={() => setCargoFilter("new")}
              >
                Новый
              </button>

              <button
                type="button"
                className={cargoFilter === "used" ? "activeFilter" : ""}
                onClick={() => setCargoFilter("used")}
              >
                БУ
              </button>

              <button
                type="button"
                className={cargoFilter === "defect" ? "activeFilter" : ""}
                onClick={() => setCargoFilter("defect")}
              >
                Брак
              </button>
                         <button
                type="button"
                className={cargoFilter === "all" ? "activeFilter" : ""}
                onClick={() => setCargoFilter("all")}
              >
                Все
              </button>
            </div>

            {form.goods.map((item, index) => (
              <div className="goodsRow" key={index}>
                <select
  value={filteredCargoItems.includes(item.name) ? item.name : ""}
  onChange={(event) => updateGoods(index, "name", event.target.value)}
>
  <option value="">Выберите груз</option>

  {groupedCargoItems.map((group) => (
    <optgroup key={group.groupName} label={group.groupName}>
      {group.items.map((cargo) => (
        <option key={cargo} value={cargo}>
          {cargo}
        </option>
      ))}
    </optgroup>
  ))}
</select>

                <input
                  value={item.weight}
                  onChange={(event) => updateGoods(index, "weight", event.target.value)}
                  placeholder="Вес"
                />

                <input
                  value={item.quantity}
                  onChange={(event) => updateGoods(index, "quantity", event.target.value)}
                  placeholder="Кол-во"
                />

                <button
                  type="button"
                  className="iconButton"
                  onClick={() => removeGoodsRow(index)}
                  title="Удалить строку"
                >
                  ×
                </button>
              </div>
            ))}
          </div>

          <Field label="Иная информация о грузе">
            <textarea
              value={form.cargoInfo}
              onChange={(event) => updateField("cargoInfo", event.target.value)}
              rows={2}
              placeholder="Дополнительная информация, если есть"
            />
          </Field>

          <button
            type="button"
            className="download"
            disabled={!hasRequiredFields || isGenerating}
            onClick={handleDownload}
          >
            {isGenerating ? "Создаю DOCX..." : "Скачать DOCX"}
          </button>

          {!hasRequiredFields ? (
            <p className="warning">
              Заполните город, даты, время, номер авто и ФИО водителя.
            </p>
          ) : null}
        </form>

        <aside className="preview">
          <div className="paper">
            <p>
              № от{" "}
              {form.docDate
                ? formatDateLong(form.docDate)
                : "«__» ________ 2026 г."}
            </p>

            <h3>{title}</h3>

            <p className="center">
              г. {form.city || "________"}{" "}
              {form.docDate
                ? formatDateLong(form.docDate)
                : "«__» ________ 2026 г."}
            </p>

            <p>
              <b>ХРАНИТЕЛЬ-ОПЕРАТОР :</b> {form.operator}
            </p>

            <p>
              <b>ПОКЛАЖЕДАТЕЛЬ-ЗАКАЗЧИК:</b> {form.customer}
            </p>

            <table>
              <thead>
                <tr>
                  <th>№</th>
                  <th>Дата/время</th>
                  <th>№ а/м</th>
                  <th>ФИО водителя</th>
                  <th>Наименование</th>
                  <th>Вес</th>
                  <th>Кол-во</th>
                </tr>
              </thead>

              <tbody>
                {(previewGoods.length
                  ? previewGoods
                  : [{ name: "", weight: "", quantity: "" }]
                ).map((item, index) => (
                  <tr key={index}>
                    <td>{index === 0 ? "1" : ""}</td>

                    <td>
                      {index === 0
                        ? `${formatDateDots(form.arrivalDate)} ${form.timeRange}`
                        : ""}
                    </td>

                    <td>{index === 0 ? form.vehicleNumber : ""}</td>

                    <td>
                      {index === 0 ? form.driverName.replaceAll("\n", " ") : ""}
                    </td>

                    <td>{item.name}</td>
                    <td>{item.weight}</td>
                    <td>{item.quantity}</td>
                  </tr>
                ))}
              </tbody>
            </table>

            <p>Иная информация о грузе: {form.cargoInfo}</p>

            <div className="signatures">
              <span>
                От Хранителя-Оператора:
                <br />
                ______________________
                <br />
                {form.operator}
                <br />
                М.П.
              </span>

              <span>
                От Поклажедателя-Заказчика:
                <br />
                {form.signatoryTitle}
                <br />
                {form.customer}
                <br />
                ___________________
                <br />
                М.П.
              </span>
            </div>
          </div>
        </aside>
      </section>
    </main>
  );
}

export default App;
