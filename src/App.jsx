import { useMemo, useState } from "react";
import { downloadDocx, formatDateDots, formatDateLong } from "./docxGenerator.js";

const today = "2026-06-03";

const allowedCities = [
  "Андижан",
  "Самарканд",
  "Ташкент",
  "Фергана",
  "Наманган",
  "Бухара",
];

const sortedCities = [...allowedCities].sort((a, b) => a.localeCompare(b, "ru"));

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

const sortedCargoItems = [...allowedCargoItems].sort((a, b) =>
  a.localeCompare(b, "ru")
);

const initialForm = {
  applicationType: "issue",
  city: "Андижан",
  docDate: today,
  arrivalDate: today,
  timeRange: "14:00-16:00",
  vehicleNumber: "40 E 366 KB",
  driverName: "KADIROV UMID\nABDUVAXOBOVICH",
  operator: "ООО «SIMBA»",
  customer: "ИП OOO «KTNG GLOBAL TAS»",
  signatoryTitle: "Специалист по Логистике",
  cargoInfo: "",
  goods: [
    { name: "Торговое оборудование 6C9T (верхняя часть)", weight: "", quantity: "2" },
    { name: "Торговое оборудование 6C9T (нижняя часть)", weight: "", quantity: "2" },
    { name: "Кастомная витрина (без освещения)", weight: "", quantity: "2" },
    { name: "Монтажные материалы (без освещения)", weight: "", quantity: "2" },
  ],
};

const samarkandExample = {
  applicationType: "acceptance",
  city: "Самарканд",
  docDate: today,
  arrivalDate: today,
  timeRange: "14:00-16:00",
  vehicleNumber: "30 M 463 SB",
  driverName: "XUSAINOV\nAMRIDDIN\nSADRIDDINOVICH",
  operator: "ООО «SIMBA»",
  customer: "ИП OOO «KTNG GLOBAL TAS»",
  signatoryTitle: "Специалист по Логистике",
  cargoInfo: "",
  goods: [
    { name: "Торговое оборудование 10x2 METAWERK б/у рабочий", weight: "", quantity: "1" },
  ],
};

const emptyForm = {
  ...initialForm,
  city: "",
  vehicleNumber: "",
  driverName: "",
  cargoInfo: "",
  goods: [{ name: "", weight: "", quantity: "" }],
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
  const [cargoFilter, setCargoFilter] = useState("all");

  const filteredCargoItems = sortedCargoItems.filter((cargo) => {
    if (cargoFilter === "all") {
      return true;
    }

    return getCargoCondition(cargo) === cargoFilter;
  });

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
        <button type="button" onClick={() => setForm(initialForm)}>
          Пример: Андижан выдача
        </button>

        <button type="button" onClick={() => setForm(samarkandExample)}>
          Пример: Самарканд приемка
        </button>

        <button type="button" className="secondary" onClick={() => setForm(emptyForm)}>
          Очистить форму
        </button>
      </section>

      <section className="workspace">
        <form className="panel" onSubmit={(event) => event.preventDefault()}>
          <div className="sectionTitle">
            <h2>Данные заявки</h2>
            <p>Поля ниже попадут в готовый документ.</p>
          </div>

          <div className="grid two">
            <Field label="Тип заявки">
              <select
                value={form.applicationType}
                onChange={(event) => updateField("applicationType", event.target.value)}
              >
                <option value="issue">Заявка на выдачу товара</option>
                <option value="acceptance">Заявка на приемку</option>
              </select>
            </Field>

            <Field label="Город">
              <select
                value={form.city}
                onChange={(event) => updateField("city", event.target.value)}
              >
                <option value="">Выберите город</option>
                {sortedCities.map((city) => (
                  <option key={city} value={city}>
                    {city}
                  </option>
                ))}
              </select>
            </Field>

            <Field label="Дата документа">
              <input
                type="date"
                value={form.docDate}
                onChange={(event) => updateField("docDate", event.target.value)}
              />
            </Field>

            <Field label="Дата прихода автомобиля">
              <input
                type="date"
                value={form.arrivalDate}
                onChange={(event) => updateField("arrivalDate", event.target.value)}
              />
            </Field>

            <Field label="Время прихода автомобиля">
              <input
                value={form.timeRange}
                onChange={(event) => updateField("timeRange", event.target.value)}
                placeholder="14:00-16:00"
              />
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
                className={cargoFilter === "all" ? "activeFilter" : ""}
                onClick={() => setCargoFilter("all")}
              >
                Все
              </button>

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
            </div>

            {form.goods.map((item, index) => (
              <div className="goodsRow" key={index}>
                <select
                  value={item.name}
                  onChange={(event) => updateGoods(index, "name", event.target.value)}
                >
                  <option value="">Выберите груз</option>

                  {item.name && !filteredCargoItems.includes(item.name) ? (
                    <option value={item.name}>{item.name}</option>
                  ) : null}

                  {filteredCargoItems.map((cargo) => (
                    <option key={cargo} value={cargo}>
                      {cargo}
                    </option>
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
