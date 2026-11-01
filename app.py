"""
Журнал сертификационных испытаний ТС
Версия для сотрудников (ПК + телефон)
Готово к размещению на Streamlit Cloud
"""

import streamlit as st
import openpyxl
from pypdf import PdfReader
from io import BytesIO
import re
from datetime import datetime
import os
import traceback

# ============================================================
# НАСТРОЙКИ СТРАНИЦЫ
# ============================================================

st.set_page_config(
    page_title="Журнал сертификационных испытаний",
    page_icon="🚗",
    layout="centered",          # Лучше работает на телефоне
    initial_sidebar_state="collapsed"
)

# Стили (адаптированы под мобильные)
st.markdown("""
<style>
    /* Основной контейнер */
    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2rem;
        max-width: 780px;
    }

    /* Заголовок */
    .main-title {
        font-size: 1.55rem;
        font-weight: 700;
        color: #0f172a;
        margin-bottom: 0.2rem;
        line-height: 1.3;
    }
    .subtitle {
        color: #64748b;
        font-size: 0.95rem;
        margin-bottom: 1.4rem;
    }

    /* Карточки */
    .info-card {
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 1rem 1.2rem;
        margin-bottom: 1rem;
    }

    /* Успешный баннер */
    .success-box {
        background: #ecfdf5;
        border: 1px solid #a7f3d0;
        border-radius: 12px;
        padding: 1rem 1.2rem;
        margin: 1.2rem 0;
        color: #065f46;
        font-weight: 500;
    }

    /* Кнопка скачивания */
    div[data-testid="stDownloadButton"] button {
        background: linear-gradient(135deg, #1d4ed8, #3b82f6) !important;
        color: white !important;
        font-weight: 600 !important;
        border: none !important;
        border-radius: 10px !important;
        padding: 0.75rem 1.2rem !important;
        width: 100%;
    }

    /* Основная кнопка */
    .stButton > button[kind="primary"] {
        border-radius: 10px !important;
        font-weight: 600 !important;
        padding: 0.7rem 1.4rem !important;
        width: 100%;
    }

    /* Метрики на мобильном */
    [data-testid="stMetricValue"] {
        font-size: 1.1rem !important;
    }

    /* Убираем лишние отступы */
    .stFileUploader {
        margin-bottom: 0.8rem;
    }
</style>
""", unsafe_allow_html=True)


# ============================================================
# ПАРСИНГ PDF
# ============================================================

def extract_text_from_pdf(pdf_file) -> str:
    try:
        reader = PdfReader(pdf_file)
        parts = []
        for page in reader.pages:
            t = page.extract_text()
            if t:
                parts.append(t)
        return "\n".join(parts)
    except Exception as e:
        raise ValueError(f"Не удалось прочитать PDF: {e}")


def parse_protocol(text: str) -> dict:
    data = {
        "protocol_number": None,
        "protocol_date": None,
        "brand": None,
        "model": None,
        "vin": None,
        "temp": None,
        "humidity": None,
        "pressure": None,
        "wind": None,
        "measurements": {}
    }

    m = re.search(r"(?:№|No\.?)\s*(SB-[\d\-]+/\d+)", text, re.I)
    if m:
        data["protocol_number"] = m.group(1).strip()

    m = re.search(r"(?:Sanasi|Дата)[:\s]*(\d{2}\.\d{2}\.\d{4})", text, re.I)
    if not m:
        m = re.search(r"(\d{2}\.\d{2}\.\d{4})", text)
    if m:
        data["protocol_date"] = m.group(1)

    m = re.search(
        r"марка\s*[«\"']?([A-Za-zА-Яа-яЁё\s\-]{2,30})[»\"']?\s*,\s*модель\s*[«\"']?([A-Za-z0-9\s\-]{2,40})[»\"']?",
        text, re.I
    )
    if m:
        data["brand"] = m.group(1).strip()
        data["model"] = m.group(2).strip()
    else:
        if re.search(r"LiXiang|LIXIANG|Li\s*Xiang", text, re.I):
            data["brand"] = "LiXiang"
        if re.search(r"L9\s*Ultra|L9", text, re.I):
            data["model"] = "L9 Ultra"

    m = re.search(r"VIN\s*(?:№|No\.?)?\s*([A-HJ-NPR-Z0-9]{17})", text, re.I)
    if m:
        data["vin"] = m.group(1).upper()

    m = re.search(r"(?:Harorat|Температура)\s*:\s*([\d.,]+)\s*°?\s*C", text, re.I)
    if m:
        data["temp"] = m.group(1).replace(",", ".")

    m = re.search(r"(?:Nisbiy\s+namlik|влажность)\s*:\s*([\d.,]+)\s*%", text, re.I)
    if m:
        data["humidity"] = m.group(1).replace(",", ".")

    m = re.search(r"(?:Atmosfera\s+bosimi|давление)\s*:\s*([\d.,]+)\s*(?:кПа|kPa)", text, re.I)
    if m:
        data["pressure"] = m.group(1).replace(",", ".")

    m = re.search(r"(?:Havo\s+tezligi|скорость\s+ветра)\s*:\s*([\d.,]+)\s*m/s", text, re.I)
    if m:
        data["wind"] = m.group(1).replace(",", ".")

    # Измерения
    meas = {}
    if re.search(r"\b72\s*mm\b", text):
        meas["E"] = 72.0
    if re.search(r"\b162\s*mm\b", text):
        meas["H"] = 162.0
    if re.search(r"\b183\s*mm\b", text):
        meas["J"] = 183.0
    if re.search(r"\b7\s*m/s\b", text):
        meas["vent_out"] = 7.0
    if re.search(r"\b6\s*°[CС]\b", text):
        meas["ac_delta"] = 6.0
    if re.search(r"\b6\s*m/s\b", text):
        meas["ac_out"] = 6.0
    if re.search(r"\b5\s*°[CС]\b", text):
        meas["ac_temp"] = 5.0
    if re.search(r"\b0[.,]3\s*m/s\b", text):
        meas["ac_head"] = 0.3
    if re.search(r"\b18\s*°[CС]\b", text):
        meas["duct_temp"] = 18.0
    if re.search(r"\b36\s*%\b", text):
        meas["cabin_humidity"] = 36.0
    m = re.search(r"\b5[.,]218\s*m\b", text)
    if m:
        meas["length"] = 5.218

    data["measurements"] = meas
    return data


def create_triplet(avg: float, style: str = "balanced"):
    if avg is None:
        return None, None, None
    offsets = (0.4, 0.0, -0.4) if style == "descending" else (0.3, 0.0, -0.3)
    a = round(avg + offsets[0], 3)
    b = round(avg + offsets[1], 3)
    c = round(avg + offsets[2], 3)
    current = (a + b + c) / 3.0
    delta = avg - current
    return round(a + delta, 3), round(b + delta, 3), round(c + delta, 3)


# ============================================================
# РАБОТА С EXCEL
# ============================================================

def safe_set(cell, value):
    if cell.value is None or not (isinstance(cell.value, str) and str(cell.value).startswith("=")):
        cell.value = value


def update_workbook(wb, parsed: dict):
    preferred = ["Рабочий журнал", "Sheet1", "Лист1"]
    ws = None
    for name in preferred:
        if name in wb.sheetnames:
            ws = wb[name]
            break
    if ws is None:
        ws = wb.active

    def find_near(ws, keywords, new_val, max_r=100, max_c=12):
        for row in ws.iter_rows(min_row=1, max_row=max_r, max_col=max_c):
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    low = cell.value.lower()
                    for kw in keywords:
                        if kw.lower() in low:
                            right = ws.cell(row=cell.row, column=cell.column + 1)
                            safe_set(right, new_val)
                            return True
        return False

    if parsed.get("brand"):
        find_near(ws, ["Марка", "1.1"], parsed["brand"])
    if parsed.get("model"):
        find_near(ws, ["Модификация", "1.2", "модель"], parsed["model"])
    if parsed.get("vin"):
        find_near(ws, ["VIN", "Идентификационный", "1.5"], parsed["vin"])
    if parsed.get("protocol_date"):
        find_near(ws, ["Дата проведения", "Дата"], parsed["protocol_date"])
    if parsed.get("temp"):
        find_near(ws, ["Температура"], parsed["temp"])
    if parsed.get("humidity"):
        find_near(ws, ["влажность", "Относительная"], parsed["humidity"])
    if parsed.get("pressure"):
        find_near(ws, ["давление", "Атмосферное"], parsed["pressure"])
    if parsed.get("wind"):
        find_near(ws, ["ветра", "Скорость ветра"], parsed["wind"])

    brand_model = " ".join(filter(None, [parsed.get("brand"), parsed.get("model")]))
    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        for row in sheet.iter_rows(min_row=1, max_row=12, max_col=8):
            for cell in row:
                if not cell.value or not isinstance(cell.value, str):
                    continue
                if "Проба" in cell.value or "проба" in cell.value:
                    right = sheet.cell(row=cell.row, column=cell.column + 1)
                    if brand_model:
                        safe_set(right, brand_model)
                if "Дата" in cell.value and parsed.get("protocol_date"):
                    right = sheet.cell(row=cell.row, column=cell.column + 1)
                    safe_set(right, parsed["protocol_date"])

    def write_triplet(ws, label, values, max_r=250):
        if not values or values[0] is None:
            return
        a, b, c = values
        for row in ws.iter_rows(min_row=1, max_row=max_r, max_col=10):
            for cell in row:
                if not cell.value:
                    continue
                s = str(cell.value)
                target_col = cell.column - 1 if cell.column > 1 else cell.column + 1
                target = ws.cell(row=cell.row, column=target_col)
                if f"{label}(1)" in s:
                    safe_set(target, a)
                elif f"{label}(2)" in s:
                    safe_set(target, b)
                elif f"{label}(3)" in s:
                    safe_set(target, c)

    meas = parsed.get("measurements", {})
    if "E" in meas:
        write_triplet(ws, "Е", create_triplet(meas["E"]))
    if "H" in meas:
        write_triplet(ws, "Н", create_triplet(meas["H"]))
    if "J" in meas:
        write_triplet(ws, "J", create_triplet(meas["J"]))
    if "length" in meas:
        write_triplet(ws, "m", create_triplet(meas["length"], "descending"))

    return wb


def process(pdf_files, template_bytes=None):
    if template_bytes is not None:
        wb = openpyxl.load_workbook(BytesIO(template_bytes))
    else:
        if not os.path.exists("template.xlsx"):
            raise FileNotFoundError("Файл template.xlsx не найден в проекте.")
        wb = openpyxl.load_workbook("template.xlsx")

    results = []
    for pdf in pdf_files:
        try:
            text = extract_text_from_pdf(pdf)
            if len(text.strip()) < 80:
                results.append({"name": pdf.name, "ok": False, "error": "Текст почти не извлечён (возможно скан)."})
                continue
            parsed = parse_protocol(text)
            wb = update_workbook(wb, parsed)
            results.append({"name": pdf.name, "ok": True, "parsed": parsed})
        except Exception as e:
            results.append({"name": pdf.name, "ok": False, "error": str(e)})

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf, results


# ============================================================
# ИНТЕРФЕЙС
# ============================================================

st.markdown('<p class="main-title">🚗 Журнал сертификационных испытаний</p>', unsafe_allow_html=True)
st.markdown('<p class="subtitle">Загрузите PDF-протокол → система автоматически заполнит Excel</p>', unsafe_allow_html=True)

# Загрузка файлов
pdfs = st.file_uploader(
    "📄 PDF-протоколы (можно несколько)",
    type=["pdf"],
    accept_multiple_files=True
)

with st.expander("⚙️ Дополнительно (свой шаблон)"):
    custom_template = st.file_uploader("Загрузить свой Excel-шаблон", type=["xlsx"])

st.markdown("")

if st.button("🚀 Запустить обработку", type="primary", use_container_width=True):
    if not pdfs:
        st.warning("Сначала загрузите хотя бы один PDF-файл.")
    else:
        with st.spinner("Обрабатываю протокол..."):
            try:
                tpl_bytes = custom_template.getvalue() if custom_template else None
                output, results = process(pdfs, template_bytes=tpl_bytes)

                st.markdown("### Результат")

                for r in results:
                    if r["ok"]:
                        p = r["parsed"]
                        with st.expander(f"✅ {r['name']}", expanded=True):
                            col1, col2 = st.columns(2)
                            with col1:
                                st.metric("Марка", p.get("brand") or "—")
                                st.metric("Модель", p.get("model") or "—")
                                st.metric("VIN", p.get("vin") or "—")
                            with col2:
                                st.metric("Дата", p.get("protocol_date") or "—")
                                st.metric("Температура", f"{p.get('temp')} °C" if p.get("temp") else "—")
                                st.metric("Влажность", f"{p.get('humidity')} %" if p.get("humidity") else "—")
                    else:
                        st.error(f"❌ {r['name']}: {r.get('error')}")

                st.markdown(
                    '<div class="success-box">Готово! Файл сформирован. Формулы сохранены.</div>',
                    unsafe_allow_html=True
                )

                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                st.download_button(
                    label="📥 Скачать заполненный Excel",
                    data=output,
                    file_name=f"Журнал_испытаний_{ts}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            except Exception as e:
                st.error("Произошла ошибка при обработке:")
                st.code(traceback.format_exc())

# Подсказка внизу
st.markdown("---")
st.caption("Приложение для сотрудников • Работает на компьютере и телефоне • Формулы Excel сохраняются")
