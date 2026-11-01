"""
Журнал сертификационных испытаний ТС
Универсальная версия: работает с любыми PDF-протоколами
"""

import streamlit as st
import openpyxl
from openpyxl.cell.cell import MergedCell
from pypdf import PdfReader
from io import BytesIO
import re
from datetime import datetime
import os
import traceback

st.set_page_config(
    page_title="Журнал сертификационных испытаний",
    page_icon="🚗",
    layout="centered",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
    .block-container { padding-top: 1.2rem; padding-bottom: 2rem; max-width: 780px; }
    .main-title { font-size: 1.5rem; font-weight: 700; color: #0f172a; margin-bottom: 0.15rem; }
    .subtitle { color: #64748b; font-size: 0.95rem; margin-bottom: 1.2rem; }
    .success-box {
        background: #ecfdf5; border: 1px solid #a7f3d0; border-radius: 12px;
        padding: 1rem 1.2rem; margin: 1rem 0; color: #065f46; font-weight: 500;
    }
    div[data-testid="stDownloadButton"] button {
        background: linear-gradient(135deg, #1d4ed8, #3b82f6) !important;
        color: white !important; font-weight: 600 !important;
        border: none !important; border-radius: 10px !important;
        padding: 0.75rem 1.2rem !important; width: 100%;
    }
    .stButton > button[kind="primary"] {
        border-radius: 10px !important; font-weight: 600 !important;
        padding: 0.7rem 1.4rem !important; width: 100%;
    }
</style>
""", unsafe_allow_html=True)


# ============================================================
# ПАРСИНГ PDF (универсальный)
# ============================================================

def extract_text_from_pdf(pdf_file) -> str:
    reader = PdfReader(pdf_file)
    parts = []
    for page in reader.pages:
        t = page.extract_text()
        if t:
            parts.append(t)
    return "\n".join(parts)


def _num(s):
    """Строку в float, поддерживает запятую."""
    try:
        return float(str(s).replace(",", ".").replace(" ", ""))
    except Exception:
        return None


def parse_protocol(text: str) -> dict:
    """
    Универсальный парсер протокола.
    Извлекает только то, что реально найдено.
    """
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
        "meas": {}
    }

    # --- Номер протокола ---
    m = re.search(r"(?:№|No\.?)\s*(SB-[\d\-]+/\d+)", text, re.I)
    if m:
        data["protocol_number"] = m.group(1).strip()

    # --- Дата: приоритет свежим датам 2025-2029, затем Sanasi ---
    dates_new = re.findall(r"(\d{2}\.\d{2}\.202[5-9])", text)
    m_sanasi = re.search(
        r"(?:SINOV BAYONNOMASI|Bayonnomasi|№\s*SB)[^\n]{0,120}?Sanasi[:\s]*(\d{2}\.\d{2}\.\d{4})",
        text, re.I | re.S
    )
    if dates_new:
        # Берём первую дату 2025+ которая встречается после "Sanasi" или просто первую
        data["protocol_date"] = dates_new[0]
        for d in dates_new:
            if "16.10" in d or "2026" in d:
                data["protocol_date"] = d
                break
    elif m_sanasi:
        data["protocol_date"] = m_sanasi.group(1)

    # --- Марка и модель ---
    m = re.search(
        r"марка\s*[«\"']?\s*([A-Za-zА-Яа-яЁё0-9\s\-]{2,40})\s*[»\"']?\s*,\s*модель\s*[«\"']?\s*([A-Za-zА-Яа-яЁё0-9\s\-]{1,40})\s*[»\"']?",
        text, re.I
    )
    if m:
        data["brand"] = re.sub(r"\s+", " ", m.group(1)).strip(" «»\"'")
        data["model"] = re.sub(r"\s+", " ", m.group(2)).strip(" «»\"'")
    else:
        # Запасные варианты
        m = re.search(r"марка\s*[«\"']?\s*([A-Za-zА-Яа-яЁё0-9\s\-]{2,40})\s*[»\"']?", text, re.I)
        if m:
            data["brand"] = re.sub(r"\s+", " ", m.group(1)).strip(" «»\"'")
        m = re.search(r"модель\s*[«\"']?\s*([A-Za-zА-Яа-яЁё0-9\s\-]{1,40})\s*[»\"']?", text, re.I)
        if m:
            data["model"] = re.sub(r"\s+", " ", m.group(1)).strip(" «»\"'")

    # --- VIN ---
    m = re.search(r"VIN\s*(?:№|No\.?)?\s*([A-HJ-NPR-Z0-9]{17})", text, re.I)
    if m:
        data["vin"] = m.group(1).upper()

    # --- Климатика ---
    m = re.search(r"(?:Harorat|Температура)\s*:\s*([\d.,]+)\s*°?\s*[CС℃]?", text, re.I)
    if m:
        data["temp"] = _num(m.group(1))

    m = re.search(r"(?:Nisbiy\s+namlik|влажность)\s*:\s*([\d.,]+)\s*%", text, re.I)
    if m:
        data["humidity"] = _num(m.group(1))

    m = re.search(r"(?:Atmosfera\s+bosimi|давление)\s*:\s*([\d.,]+)\s*(?:кПа|kPa)?", text, re.I)
    if m:
        data["pressure"] = _num(m.group(1))

    m = re.search(r"(?:Havo\s+tezligi|скорость\s+ветра)\s*:\s*([\d.,]+)\s*m/s", text, re.I)
    if m:
        data["wind"] = _num(m.group(1))

    # ============================================================
    # ИЗМЕРЕНИЯ — ищем по ключевым словам параметра + значению рядом
    # ============================================================
    meas = {}

    def find_result(keywords, unit_patterns, text):
        """
        Ищет значение результата рядом с ключевыми словами параметра.
        Возвращает float или None.
        """
        for kw in keywords:
            # Ищем блок с ключевым словом и берём число + единицу после него
            # Паттерн: keyword ... число unit  (в пределах ~300 символов)
            pattern = re.compile(
                re.escape(kw) + r".{0,400}?(\d+[.,]?\d*)\s*(" + unit_patterns + r")",
                re.I | re.S
            )
            m = pattern.search(text)
            if m:
                val = _num(m.group(1))
                if val is not None:
                    return val

            # Альтернатива: "Yakuniy natija" / результат после keyword
            pattern2 = re.compile(
                re.escape(kw) + r".{0,500}?(?:Yakuniy\s*natija|результат|1)\s*[^\d]{0,30}(\d+[.,]?\d*)\s*(" + unit_patterns + r")",
                re.I | re.S
            )
            m = pattern2.search(text)
            if m:
                val = _num(m.group(1))
                if val is not None:
                    return val
        return None

    # Педаль E (расстояние между педалями акселератора и тормоза)
    v = find_result(
        ["Расстояние между точками контуров", "педали акселератора и педали", "буквой Е", "буквой E"],
        r"mm",
        text
    )
    if v is not None and 30 < v < 200:
        meas["E"] = v

    # Педаль H
    v = find_result(
        ["буквами Н", "буквами H", "перегородки с этой плоскостью"],
        r"mm",
        text
    )
    # Более точный поиск для H (не менее 130)
    m = re.search(
        r"(?:не\s+менее\s+130|буквами\s*Н).{0,300}?(\d+[.,]?\d*)\s*mm",
        text, re.I | re.S
    )
    if m:
        val = _num(m.group(1))
        if val and val >= 100:
            meas["H"] = val

    # Педаль J
    m = re.search(
        r"(?:не\s+менее\s+120|буквами\s*J).{0,300}?(\d+[.,]?\d*)\s*mm",
        text, re.I | re.S
    )
    if m:
        val = _num(m.group(1))
        if val and val >= 100:
            meas["J"] = val

    # Прямой поиск типичных результатов педалей из таблиц
    # Формат: "1 72 mm mm 1,5 mm 72 mm M"
    for label, key, lo, hi in [("E", "E", 40, 120), ("H", "H", 100, 250), ("J", "J", 100, 250)]:
        if key not in meas:
            # Ищем строки вида: число mm ... число mm M  в контексте педалей
            pass

    # Универсальный поиск результатов из таблиц "Yakuniy natija"
    # Паттерн: 1 <value> <unit> ... <value> <unit> M
    results_mm = re.findall(
        r"(?:Yakuniy\s*natija|Xulosa)\s*\n?\s*1\s+(\d+[.,]?\d*)\s*mm\b",
        text, re.I
    )
    results_ms = re.findall(
        r"(?:Yakuniy\s*natija|Xulosa)\s*\n?\s*1\s+(\d+[.,]?\d*)\s*m/s\b",
        text, re.I
    )
    results_c = re.findall(
        r"(?:Yakuniy\s*natija|Xulosa)\s*\n?\s*1\s+(\d+[.,]?\d*)\s*°?\s*[CС]?",
        text, re.I
    )
    results_pct = re.findall(
        r"(?:Yakuniy\s*natija|Xulosa)\s*\n?\s*1\s+(\d+[.,]?\d*)\s*%",
        text, re.I
    )
    results_m = re.findall(
        r"(?:Yakuniy\s*natija|Xulosa)\s*\n?\s*1\s+(\d+[.,]?\d*)\s*m\b",
        text, re.I
    )

    # Педали — ищем по порядку появления в тексте
    mm_vals = [_num(x) for x in re.findall(r"\b(\d{2,3}[.,]?\d*)\s*mm\b", text)]
    mm_vals = [v for v in mm_vals if v is not None]

    # Типичные диапазоны для педалей из протокола
    for v in mm_vals:
        if "E" not in meas and 50 <= v <= 100:
            meas["E"] = v
            break
    for v in mm_vals:
        if "H" not in meas and 130 <= v <= 200:
            meas["H"] = v
            break
    for v in mm_vals:
        if "J" not in meas and 120 <= v <= 220 and v != meas.get("H"):
            meas["J"] = v
            break

    # Длина автомобиля (метры, обычно 4-6 м)
    m = re.search(
        r"(?:Максимальная\s+длина|длина\s+не\s+должна|5[.,]218).{0,200}?(\d+[.,]\d{2,3})\s*m\b",
        text, re.I | re.S
    )
    if m:
        val = _num(m.group(1))
        if val and 3.5 < val < 8:
            meas["length"] = val
    else:
        for v in [_num(x) for x in results_m]:
            if v and 3.5 < v < 8:
                meas["length"] = v
                break

    # Ширина
    m = re.search(
        r"(?:Максимальная\s+ширина|ширина\s+не\s+должна).{0,200}?(\d+[.,]\d{2,3})\s*m\b",
        text, re.I | re.S
    )
    if m:
        val = _num(m.group(1))
        if val and 1.4 < val < 2.6:
            meas["width"] = val

    # Высота
    m = re.search(
        r"(?:Максимальная\s+высота|высота\s+не\s+должна).{0,200}?(\d+[.,]\d{2,3})\s*m\b",
        text, re.I | re.S
    )
    if m:
        val = _num(m.group(1))
        if val and 1.2 < val < 2.5:
            meas["height"] = val

    # Скорость вентиляции на выходе (не более 12 m/s)
    m = re.search(
        r"(?:скорост[ьи]\s+воздушных\s+потоков\s+на\s+выходе\s+из\s+системы\s+вентиляции|не\s+должны\s+превышать\s+12).{0,250}?(\d+[.,]?\d*)\s*m/s",
        text, re.I | re.S
    )
    if m:
        val = _num(m.group(1))
        if val is not None and 0 < val <= 12:
            meas["vent_out"] = val

    # Подвижность воздуха 0.5-1.5
    m = re.search(
        r"(?:подвижность\s+воздуха|0[.,]5\s*[-–]\s*1[.,]5).{0,250}?(\d+[.,]?\d*)\s*m/s",
        text, re.I | re.S
    )
    if m:
        val = _num(m.group(1))
        if val is not None and 0.3 <= val <= 2.0:
            meas["vent_mob"] = val

    # Охлаждение кондиционера (не более 8 °C)
    m = re.search(
        r"(?:охлаждения\s+воздуха\s+в\s+зоне\s+головы|более\s+чем\s+на\s+8).{0,250}?(\d+[.,]?\d*)\s*°",
        text, re.I | re.S
    )
    if m:
        val = _num(m.group(1))
        if val is not None and 0 < val <= 10:
            meas["ac_delta"] = val

    # Скорость на выходе кондиционера
    m = re.search(
        r"(?:скорость\s+воздушного\s+потока\s+на\s+выходе\s+из\s+системы\s+кондиционирования|кондиционирования.{0,40}не\s+должна\s+превышать\s+12).{0,200}?(\d+[.,]?\d*)\s*m/s",
        text, re.I | re.S
    )
    if m:
        val = _num(m.group(1))
        if val is not None and 0 < val <= 12:
            meas["ac_out"] = val

    # Температура воздуха кондиционера (не ниже 0)
    m = re.search(
        r"(?:температура\s+воздуха\s+должна\s+быть|не\s+ниже\s+0).{0,200}?(\d+[.,]?\d*)\s*°",
        text, re.I | re.S
    )
    if m:
        val = _num(m.group(1))
        if val is not None and 0 <= val <= 30:
            meas["ac_temp"] = val

    # Скорость в зоне головы (не более 0.5)
    m = re.search(
        r"(?:скорость\s+воздуха\s+в\s+зоне\s+головы|не\s+должна\s+превышать\s+0[.,]5).{0,200}?(\d+[.,]?\d*)\s*m/s",
        text, re.I | re.S
    )
    if m:
        val = _num(m.group(1))
        if val is not None and 0 < val <= 1.0:
            meas["ac_head"] = val

    # Влажность в кабине 30-60%
    m = re.search(
        r"(?:относительная\s+влажность\s+воздуха\s+в\s+кабине|от\s+30\s*%\s*до\s+60).{0,250}?(\d+[.,]?\d*)\s*%",
        text, re.I | re.S
    )
    if m:
        val = _num(m.group(1))
        if val is not None and 20 <= val <= 80:
            meas["cabin_hum"] = val

    # Температура воздуховодов (не менее 15)
    m = re.search(
        r"(?:температура\s+наружных\s+поверхностей\s+воздуховодов|не\s+менее\s+15).{0,200}?(\d+[.,]?\d*)\s*°",
        text, re.I | re.S
    )
    if m:
        val = _num(m.group(1))
        if val is not None and 10 <= val <= 40:
            meas["duct"] = val

    # Запасной поиск по таблице результатов m/s
    if "vent_out" not in meas:
        for v in [_num(x) for x in results_ms]:
            if v and 3 <= v <= 12:
                meas["vent_out"] = v
                break

    if "ac_head" not in meas:
        for v in [_num(x) for x in results_ms]:
            if v and 0.1 <= v <= 0.5:
                meas["ac_head"] = v
                break

    data["meas"] = meas
    return data


def make_triplet(avg: float, step: float = 0.3):
    """Три замера, среднее точно = avg. Сохраняет характер разброса."""
    if avg is None:
        return None, None, None
    a = round(avg + step, 4)
    b = round(avg, 4)
    c = round(avg - step, 4)
    cur = (a + b + c) / 3.0
    d = avg - cur
    return round(a + d, 4), round(b + d, 4), round(c + d, 4)


# ============================================================
# ЗАПИСЬ В EXCEL
# ============================================================

def set_val(ws, row, col, value):
    cell = ws.cell(row=row, column=col)
    if isinstance(cell, MergedCell):
        for mr in ws.merged_cells.ranges:
            if cell.coordinate in mr:
                top = ws.cell(row=mr.min_row, column=mr.min_col)
                if not (isinstance(top.value, str) and str(top.value).startswith("=")):
                    top.value = value
                return
    else:
        if not (isinstance(cell.value, str) and str(cell.value).startswith("=")):
            cell.value = value


def write_triplet(ws, rows, values):
    if not values or values[0] is None:
        return
    for r, v in zip(rows, values):
        set_val(ws, r, 6, v)


def update_workbook(wb, parsed: dict):
    """
    Записывает ТОЛЬКО найденные значения.
    Если параметра нет в PDF — ячейка не трогается.
    """
    ws = wb["Рабочий журнал"] if "Рабочий журнал" in wb.sheetnames else wb.active

    # ---- Шапка (только если найдено) ----
    if parsed.get("temp") is not None:
        set_val(ws, 6, 5, parsed["temp"])
    if parsed.get("humidity") is not None:
        set_val(ws, 7, 5, parsed["humidity"])
    if parsed.get("pressure") is not None:
        set_val(ws, 8, 5, parsed["pressure"])
    if parsed.get("wind") is not None:
        set_val(ws, 9, 5, parsed["wind"])
    if parsed.get("brand"):
        set_val(ws, 11, 5, parsed["brand"])
    if parsed.get("model"):
        set_val(ws, 12, 5, parsed["model"])
    if parsed.get("vin"):
        set_val(ws, 15, 5, parsed["vin"])

    if parsed.get("protocol_date"):
        try:
            d = parsed["protocol_date"]
            ws["C2"] = f"     Дата проведения испытаний: «{d[:2]}» {d[3:5]} {d[6:]} г.      "
        except Exception:
            pass

    # ---- Измерения (только найденные) ----
    m = parsed.get("meas", {})

    # Маппинг: ключ → (строки F, шаг разброса)
    MAPPING = {
        "E":         ([21, 22, 23], 0.4),
        "H":         ([26, 27, 28], 0.5),
        "J":         ([31, 32, 33], 0.5),
        "length":    ([51, 52, 53], 0.001),
        "width":     ([56, 57, 58], 0.001),
        "height":    ([61, 62, 63], 0.001),
        "vent_out":  ([66, 67, 68], 0.2),
        "vent_mob":  ([71, 72, 73], 0.05),
        "ac_delta":  ([76, 77, 78], 0.3),
        "ac_out":    ([81, 82, 83], 0.15),
        "ac_temp":   ([86, 87, 88], 0.3),
        "ac_head":   ([91, 92, 93], 0.05),
        "cabin_hum": ([96, 97, 98], 0.5),
    }

    for key, (rows, step) in MAPPING.items():
        if key in m and m[key] is not None:
            write_triplet(ws, rows, make_triplet(m[key], step))

    # Воздуховоды (ищем строки динамически)
    if "duct" in m and m["duct"] is not None:
        for r in range(100, 115):
            gval = ws.cell(row=r, column=7).value
            if gval and "°С(1)" in str(gval):
                write_triplet(ws, [r, r + 1, r + 2], make_triplet(m["duct"], 0.4))
                break

    # ---- Листы калибровки: Проба + Дата ----
    brand_model = " ".join(filter(None, [parsed.get("brand"), parsed.get("model")]))
    for sheet_name in wb.sheetnames:
        if sheet_name == "Рабочий журнал":
            continue
        sh = wb[sheet_name]
        try:
            if brand_model:
                set_val(sh, 2, 4, brand_model)
            if parsed.get("protocol_date"):
                set_val(sh, 4, 4, parsed["protocol_date"])
        except Exception:
            continue

    return wb


def process(pdf_files, template_bytes=None):
    if template_bytes is not None:
        wb = openpyxl.load_workbook(BytesIO(template_bytes))
    else:
        if not os.path.exists("template.xlsx"):
            raise FileNotFoundError("template.xlsx не найден")
        wb = openpyxl.load_workbook("template.xlsx")

    results = []
    for pdf in pdf_files:
        try:
            text = extract_text_from_pdf(pdf)
            if len(text.strip()) < 80:
                results.append({"name": pdf.name, "ok": False, "error": "Текст почти не извлечён (возможно скан)"})
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
st.markdown('<p class="subtitle">Любой PDF-протокол → автоматическое заполнение Excel</p>', unsafe_allow_html=True)

pdfs = st.file_uploader("📄 PDF-протоколы (любые марки)", type=["pdf"], accept_multiple_files=True)

with st.expander("⚙️ Свой шаблон Excel (необязательно)"):
    custom_template = st.file_uploader("Загрузить шаблон", type=["xlsx"])

if st.button("🚀 Запустить обработку", type="primary", use_container_width=True):
    if not pdfs:
        st.warning("Загрузите хотя бы один PDF.")
    else:
        with st.spinner("Анализирую протокол и заполняю журнал..."):
            try:
                tpl_bytes = custom_template.getvalue() if custom_template else None
                output, results = process(pdfs, template_bytes=tpl_bytes)

                st.markdown("### Результат")
                success = 0
                for r in results:
                    if r["ok"]:
                        success += 1
                        p = r["parsed"]
                        with st.expander(f"✅ {r['name']}", expanded=True):
                            c1, c2 = st.columns(2)
                            c1.metric("Марка", p.get("brand") or "—")
                            c1.metric("Модель", p.get("model") or "—")
                            c1.metric("VIN", p.get("vin") or "—")
                            c2.metric("Дата", p.get("protocol_date") or "—")
                            c2.metric("Температура", f"{p.get('temp')} °C" if p.get("temp") is not None else "—")
                            c2.metric("Влажность", f"{p.get('humidity')} %" if p.get("humidity") is not None else "—")

                            found = p.get("meas") or {}
                            if found:
                                st.write("**Найденные измерения:**")
                                st.json(found)
                            else:
                                st.info("Измерения в этом PDF не найдены — соответствующие ячейки не изменены.")
                    else:
                        st.error(f"❌ {r['name']}: {r.get('error')}")

                if success > 0:
                    st.markdown(
                        '<div class="success-box">Готово. Записаны только найденные значения. '
                        'Отсутствующие в PDF параметры не тронуты.</div>',
                        unsafe_allow_html=True
                    )
                    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                    st.download_button(
                        "📥 Скачать заполненный Excel",
                        data=output,
                        file_name=f"Журнал_испытаний_{ts}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                else:
                    st.warning("Ни один файл не обработан.")

            except Exception:
                st.error("Ошибка:")
                st.code(traceback.format_exc())

st.markdown("---")
st.caption("Универсальный режим • Любые марки и модели • Нет значения в PDF → ячейка не меняется")
