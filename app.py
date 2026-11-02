# -*- coding: utf-8 -*-
"""
Заполнение рабочего журнала сертификационных испытаний ТС категории M1
по PDF-протоколу испытаний (шаблон: template.xlsx).

Запуск:  streamlit run app.py
Зависимости: streamlit, openpyxl, pypdf
"""
import io
import re
import zipfile
import warnings
from datetime import date
from decimal import Decimal as D
from fractions import Fraction as Fr
from pathlib import Path
from xml.sax.saxutils import escape

import openpyxl
import streamlit as st
from pypdf import PdfReader

warnings.filterwarnings("ignore")

JOURNAL_SHEET = "Рабочий журнал"
TEMPLATE_NAME = "template.xlsx"

# --------------------------------------------------------------------------
# Раскладка журнала: ключ -> (подпись, первая строка тройки, колонка)
# --------------------------------------------------------------------------
SIMPLE = {
    "ped_e": ("Педаль E", 21, "F"),
    "ped_h": ("Педаль H", 26, "F"),
    "ped_j": ("Педаль J", 31, "F"),
    "co": ("Оксид углерода CO", 36, "F"),
    "ch": ("Углеводороды CH", 41, "F"),
    "horn": ("Уровень звука сигнала", 46, "F"),
    "length": ("Длина", 51, "F"),
    "width": ("Ширина", 56, "F"),
    "height": ("Высота", 61, "F"),
    "vent_speed": ("Скорость воздуха на выходе вентиляции", 66, "F"),
    "vent_mob": ("Подвижность воздуха в зоне головы", 71, "F"),
    "ac_cool": ("Охлаждение воздуха кондиционером", 76, "F"),
    "ac_speed": ("Скорость воздуха на выходе кондиционера", 81, "F"),
    "ac_temp": ("Температура воздуха", 86, "F"),
    "ac_head": ("Скорость воздуха в зоне головы", 91, "F"),
    "humidity": ("Относительная влажность", 96, "F"),
    "duct_temp": ("Температура поверхностей воздуховодов", 101, "F"),
    "heat": ("Температура системы отопления", 106, "F"),
    "protr": ("Выступающие части (радиус)", 112, "F"),
    "grille_gap": ("Решётки и щели (зазоры)", 118, "F"),
    "grille_rad": ("Решётки и щели (радиус кривизны)", 123, "F"),
    "wiper": ("Стеклоочиститель", 128, "F"),
    "bumper": ("Бамперы", 133, "F"),
    "door": ("Двери", 138, "F"),
    "handle": ("Ручки двери", 143, "F"),
    "trim": ("Обшивка кузова", 149, "F"),
    "glass_w": ("Светопропускание ветрового стекла", 158, "F"),
    "glass_s": ("Светопропускание переднего бокового стекла", 163, "F"),
    "plate_a1": ("Номерной знак: угол (б)", 169, "F"),
    "plate_a2": ("Номерной знак: угол (в)", 174, "F"),
    "plate_h1": ("Номерной знак: высота min", 179, "F"),
    "plate_h2": ("Номерной знак: высота max", 184, "F"),
    "noise": ("Внешний шум", 206, "F"),
}
LAMPS = {
    "dipped": ("Фары ближнего света", 214),
    "fog_front": ("Передняя противотуманная фара", 219),
    "reverse": ("Огонь заднего хода", 224),
    "pos_front": ("Габаритные огни передние", 229),
    "pos_rear": ("Габаритные огни задние", 234),
    "brake": ("Сигнал торможения", 239),
    "turn_front": ("Указатели поворота передние", 244),
    "turn_rear": ("Указатели поворота задние", 249),
    "side": ("Боковые повторители", 254),
    "rear_fog": ("Задний противотуманный фонарь", 261),
    "refl_rear": ("Световозвращатели задние", 266),
    "refl_front": ("Световозвращатели передние", 271),
}
LAMP_COLS = {"F": "верх", "G": "низ", "H": "между", "I": "по ширине"}

MONTHS = ["января", "февраля", "марта", "апреля", "мая", "июня", "июля",
          "августа", "сентября", "октября", "ноября", "декабря"]


def all_blocks():
    """Все блоки замеров журнала: ключ -> (подпись, строка, колонка)."""
    res = {}
    for k, (label, row, col) in SIMPLE.items():
        res[k] = (label, row, col)
    for k, (label, row) in LAMPS.items():
        for c, pos in LAMP_COLS.items():
            res[f"{k}|{c}"] = (f"{label}: {pos}", row, c)
    return res


# --------------------------------------------------------------------------
# Разбор PDF
# --------------------------------------------------------------------------
def norm(s):
    s = s.replace("\u0451", "\u0435").replace("\u0401", "\u0415")
    return re.sub(r"\s+", " ", s).strip()


def pdf_text(pdf_bytes):
    reader = PdfReader(io.BytesIO(pdf_bytes))
    return "\n".join((p.extract_text() or "") for p in reader.pages)


def to_dec(s):
    return D(s.replace(",", ".").replace("\u2212", "-"))


def decimals_of(s):
    s = s.replace(",", ".")
    return len(s.split(".")[1]) if "." in s else 0


def parse_meta(raw):
    flat = norm(raw)
    meta = {}
    m = re.search(r"№\s*(SB-[\w\-/]+)", flat)
    meta["protocol"] = m.group(1) if m else ""
    m = re.search(r"марка\s*[«\"“]([^»\"”]+)[»\"”]", flat)
    meta["brand"] = m.group(1).strip() if m else ""
    m = re.search(r"модель\s*[«\"“]([^»\"”]+)[»\"”]", flat)
    meta["model"] = m.group(1).strip() if m else ""
    m = re.search(r"VIN\s*№?\s*([A-Za-z0-9]{17})", flat)
    meta["vin"] = m.group(1).upper() if m else ""
    m = (re.search(r"Sinov boshlanish vaqti:\s*(\d{2})\.(\d{2})\.(\d{4})", flat)
         or re.search(r"Sanasi:\s*(\d{2})\.(\d{2})\.(\d{4})", flat))
    meta["date"] = date(int(m.group(3)), int(m.group(2)), int(m.group(1))) if m else None

    def num(pat):
        mm = re.search(pat, flat)
        return (to_dec(mm.group(1)), (mm.group(2) or "")) if mm else (None, "")

    meta["temp"], _ = num(r"Harorat\s*:\s*(-?[\d.,]+)\s*(\S+)?")
    meta["hum"], _ = num(r"Nisbiy namlik\s*:\s*([\d.,]+)\s*(\S+)?")
    meta["wind"], _ = num(r"Havo tezligi\s*:\s*([\d.,]+)\s*(\S+)?")
    p, unit = num(r"Atmosfera bosimi\s*:\s*([\d.,]+)\s*(\S+)?")
    if p is not None:
        u = unit.lower()
        if u in ("гпа", "hpa", "mbar", "мбар"):
            p = p / 10
        elif u.startswith("мм") or u.startswith("mm"):
            p = (p * D("0.133322")).quantize(D("1"))
    meta["press"] = p
    return meta


def parse_params(raw):
    """Список параметров протокола: раздел, название, итоговые значения."""
    text = re.sub(r"Sinov bayonnomasi № SB-.*?etiladi\.", "", raw, flags=re.S)
    events = []
    for m in re.finditer(r"(\d+)\.\s*Sinalayotgan ko'rsatgich\s+(.*?)\s*Texnik talablarni", text, re.S):
        events.append((m.start(), m.end(), "sec", norm(m.group(2))))
    for m in re.finditer(r"(\d+)\.\s*Aniqlanayotgan parametr nomi", text):
        events.append((m.start(), m.end(), "par", None))
    events.sort(key=lambda e: e[0])

    params, section = [], ""
    for i, (st_, en, kind, val) in enumerate(events):
        if kind == "sec":
            section = val
            continue
        end = events[i + 1][0] if i + 1 < len(events) else len(text)
        block = text[en:end]
        name_part = re.split(r"Maxsulotga", block, maxsplit=1)[0]
        name = norm(name_part)
        name = re.sub(r"^\d+\s+(?=\D)", "", name)
        rows = []
        after = re.split(r"Xulosa", block, maxsplit=1)
        if len(after) == 2:
            for line in after[1].splitlines():
                mm = re.match(r"^\s*(\d+)\s+(.+?)\s+([MN])\s*$", line)
                if not mm:
                    continue
                toks = mm.group(2).split()
                if len(toks) >= 2 and re.fullmatch(r"-?[\d.,]+", toks[-2]):
                    rows.append((toks[-2], toks[-1], mm.group(3)))
        if not rows:
            continue
        vals = [to_dec(r[0]) for r in rows]
        dp = max(decimals_of(r[0]) for r in rows)
        value = (sum(vals) / len(vals)).quantize(D(1).scaleb(-dp)) if dp else \
            (sum(vals) / len(vals)).quantize(D(1))
        params.append({
            "section": norm(section).lower(),
            "name": name,
            "name_l": name.lower(),
            "value": value,
            "dp": dp,
            "unit": rows[0][1],
            "ok": all(r[2] == "M" for r in rows),
        })
    return params


# --------------------------------------------------------------------------
# Сопоставление параметров протокола с блоками журнала
# --------------------------------------------------------------------------
def detect_lamp(n):
    if "от передней части атс" in n:
        return "side"
    if "фары ближнего" in n or "ближнего света" in n:
        return "dipped"
    if "заднего хода" in n:
        return "reverse"
    if "противотуман" in n:
        return "rear_fog" if "задн" in n else "fog_front"
    if "габаритные огни" in n:
        return "pos_front" if "передн" in n else ("pos_rear" if "задн" in n else None)
    if "сигнал торможения" in n:
        return "brake"
    if "указатели поворота" in n:
        if "боковые" in n:
            return "side"
        if "передн" in n:
            return "turn_front"
        if "задн" in n:
            return "turn_rear"
        return None
    if "световозвращатели" in n:
        return "refl_front" if "передн" in n else ("refl_rear" if "задн" in n else None)
    return None


def lamp_col(n):
    if "от передней части атс" in n:
        return "H"
    if "верхней точки" in n:
        return "F"
    if "нижней точки" in n:
        return "G"
    if "между крайними внутренними" in n:
        return "H"
    if "габарита кузова по ширине" in n:
        return "I"
    return None


def classify(section, n, cur_lamp):
    """Возвращает (ключ блока или None, текущая лампа)."""
    if "внешним световым" in section or "световым приборам" in section:
        lamp = detect_lamp(n) or cur_lamp
        col = lamp_col(n)
        if lamp and col:
            return f"{lamp}|{col}", lamp
        return None, lamp
    cur_lamp = None
    if "педалей" in section or "обозначенное в приложении 4" in n or "перегородки" in n:
        m = re.search(r"букв\w*\s+([A-Za-zА-Яа-я])\s*$", n)
        if m:
            ch = m.group(1).upper()
            if ch in ("E", "Е"):
                return "ped_e", None
            if ch in ("H", "Н"):
                return "ped_h", None
            if ch == "J":
                return "ped_j", None
    if "оксида углерода" in n and "кабин" not in n:
        return "co", None
    if "углеводороды" in n:
        return "ch", None
    if "максимальная длина" in n:
        return "length", None
    if "максимальная ширина" in n:
        return "width", None
    if "максимальная высота" in n:
        return "height", None
    if "вентиляци" in section or "кондиционирован" in section:
        if "потоков на выходе из системы вентиляции" in n:
            return "vent_speed", None
        if "система вентиляции должна" in n:
            return "vent_mob", None
        if "исключать возможность охлаждения" in n:
            return "ac_cool", None
        if "потока на выходе из системы кондиционирования" in n:
            return "ac_speed", None
        if "температура воздуха должна" in n:
            return "ac_temp", None
        if "в зоне головы" in n:
            return "ac_head", None
        if "относительная влажность" in n:
            return "humidity", None
        if "температура наружных поверхностей воздуховодов" in n:
            return "duct_temp", None
    if "температура поверхности без покрытия" in n:
        return "heat", None
    if "уровень внешнего шума" in n:
        if "разгона" in n:
            return "noise", None
        return None, None
    if "травмобезопасность" in section or "выступ" in section:
        if "выступающая часть наружной поверхности" in n:
            return "protr", None
        if "решетки и щели" in n:
            return "grille_gap", None
        if "стеклоочиститель" in n:
            return "wiper", None
        if "бамперы" in n:
            return "bumper", None
        if n == "двери":
            return "door", None
        if "для ручек двери" in n:
            return "handle", None
        if "обшивка кузова" in n:
            return "trim", None
    return None, None


def map_params(params):
    values, unmatched, dup = {}, [], []
    lamp = None
    for p in params:
        key, lamp = classify(p["section"], p["name_l"].replace("\u0451", "\u0435"), lamp)
        if key is None:
            unmatched.append(p)
            continue
        if key in values:
            dup.append(key)
            continue
        values[key] = p
    return values, unmatched, dup


# --------------------------------------------------------------------------
# Генерация троек замеров по образцу разброса из шаблона
# --------------------------------------------------------------------------
def gen_triplet(seed, P, dp):
    """Тройка с тем же характером разброса, что и в шаблоне."""
    if seed and all(x is not None for x in seed):
        t = list(seed)
        k = max([max(-x.as_tuple().exponent, 0) for x in t] + [dp])
        at = Fr(sum(Fr(x) for x in t), 3)
        delta = Fr(P) - at
        if (delta * 10 ** k).denominator == 1:
            dl = D(delta.numerator) / D(delta.denominator)
            exact = True
        else:
            rnd = (D(at.numerator) / D(at.denominator)).quantize(D(1).scaleb(-dp))
            dl = P - rnd
            exact = False
        return [x + dl for x in t], exact
    step = D(1).scaleb(-dp)
    return [P + step, P, P - step], True


def fmt(x):
    s = format(x.normalize(), "f")
    return "0" if s in ("-0", "") else s


# --------------------------------------------------------------------------
# Работа с xlsx как с zip/XML (картинки и стили остаются нетронутыми)
# --------------------------------------------------------------------------
def cell_re(ref):
    return re.compile(r'<c r="%s"((?: [A-Za-z:]+="[^"]*")*)\s*(?:/>|>.*?</c>)' % ref, re.S)


def set_num(xml, ref, val):
    m = cell_re(ref).search(xml)
    if not m:
        return xml, False
    attrs = re.sub(r' t="[^"]*"', "", m.group(1))
    return xml[:m.start()] + f'<c r="{ref}"{attrs}><v>{val}</v></c>' + xml[m.end():], True


def set_str(xml, ref, text):
    m = cell_re(ref).search(xml)
    if not m:
        return xml, False
    attrs = re.sub(r' t="[^"]*"', "", m.group(1))
    new = f'<c r="{ref}"{attrs} t="inlineStr"><is><t xml:space="preserve">{escape(text)}</t></is></c>'
    return xml[:m.start()] + new + xml[m.end():], True


def clear_cell(xml, ref):
    m = cell_re(ref).search(xml)
    if not m:
        return xml, False
    if "<f" in m.group(0):  # формулы не трогаем
        return xml, False
    attrs = re.sub(r' t="[^"]*"', "", m.group(1))
    return xml[:m.start()] + f'<c r="{ref}"{attrs}/>' + xml[m.end():], True


FORMULA_RE = re.compile(
    r'<c r="([A-Z]+\d+)"((?: [A-Za-z:]+="[^"]*")*)>(<f[^>]*?(?:/>|>.*?</f>))<v>[^<]*</v></c>', re.S)


T_ATTR_RE = re.compile(r' t="[^"]*"')


def _strip_repl(m):
    attrs = T_ATTR_RE.sub("", m.group(2))
    return '<c r="' + m.group(1) + '"' + attrs + ">" + m.group(3) + "</c>"


def strip_formula_cache(xml):
    return FORMULA_RE.sub(_strip_repl, xml)


def sheet_paths(z):
    wb = z.read("xl/workbook.xml").decode("utf8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf8")
    rid2t = {}
    for m in re.finditer(r"<Relationship\b[^>]*>", rels):
        tag = m.group(0)
        i = re.search(r'Id="([^"]+)"', tag)
        t = re.search(r'Target="([^"]+)"', tag)
        if i and t:
            tg = t.group(1).lstrip("/")
            rid2t[i.group(1)] = tg if tg.startswith("xl/") else "xl/" + tg
    res = []
    for m in re.finditer(r"<sheet\b[^>]*>", wb):
        tag = m.group(0)
        nm = re.search(r'name="([^"]*)"', tag).group(1)
        rid = re.search(r'r:id="([^"]+)"', tag).group(1)
        nm = (nm.replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")
                .replace("&quot;", '"').replace("&apos;", "'"))
        res.append((nm, rid2t[rid]))
    return res


def read_seeds(template_bytes):
    """Значения троек из шаблона (образец разброса)."""
    wb = openpyxl.load_workbook(io.BytesIO(template_bytes), data_only=True, read_only=True)
    ws = wb[JOURNAL_SHEET] if JOURNAL_SHEET in wb.sheetnames else wb.worksheets[0]
    cells = {}
    for row in ws.iter_rows(min_row=1, max_row=300, min_col=1, max_col=13):
        for c in row:
            if c.value is not None:
                cells[c.coordinate] = c.value
    wb.close()
    seeds = {}
    for key, (label, row, col) in all_blocks().items():
        vals = []
        for k in range(3):
            v = cells.get(f"{col}{row + k}")
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                vals.append(D(repr(v)))
            else:
                vals.append(None)
        seeds[key] = vals
    return seeds


def fill_workbook(template_bytes, meta, values):
    seeds = read_seeds(template_bytes)
    blocks = all_blocks()
    log = {"filled": [], "cleared": [], "inexact": [], "missing_cells": []}

    zin = zipfile.ZipFile(io.BytesIO(template_bytes))
    paths = sheet_paths(zin)
    jpath = next((p for n, p in paths if n == JOURNAL_SHEET), paths[0][1])
    other = [p for n, p in paths if p != jpath]

    out = {}
    j = zin.read(jpath).decode("utf8")

    # --- шапка журнала
    d = meta["date"]
    if d:
        txt = f"     Дата проведения испытаний: «{d.day:02d}» {MONTHS[d.month - 1]} {d.year} г.      "
        j, ok = set_str(j, "C2", txt)
        if not ok:
            log["missing_cells"].append("C2")
    for ref, v in (("E11", meta["brand"] + " "), ("E12", meta["model"]), ("E15", meta["vin"] + " ")):
        if v.strip():
            j, ok = set_str(j, ref, v)
            if not ok:
                log["missing_cells"].append(ref)
    for ref, v in (("D6", meta["temp"]), ("D7", meta["hum"]), ("D8", meta["press"]), ("D9", meta["wind"])):
        if v is not None:
            j, ok = set_num(j, ref, fmt(D(v)))
            if not ok:
                log["missing_cells"].append(ref)

    # --- замеры
    for key, (label, row, col) in blocks.items():
        cells = [f"{col}{row + k}" for k in range(3)]
        if key in values:
            p = values[key]
            new, exact = gen_triplet(seeds[key], p["value"], p["dp"])
            for ref, x in zip(cells, new):
                j, ok = set_num(j, ref, fmt(x))
                if not ok:
                    log["missing_cells"].append(ref)
            log["filled"].append((label, p["value"], [fmt(x) for x in new], exact, p["ok"]))
            if not exact:
                log["inexact"].append(label)
        else:
            had = any(v is not None for v in seeds[key])
            for ref in cells:
                j, _ = clear_cell(j, ref)
            if had:
                log["cleared"].append(label)
    j = strip_formula_cache(j)
    out[jpath] = j.encode("utf8")

    # --- расчётные листы: объект и дата
    cal_name = f'{meta["brand"]} {meta["model"]}'.strip()
    serial = (d - date(1899, 12, 30)).days if d else None
    for p in other:
        x = zin.read(p).decode("utf8")
        if cal_name:
            x, _ = set_str(x, "D2", cal_name)
        if serial is not None:
            x, _ = set_num(x, "D4", str(serial))
        out[p] = strip_formula_cache(x).encode("utf8")

    wbx = zin.read("xl/workbook.xml").decode("utf8")
    if "<calcPr" in wbx:
        wbx = re.sub(r"<calcPr([^>]*?)/>", lambda m: "<calcPr" + re.sub(r' fullCalcOnLoad="[^"]*"', "", m.group(1))
                     + ' fullCalcOnLoad="1"/>', wbx, count=1)
    else:
        wbx = wbx.replace("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>')
    out["xl/workbook.xml"] = wbx.encode("utf8")

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as zout:
        for info in zin.infolist():
            if info.filename.endswith("/"):
                continue
            data = out.get(info.filename)
            if data is None:
                data = zin.read(info.filename)
            zout.writestr(info.filename, data)
    return buf.getvalue(), log


def process(pdf_bytes, template_bytes):
    raw = pdf_text(pdf_bytes)
    if "Aniqlanayotgan parametr nomi" not in raw:
        raise ValueError("В PDF не найден текст протокола (возможно, это скан без текстового слоя).")
    meta = parse_meta(raw)
    params = parse_params(raw)
    values, unmatched, dup = map_params(params)
    data, log = fill_workbook(template_bytes, meta, values)
    log["unmatched"] = [p["name"] for p in unmatched]
    log["dup"] = dup
    log["params"] = len(params)
    log["nonconf"] = [p["name"] for p in params if not p["ok"]]
    return data, meta, log


def safe_name(s):
    return re.sub(r"[^\w\-.]+", "_", s, flags=re.U).strip("_")


# --------------------------------------------------------------------------
# Интерфейс Streamlit
# --------------------------------------------------------------------------
def main():
    st.set_page_config(page_title="Журнал M1", page_icon="📋", layout="centered")
    st.markdown(
        "<style>.stButton>button,.stDownloadButton>button{min-height:3rem;font-size:1.05rem}"
        ".block-container{padding-top:1.2rem;padding-bottom:2rem}</style>",
        unsafe_allow_html=True,
    )
    st.title("📋 Журнал испытаний M1")
    st.caption("Загрузите PDF-протокол — получите заполненный журнал по шаблону.")

    tpl_path = Path(__file__).resolve().parent / TEMPLATE_NAME
    template_bytes = tpl_path.read_bytes() if tpl_path.exists() else None
    if template_bytes is None:
        up_t = st.file_uploader("Шаблон журнала (template.xlsx)", type=["xlsx"], key="tpl")
        if up_t is not None:
            template_bytes = up_t.getvalue()
    else:
        with st.expander("Шаблон"):
            st.write(f"Используется {TEMPLATE_NAME} рядом с app.py.")
            up_t = st.file_uploader("Заменить шаблон", type=["xlsx"], key="tpl2")
            if up_t is not None:
                template_bytes = up_t.getvalue()

    pdf = st.file_uploader("PDF-протокол испытаний", type=["pdf"], key="pdf")

    if st.button("Заполнить журнал", type="primary", use_container_width=True,
                 disabled=(pdf is None or template_bytes is None)):
        try:
            with st.spinner("Обработка протокола…"):
                data, meta, log = process(pdf.getvalue(), template_bytes)
            st.session_state["res"] = (data, meta, log, pdf.name)
        except Exception as e:  # noqa: BLE001
            st.session_state.pop("res", None)
            st.error(f"Не удалось обработать файл: {e}")

    if template_bytes is None:
        st.info("Положите template.xlsx рядом с app.py или загрузите шаблон выше.")
    elif pdf is None:
        st.info("Загрузите PDF-протокол.")

    res = st.session_state.get("res")
    if not res:
        return
    data, meta, log, pdf_name = res

    st.success("Журнал готов")
    name = f'Журнал_{safe_name(meta["brand"])}_{safe_name(meta["model"])}_{meta["vin"]}.xlsx'
    st.download_button(
        "⬇️ Скачать журнал (.xlsx)", data=data, file_name=name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    st.subheader("Реквизиты")
    d = meta["date"]
    st.markdown(
        f"- **Протокол:** {meta['protocol'] or '—'}\n"
        f"- **Марка / модель:** {meta['brand']} / {meta['model']}\n"
        f"- **VIN:** {meta['vin'] or '—'}\n"
        f"- **Дата:** {d.strftime('%d.%m.%Y') if d else '—'}\n"
        f"- **Условия:** {meta['temp']} °C, {meta['hum']} %, {meta['press']} кПа, {meta['wind']} м/с"
    )

    c1, c2, c3 = st.columns(3)
    c1.metric("Заполнено", len(log["filled"]))
    c2.metric("Очищено", len(log["cleared"]))
    c3.metric("Округл.", len(log["inexact"]))

    if log["nonconf"]:
        st.error("Несоответствия (N) в протоколе: " + "; ".join(log["nonconf"]))
    if log["missing_cells"]:
        st.warning("Не найдены ячейки в шаблоне: " + ", ".join(sorted(set(log["missing_cells"]))))
    if log["dup"]:
        st.warning("Повторяющиеся параметры в протоколе (взят первый): " + ", ".join(log["dup"]))

    with st.expander(f"Очищено — нет в протоколе ({len(log['cleared'])})"):
        st.write("\n".join(f"- {x}" for x in log["cleared"]) or "—")
    with st.expander(f"Нет места в журнале ({len(log['unmatched'])})"):
        st.write("\n".join(f"- {x[:140]}" for x in log["unmatched"]) or "—")
    with st.expander(f"Заполненные блоки ({len(log['filled'])})"):
        for label, val, trip, exact, ok in log["filled"]:
            mark = "" if exact else " ≈"
            bad = "" if ok else " ⚠️N"
            st.write(f"**{label}**: {fmt(val)} → {' / '.join(trip)}{mark}{bad}")
    st.caption(
        "≈ — среднее совпадает с протоколом с учётом его округления: при сохранении шага "
        "разброса из шаблона точное значение невозможно. Формулы пересчитаются при открытии файла в Excel."
    )


if __name__ == "__main__":
    main()
