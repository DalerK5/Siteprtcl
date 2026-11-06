# -*- coding: utf-8 -*-
"""
Заполнение рабочего журнала сертификационных испытаний ТС категории M1
по PDF-протоколу испытаний (шаблон: template.xlsx).

Запуск:  streamlit run app.py
Зависимости: streamlit, openpyxl, pypdf
"""
import io
import functools
import itertools
import math
import re
import zipfile
import warnings
from datetime import date
from decimal import Decimal as D, InvalidOperation, ROUND_HALF_UP
from fractions import Fraction as Fr
from pathlib import Path
from xml.sax.saxutils import escape

import openpyxl
from openpyxl.utils.cell import column_index_from_string, coordinate_from_string, get_column_letter
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
    "cab_co": ("Воздух кабины: оксид углерода CO", 191, "F"),
    "cab_no2": ("Воздух кабины: диоксид азота NO2", 196, "F"),
    "cab_no": ("Воздух кабины: оксид азота NO", 201, "F"),
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
HOMO = str.maketrans("abcehkmoptxy", "авсенкмортху")


def norm(s):
    s = s.replace("\u0451", "\u0435").replace("\u0401", "\u0415")
    return re.sub(r"\s+", " ", s).strip()


def fold(s):
    """Нижний регистр, ё->е, латинские двойники -> кириллица."""
    return norm(s).lower().translate(HOMO)


def pdf_text(pdf_bytes):
    try:
        reader = PdfReader(io.BytesIO(pdf_bytes))
        return "\n".join((p.extract_text() or "") for p in reader.pages)
    except Exception as e:  # noqa: BLE001
        raise ValueError("файл не читается как PDF (%s)" % type(e).__name__)


def to_dec(s):
    return D(s.replace(",", ".").replace("\u2212", "-").strip())


def decimals_of(s):
    s = s.replace(",", ".")
    return len(s.split(".")[1]) if "." in s else 0


def parse_meta(raw):
    flat = norm(raw)
    meta = {}
    m = re.search(r"№\s*(SB-[\w\-/]+)", flat)
    meta["protocol"] = m.group(1) if m else ""
    m = re.search(r"марк[аиеу]?\s*[«\"“]([^»\"”]+)[»\"”]", flat)
    meta["brand"] = m.group(1).strip() if m else ""
    m = re.search(r"модел[ьиюе]\s*[«\"“]([^»\"”]+)[»\"”]", flat)
    meta["model"] = m.group(1).strip() if m else ""
    if not meta["brand"] or not meta["model"]:
        pats = [
            # «марки DONGFENG модели EII008,»
            r"марк[аиеу]?\s+([^,«»\"”]+?)\s+модел[ьиюе]?\s+([^,«»\"”]+?)\s*(?:,|VIN|идент|$)",
            # «markasi YUTONG, model: ZK6890HQG,»
            r"markasi\s+([^,]+?)\s*,?\s+model[i]?\s*:?\s*([^,\s]+(?:\s[^,\s:]+)?)\s*(?:,|VIN)",
            # «FOTON markali BJ6931 modelli»
            r"([A-Za-zА-Яа-я0-9\-]+)\s+markali\s+(.+?)\s+modelli",
            # «marka HONGQI E-QM5, VIN:» — марка = первое слово, модель = остальное
            r"\bmarka\s+([A-Za-zА-Яа-я0-9\-]+)\s+([^,]+?)\s*,\s*VIN",
        ]
        for pat in pats:
            m = re.search(pat, flat, re.I)
            if m:
                meta["brand"] = meta["brand"] or m.group(1).strip()
                meta["model"] = meta["model"] or m.group(2).strip()
                break
    m = (re.search(r"(?:VIN|идент\.?|идентификационный)\s*(?:№|N|номер)?\s*[:\-]?\s*([A-HJ-NPR-Za-hj-npr-z0-9]{17})\b", flat)
         or re.search(r"\b([A-HJ-NPR-Z0-9]{17})\b", flat))
    meta["vin"] = m.group(1).upper() if m else ""
    m = (re.search(r"Sinov boshlanish vaqti:\s*(\d{2})\.(\d{2})\.(\d{4})", flat)
         or re.search(r"Sanasi:\s*(\d{2})\.(\d{2})\.(\d{4})", flat))
    meta["date"] = date(int(m.group(3)), int(m.group(2)), int(m.group(1))) if m else None

    def num(pat):
        mm = re.search(pat, flat)
        return (to_dec(mm.group(1)), (mm.group(2) or "")) if mm else (None, "")

    # диапазоны вида «20 - 21 ℃» — берём первое число
    flat = re.sub(r"(:\s*-?\d+[.,]?\d*)\s*[-–]\s*\d+[.,]?\d*", r"\1", flat)

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


def parse_req(text):
    """Границы требования (нижняя, верхняя) — только явные."""
    t = norm(text).lower().replace(",", ".")
    lo = hi = None
    m = re.search(r"min\s*(-?\d+\.?\d*)", t)
    if m:
        lo = D(m.group(1))
    m = re.search(r"max\s*(-?\d+\.?\d*)", t)
    if m:
        hi = D(m.group(1))
    m = re.search(r"(?:не менее|не ниже|минимум)[^\d]{0,15}(\d+\.?\d*)", t)
    if m:
        lo = D(m.group(1))
    m = re.search(r"(?:не более|не превышает|не должн\w* превышать|максимум)[^\d]{0,15}(\d+\.?\d*)", t)
    if m:
        hi = D(m.group(1))
    m = re.search(r"от\s*(\d+\.?\d*)\s*%?\s*до\s*(\d+\.?\d*)", t)
    if m:
        lo, hi = D(m.group(1)), D(m.group(2))
    m = re.match(r"^\s*(\d+\.?\d*)\s*-\s*(\d+\.?\d*)", t)
    if m:
        lo, hi = D(m.group(1)), D(m.group(2))
    return lo, hi


def in_req(v, req):
    lo, hi = req
    if lo is None and hi is None:
        return None
    return (lo is None or v >= lo) and (hi is None or v <= hi)


def parse_params(raw):
    """Список параметров протокола: раздел, название, строки результатов."""
    text = re.sub(r"Sinov bayonnomasi № SB-.*?etiladi\.", "", raw, flags=re.S)
    events = []
    for m in re.finditer(r"(\d+)\.\s*Sinalayotgan ko.rsatgich\s+(.*?)\s*Texnik talablarni", text, re.S):
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
        name = norm(re.split(r"Maxsulotga", block, maxsplit=1)[0])
        name = re.sub(r"^\d+\s+(?=\D)", "", name)
        mr = re.search(r"birligi\s+(.*?)\s*\n\s*P\s*\n?\s*95", block, re.S)
        req = parse_req(mr.group(1)) if mr else (None, None)
        after = re.split(r"Xulosa", block, maxsplit=1)
        body = after[1] if len(after) == 2 else block
        rows = []
        for line in body.splitlines():
            mm = re.match(r"^\s*(\d+)\s+(.+?)\s+([MN])\s*$", line)
            if not mm:
                continue
            toks = mm.group(2).split()
            if len(toks) >= 3 and re.fullmatch(r"-?[\d.,]+", toks[0]) and re.fullmatch(r"-?[\d.,]+", toks[-2]):
                nums = [t.rstrip(",") for t in toks if re.fullmatch(r"-?\d+(?:[.,]+\d+)?,?", t)]
                unc = nums[1] if len(nums) >= 3 else None
                rows.append({"meas": toks[0], "final": toks[-2], "unit": toks[-1], "concl": mm.group(3),
                             "unc": unc})
        if not rows:
            continue
        params.append({"section": fold(section), "name": name, "raw_l": name.lower(),
                       "name_f": fold(name), "rows": rows, "req": req})
    return params


def row_value(row, req):
    """(Decimal, dp, note) — выбор между измеренным и итоговым значением."""
    try:
        m = to_dec(row["meas"])
    except (InvalidOperation, ValueError):
        m = None
    try:
        f = to_dec(row["final"])
    except (InvalidOperation, ValueError):
        f = None
    if m is None and f is None:
        raise ValueError("нечитаемое значение «%s»" % row["meas"])
    if f is None:
        return m, decimals_of(row["meas"]), "итог «%s» не читается — взято измеренное %s" % (row["final"], row["meas"])
    if m is None:
        return f, decimals_of(row["final"]), "измерено «%s» не читается — взято итоговое %s" % (row["meas"], row["final"])
    if m == f:
        return f, decimals_of(row["final"]), None
    cm, cf = in_req(m, req), in_req(f, req)
    if cf is True and cm is False:
        return f, decimals_of(row["final"]), "измерено %s, в итоге %s — взято итоговое" % (row["meas"], row["final"])
    return m, decimals_of(row["meas"]), "измерено %s, в итоге %s — взято измеренное" % (row["meas"], row["final"])


def param_unc(p):
    """Неопределённость из протокола: (Decimal, знаков) или (None, 0)."""
    for r in p["rows"]:
        if r.get("unc"):
            try:
                return to_dec(r["unc"]), decimals_of(r["unc"])
            except (InvalidOperation, ValueError):
                pass
    return None, 0


def param_value(p):
    vals, dps, notes = [], [], []
    for r in p["rows"]:
        v, dp, note = row_value(r, p["req"])
        vals.append(v)
        dps.append(dp)
        if note:
            notes.append(note)
    dp = max(dps)
    avg = sum(vals) / len(vals)
    return avg.quantize(D(1).scaleb(-dp)), dp, notes


# --------------------------------------------------------------------------
# Сопоставление параметров протокола с блоками журнала
# --------------------------------------------------------------------------
def detect_lamp(n):
    if "от передней части атс" in n:
        return "side"
    if "ближнего" in n and "фар" in n:
        return "dipped"
    if "заднего хода" in n:
        return "reverse"
    if "противотуман" in n:
        return "rear_fog" if "задн" in n else "fog_front"
    if "габаритн" in n:
        return "pos_front" if "передн" in n else ("pos_rear" if "задн" in n else None)
    if "сигнал торможения" in n or "стоп-сигнал" in n:
        return "brake"
    if "указател" in n and "поворот" in n:
        if "боков" in n or "повторител" in n:
            return "side"
        if "передн" in n:
            return "turn_front"
        if "задн" in n:
            return "turn_rear"
        return None
    if "световозвращател" in n:
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


def classify(section, n, raw_l, cur_lamp):
    """Возвращает (ключ блока или None, текущая лампа)."""
    in_light = ("освещен" in section or "световой сигнал" in section
                or ("светов" in section and "прибор" in section))
    lamp_here = detect_lamp(n)
    if lamp_here or in_light:
        lamp = lamp_here or (cur_lamp if in_light else None)
        col = lamp_col(n)
        if lamp and col:
            return lamp + "|" + col, lamp
        if in_light or lamp_here:
            return None, lamp
    # --- педали
    if "обозначенное в приложении 4" in n or "перегородки" in n or "педал" in section:
        m = re.search(r"букв\w*\s+([A-Za-zА-Яа-я])\s*$", norm(raw_l))
        if m:
            ch = m.group(1).upper()
            if ch in ("E", "Е"):
                return "ped_e", None
            if ch in ("H", "Н"):
                return "ped_h", None
            if ch == "J":
                return "ped_j", None
    # --- воздух кабины / выхлоп
    if "вредных веществ" in section or ("кабин" in section and "воздух" in section and "вентиляц" not in section):
        if "диоксид азота" in n or "no2" in n:
            return "cab_no2", None
        if "оксид азота" in n:
            return "cab_no", None
        if "углерода" in n:
            return "cab_co", None
    if "оксида углерода" in n or "оксид углерода" in n:
        return "co", None
    if "углеводород" in n:
        return "ch", None
    if "звуков" in n and ("сигнал" in n or "уровень" in n or "звука" in n):
        return "horn", None
    if "светопропускан" in section or "светопропускан" in n:
        if "ветров" in n:
            return "glass_w", None
        if "боков" in n:
            return "glass_s", None
    # --- номерной знак
    if "номерн" in section or "номерн" in n:
        if "продольн" in n:
            return "plate_a1", None
        if "опорн" in n and "высот" not in n:
            return "plate_a2", None
        if "нижнего края" in n:
            return "plate_h1", None
        if "верхнего края" in n:
            return "plate_h2", None
        if "высот" in n:
            return "plate_h", None
    # --- габариты
    if "максимальная длина" in n:
        return "length", None
    if "максимальная ширина" in n:
        return "width", None
    if "максимальная высота" in n:
        return "height", None
    # --- вентиляция / кондиционирование (по названию, без привязки к разделу)
    if "потоков на выходе из системы вентиляции" in n:
        return "vent_speed", None
    if "подвижность воздуха" in n or "система вентиляции должна" in n:
        return "vent_mob", None
    if "исключать возможность охлаждения" in n or ("охлаждения воздуха" in n and "голов" in n):
        return "ac_cool", None
    if "потока на выходе из системы кондиционирования" in n:
        return "ac_speed", None
    if n.startswith("температура воздуха должна"):
        return "ac_temp", None
    if "скорость воздуха в зоне головы" in n:
        return "ac_head", None
    if "относительная влажность" in n:
        return "humidity", None
    if "температура наружных поверхностей воздуховодов" in n:
        return "duct_temp", None
    if "температура поверхности без покрытия" in n:
        return "heat", None
    # --- шум
    if "уровень внешнего шума" in n or "уровень шума" in n:
        if "разгон" in n:
            return "noise", None
        if "постоянн" in n:
            return "noise_alt1", None
        if "неподвиж" in n or "стоян" in n:
            return "noise_alt2", None
        return "noise", None
    # --- травмобезопасность (по названию)
    if "ручек двери" in n:
        return "handle", None
    if "двери" in n and ("радиус" in n or n == "двери"):
        return "door", None
    if "бампер" in n:
        return "bumper", None
    if "обшивк" in n and ("кузов" in n):
        return "trim", None
    if "стеклоочистител" in n:
        return "wiper", None
    if "решетки и щели" in n or "для зазоров" in n:
        return "grille_gap", None
    if n.startswith("выступающая часть наружной поверхности"):
        return "protr", None
    return None, None


def map_params(params):
    """-> values {ключ: {...}}, не найденные, дубли, трасса сопоставления."""
    values, unmatched, dup, trace = {}, [], [], []
    lamp = None
    for p in params:
        key, lamp = classify(p["section"], p["name_f"], p["raw_l"], lamp)
        if key is None:
            unmatched.append(p)
            trace.append((p, None))
            continue
        ok = all(r["concl"] == "M" for r in p["rows"])
        if key == "plate_h":
            rows = p["rows"]
            targets = ["plate_h1", "plate_h2"] if len(rows) >= 2 else \
                [("plate_h2" if ("max" in p["raw_l"] or "макс" in p["name_f"]) else "plate_h1")]
            for k, r in zip(targets, rows):
                try:
                    v, dp, notes = param_value({"rows": [r], "req": p["req"]})
                except (ValueError, InvalidOperation):
                    continue
                unc, unc_dp = param_unc({"rows": [r]})
                values.setdefault(k, {"value": v, "dp": dp, "ok": r["concl"] == "M",
                                      "name": p["name"], "notes": notes, "unc": unc, "unc_dp": unc_dp})
            trace.append((p, "plate_h"))
            continue
        if key in values:
            dup.append(key)
            trace.append((p, key + " (повтор)"))
            continue
        try:
            v, dp, notes = param_value(p)
        except (ValueError, InvalidOperation):
            unmatched.append(p)
            trace.append((p, None))
            continue
        unc, unc_dp = param_unc(p)
        values[key] = {"value": v, "dp": dp, "ok": ok, "name": p["name"], "notes": notes,
                       "unc": unc, "unc_dp": unc_dp}
        trace.append((p, key))
    # шум: если нет режима разгона — берём доступный режим (с наибольшим значением)
    if "noise" not in values:
        alts = [values[k] for k in ("noise_alt1", "noise_alt2") if k in values]
        if alts:
            best = max(alts, key=lambda x: x["value"])
            best = dict(best)
            best["notes"] = list(best["notes"]) + ["нет режима разгона — в журнал взят: " + best["name"]]
            values["noise"] = best
    values.pop("noise_alt1", None)
    values.pop("noise_alt2", None)
    return values, unmatched, dup, trace


# --------------------------------------------------------------------------
# Генерация троек замеров по образцу разброса из шаблона
# --------------------------------------------------------------------------
def _sgn(x):
    return (x > 0) - (x < 0)


def _std(v):
    m = sum(v) / 3
    return math.sqrt(sum((float(x) - float(m)) ** 2 for x in v) / 2)


def _fit_exact(seed, shifted, P, k):
    """Подгоняет тройку на 1-2 единицы последнего знака так, чтобы среднее
    было РОВНО равно P, сохраняя характер разброса образца."""
    u = D(1).scaleb(-k)
    n = int(((3 * P - sum(shifted)) / u).to_integral_value())
    if n == 0:
        return shifted
    best, best_cost = None, None
    for e in itertools.product(range(-4, 5), repeat=3):
        if sum(e) != n:
            continue
        cand = [x + u * ei for x, ei in zip(shifted, e)]
        if min(cand) < 0 <= P:
            continue
        changes = sum(
            _sgn(seed[i] - seed[j]) != _sgn(cand[i] - cand[j]) for i, j in ((0, 1), (0, 2), (1, 2)))
        cost = (abs(_std(cand) - _std(seed)) / float(u) + 0.5 * changes
                + 0.1 * sum(abs(x) for x in e))
        if best_cost is None or cost < best_cost:
            best, best_cost = cand, cost
    return best


def gen_triplet(seed, P, dp):
    """Тройка с характером разброса образца; среднее РОВНО равно P.
    Возвращает (тройка, exact): exact=False — шаг разброса слегка изменён."""
    if seed and all(x is not None for x in seed):
        t = list(seed)
        k = max([max(-x.as_tuple().exponent, 0) for x in t] + [dp])
        at = Fr(sum(Fr(x) for x in t), 3)
        delta = Fr(P) - at
        if (delta * 10 ** k).denominator == 1:
            dl = D(delta.numerator) / D(delta.denominator)
            new, exact = [x + dl for x in t], True
        else:
            u = D(1).scaleb(-k)
            dl = (D(delta.numerator) / D(delta.denominator)).quantize(u)
            new = _fit_exact(t, [x + dl for x in t], P, k)
            exact = False
            if new is None:
                new = [x + dl for x in t]
        if min(new) < 0 <= P:  # отрицательных замеров быть не может
            return [P, P, P], True
        return new, exact
    step = D(1).scaleb(-dp)
    new = [P + step, P, P - step]
    if min(new) < 0 <= P:
        return [P, P, P], True
    return new, True


P_COVER = 1 - 0.9545          # как в шаблоне: T.INV.2T(1-0.9545; veff)


@functools.lru_cache(maxsize=None)
def _t95(df, p=P_COVER):
    return t_inv_2t(p, df)


def u_of_s(s, ub):
    """Расширенная неопределённость по формулам шаблона M1:
    UA=STDEVA/√3, u=√(UA²+UB²), veff=u⁴/(UA⁴/2), U=T.INV.2T(1-0.9545;veff)·u."""
    if s <= 0:
        return None
    ua = s / math.sqrt(3.0)
    uu = math.sqrt(ua * ua + ub * ub)
    veff = uu ** 4 / (ua ** 4 / 2.0)
    df = int(min(veff, 1e7))
    if df < 1:
        return None
    return _t95(df) * uu


def _round_eq(x, target, nd):
    q = D(1).scaleb(-nd)
    return D(repr(x)).quantize(q, rounding=ROUND_HALF_UP) == target.quantize(q)


def fit_u_triplet(seed, center, dp, unc, unc_dp, ub):
    """Тройка со средним РОВНО center и неопределённостью U, которая после округления
    до unc_dp знаков равна unc (как в протоколе). Характер (порядок) разброса — из шаблона.
    -> (тройка, добавлено_знаков, ok)"""
    target = float(unc)
    lo_u = u_of_s(1e-12, ub) or 0.0
    lo, hi = 1e-12, max(target, 1e-6)
    while (u_of_s(hi, ub) or 0) < target and hi < 1e12:
        hi *= 2.0
    for _ in range(200):
        mid = (lo + hi) / 2.0
        if (u_of_s(mid, ub) or 0) < target:
            lo = mid
        else:
            hi = mid
    s_t = hi
    if seed and all(x is not None for x in seed) and len(set(seed)) > 1:
        m = sum(float(x) for x in seed) / 3.0
        e = [float(x) - m for x in seed]
    else:
        e = [1.0, 0.0, -1.0]
    se = math.sqrt(sum(x * x for x in e) / 2.0)
    e = [x / se for x in e]
    seed_sign = [_sgn(e[i] - e[j]) for i, j in ((0, 1), (0, 2), (1, 2))]
    best_any = None
    for k in range(dp, dp + 5):
        q = 10.0 ** -k
        scale = s_t / q
        if scale < 0.5:
            continue
        n0, n1 = round(e[0] * scale), round(e[1] * scale)
        r = int(min(40, max(3, 0.08 * scale + 3)))
        found = None
        for dx in range(-r, r + 1):
            for dy in range(-r, r + 1):
                n = (n0 + dx, n1 + dy, -(n0 + dx + n1 + dy))
                if n[0] == n[1] == n[2]:
                    continue
                mn = sum(n) / 3.0
                s = q * math.sqrt(sum((x - mn) ** 2 for x in n) / 2.0)
                u = u_of_s(s, ub)
                if u is None:
                    continue
                vals = [center + D(x).scaleb(-k) for x in n]
                if min(vals) < 0 <= center:
                    continue
                ch = sum(_sgn(n[i] - n[j]) != seed_sign[t] for t, (i, j) in enumerate(((0, 1), (0, 2), (1, 2))))
                dist = abs(u - target)
                if best_any is None or dist < best_any[0]:
                    best_any = (dist, vals, k)
                if _round_eq(u, unc, unc_dp):
                    safe = dist <= 0.3 * 10.0 ** -unc_dp
                    cost = (not safe, ch, dist)
                    if found is None or cost < found[0]:
                        found = (cost, vals)
        if found:
            return found[1], k - dp, True
    if best_any:
        return best_any[1], best_any[2] - dp, False
    return None, 0, target >= lo_u


def fmt(x):
    s = format(x.normalize(), "f")
    return "0" if s in ("-0", "") else s


# --------------------------------------------------------------------------
# Мини-вычислитель формул Excel (чтобы в файле были готовые значения)
# --------------------------------------------------------------------------
class XlErr(Exception):
    def __init__(self, code):
        Exception.__init__(self, code)
        self.code = code


_TOK = re.compile(
    r"""\s*(?:
    (?P<num>\d+\.?\d*(?:[eE][+-]?\d+)?|\.\d+)
   |(?P<bool>TRUE|FALSE)(?![A-Za-z0-9_.(])
   |(?P<fn>[A-Za-z_][A-Za-z0-9_.]*)(?=\()
   |(?P<ref>(?:(?:'(?:[^']|'')+'|[^\s!()+\-*/^,:'"=<>&]+)!)?\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)
   |(?P<str>"(?:[^"]|"")*")
   |(?P<op><>|<=|>=|[-+*/^(),=<>])
    )""", re.X)


def _num(v):
    if v is None:
        return 0.0
    if isinstance(v, bool):
        return 1.0 if v else 0.0
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(",", "."))
    except ValueError:
        raise XlErr("#VALUE!")


def _betacf(a, b, x):
    tiny = 1e-300
    qab, qap, qam = a + b, a + 1.0, a - 1.0
    c, d = 1.0, 1.0 - qab * x / qap
    if abs(d) < tiny:
        d = tiny
    d = 1.0 / d
    h = d
    for m in range(1, 1000):
        m2 = 2 * m
        aa = m * (b - m) * x / ((qam + m2) * (a + m2))
        d = 1.0 + aa * d
        d = tiny if abs(d) < tiny else d
        c = 1.0 + aa / c
        c = tiny if abs(c) < tiny else c
        d = 1.0 / d
        h *= d * c
        aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2))
        d = 1.0 + aa * d
        d = tiny if abs(d) < tiny else d
        c = 1.0 + aa / c
        c = tiny if abs(c) < tiny else c
        d = 1.0 / d
        de = d * c
        h *= de
        if abs(de - 1.0) < 3e-15:
            break
    return h


def _betai(a, b, x):
    if x <= 0.0:
        return 0.0
    if x >= 1.0:
        return 1.0
    bt = math.exp(math.lgamma(a + b) - math.lgamma(a) - math.lgamma(b)
                  + a * math.log(x) + b * math.log(1.0 - x))
    if x < (a + 1.0) / (a + b + 2.0):
        return bt * _betacf(a, b, x) / a
    return 1.0 - bt * _betacf(b, a, 1.0 - x) / b


def t_inv_2t(p, df):
    if p <= 0.0 or p > 1.0 or df < 1:
        raise XlErr("#NUM!")
    df = int(df)
    if p == 1.0:
        return 0.0

    def tail(t):
        if df > 5000000:
            return math.erfc(t / math.sqrt(2.0))
        return _betai(df / 2.0, 0.5, df / (df + t * t))

    lo, hi = 0.0, 1.0
    while tail(hi) > p and hi < 1e9:
        hi *= 2.0
    for _ in range(200):
        mid = (lo + hi) / 2.0
        if tail(mid) > p:
            lo = mid
        else:
            hi = mid
    return (lo + hi) / 2.0


class Book:
    def __init__(self, cells):
        self.cells = cells
        self.memo = {}
        self.stack = set()

    def raw(self, sheet, coord):
        if sheet not in self.cells:
            raise XlErr("#REF!")
        return self.cells[sheet].get(coord)

    def result(self, sheet, coord):
        """('v', значение) или ('e', код) для формульной ячейки."""
        key = (sheet, coord)
        if key in self.memo:
            return self.memo[key]
        if key in self.stack:
            return ("e", "#REF!")
        self.stack.add(key)
        try:
            v = Evaluator(self, sheet, self.cells[sheet][coord][1:], coord).run()
            r = ("v", 0.0 if v is None else v)
        except XlErr as e:
            r = ("e", e.code)
        except ZeroDivisionError:
            r = ("e", "#DIV/0!")
        except (ValueError, OverflowError):
            r = ("e", "#NUM!")
        finally:
            self.stack.discard(key)
        self.memo[key] = r
        return r

    def value(self, sheet, coord):
        v = self.raw(sheet, coord)
        if isinstance(v, str) and v.startswith("="):
            r = self.result(sheet, coord)
            if r[0] == "e":
                raise XlErr(r[1])
            return r[1]
        return v

    def rng(self, sheet, a, b):
        c1, r1 = coordinate_from_string(a)
        c2, r2 = coordinate_from_string(b)
        i1, i2 = sorted((column_index_from_string(c1), column_index_from_string(c2)))
        j1, j2 = sorted((r1, r2))
        return [self.value(sheet, get_column_letter(i) + str(j))
                for j in range(j1, j2 + 1) for i in range(i1, i2 + 1)]


class Evaluator:
    def __init__(self, book, sheet, text, coord=None):
        self.b, self.sheet, self.coord = book, sheet, coord
        self.t, self.i = [], 0
        pos = 0
        text = text.strip()
        while pos < len(text):
            m = _TOK.match(text, pos)
            if not m or m.end() == pos:
                raise XlErr("#NAME?")
            pos = m.end()
            kind = m.lastgroup
            self.t.append((kind, m.group(kind)))
            while pos < len(text) and text[pos].isspace():
                pos += 1

    def peek(self):
        return self.t[self.i] if self.i < len(self.t) else (None, None)

    def take(self):
        tok = self.peek()
        self.i += 1
        return tok

    def run(self):
        v = self.comp()
        if self.i < len(self.t):
            raise XlErr("#VALUE!")
        if isinstance(v, list):
            v = self.intersect()
        return v

    def intersect(self):
        """Неявное пересечение: формула-диапазон «='48-04'!F1:G1» в ячейке F1 даёт F1."""
        kind, txt = self.t[0]
        if len(self.t) != 1 or kind != "ref" or not self.coord:
            raise XlErr("#VALUE!")
        sheet, rest = self.sheet, txt
        if "!" in txt:
            sh, rest = txt.rsplit("!", 1)
            sheet = sh[1:-1].replace("''", "'") if sh.startswith("'") else sh
        a, b = rest.replace("$", "").split(":")
        (c1, r1), (c2, r2) = coordinate_from_string(a), coordinate_from_string(b)
        cc, rr = coordinate_from_string(self.coord)
        i1, i2 = sorted((column_index_from_string(c1), column_index_from_string(c2)))
        j1, j2 = sorted((r1, r2))
        ci = column_index_from_string(cc)
        if j1 == j2 and i1 <= ci <= i2:
            return self.b.value(sheet, get_column_letter(ci) + str(j1))
        if i1 == i2 and j1 <= rr <= j2:
            return self.b.value(sheet, get_column_letter(i1) + str(rr))
        raise XlErr("#VALUE!")

    def comp(self):
        v = self.expr()
        ops = {"=": lambda a, b: a == b, "<>": lambda a, b: a != b, "<": lambda a, b: a < b,
               ">": lambda a, b: a > b, "<=": lambda a, b: a <= b, ">=": lambda a, b: a >= b}
        while self.peek()[0] == "op" and self.peek()[1] in ops:
            op = self.take()[1]
            r = self.expr()
            if isinstance(v, str) or isinstance(r, str):
                a, b = ("" if v is None else str(v)), ("" if r is None else str(r))
            else:
                a, b = _num(v), _num(r)
            v = ops[op](a, b)
        return v

    def expr(self):
        v = self.term()
        while self.peek() in (("op", "+"), ("op", "-")):
            op = self.take()[1]
            r = self.term()
            v = _num(v) + _num(r) if op == "+" else _num(v) - _num(r)
        return v

    def term(self):
        v = self.power()
        while self.peek() in (("op", "*"), ("op", "/")):
            op = self.take()[1]
            r = self.power()
            if op == "*":
                v = _num(v) * _num(r)
            else:
                d = _num(r)
                if d == 0:
                    raise ZeroDivisionError
                v = _num(v) / d
        return v

    def power(self):
        v = self.unary()
        while self.peek() == ("op", "^"):
            self.take()
            r = self.unary()
            a, e = _num(v), _num(r)
            if a == 0 and e < 0:
                raise ZeroDivisionError
            v = a ** e
            if isinstance(v, complex):
                raise XlErr("#NUM!")
        return v

    def unary(self):
        if self.peek() == ("op", "-"):
            self.take()
            return -_num(self.unary())
        if self.peek() == ("op", "+"):
            self.take()
            return self.unary()
        return self.atom()

    def atom(self):
        kind, txt = self.take()
        if kind == "num":
            return float(txt)
        if kind == "bool":
            return txt == "TRUE"
        if kind == "str":
            return txt[1:-1].replace('""', '"')
        if kind == "op" and txt == "(":
            v = self.comp()
            if self.take() != ("op", ")"):
                raise XlErr("#NAME?")
            return v
        if kind == "fn" and txt.upper() == "IF":
            return self.call_if()
        if kind == "fn":
            if self.take() != ("op", "("):
                raise XlErr("#NAME?")
            args = []
            if self.peek() != ("op", ")"):
                while True:
                    args.append(self.comp())
                    tok = self.take()
                    if tok == ("op", ")"):
                        break
                    if tok != ("op", ","):
                        raise XlErr("#NAME?")
            else:
                self.take()
            return self.call(txt, args)
        if kind == "ref":
            sheet, rest = self.sheet, txt
            if "!" in txt:
                sh, rest = txt.rsplit("!", 1)
                sheet = sh[1:-1].replace("''", "'") if sh.startswith("'") else sh
            rest = rest.replace("$", "")
            if ":" in rest:
                a, b = rest.split(":")
                return self.b.rng(sheet, a, b)
            return self.b.value(sheet, rest)
        raise XlErr("#NAME?")

    def skip_arg(self):
        """Пропустить аргумент функции (до «,» или «)» своего уровня)."""
        depth = 0
        while self.i < len(self.t):
            k, v = self.peek()
            if k == "op" and v == "(":
                depth += 1
            elif k == "op" and v == ")":
                if depth == 0:
                    return
                depth -= 1
            elif k == "op" and v == "," and depth == 0:
                return
            self.take()

    def call_if(self):
        """IF(условие, да, нет) — вычисляется только нужная ветка (как в Excel)."""
        if self.take() != ("op", "("):
            raise XlErr("#NAME?")
        cond = self.comp()
        c = cond if isinstance(cond, bool) else _num(cond) != 0
        if self.take() != ("op", ","):
            raise XlErr("#NAME?")
        res = False
        if c:
            res = self.comp()
        else:
            self.skip_arg()
        tok = self.take()
        if tok == ("op", ","):
            if c:
                self.skip_arg()
            else:
                res = self.comp()
            tok = self.take()
        if tok != ("op", ")"):
            raise XlErr("#NAME?")
        return res

    def call(self, name, args):
        n = name.upper().replace("_XLFN.", "")
        if n in ("TRUE", "FALSE") and not args:
            return n == "TRUE"
        if n == "COUNT":
            cnt = 0
            for a in args:
                for x in (a if isinstance(a, list) else [a]):
                    if isinstance(x, (int, float)) and not isinstance(x, bool):
                        cnt += 1
            return float(cnt)
        if n in ("AVERAGE", "STDEVA"):
            vals = []
            for a in args:
                if isinstance(a, list):
                    for x in a:
                        if isinstance(x, bool) or x is None:
                            continue
                        if isinstance(x, (int, float)):
                            vals.append(float(x))
                        elif n == "STDEVA":
                            vals.append(0.0)
                else:
                    vals.append(_num(a))
            if n == "AVERAGE":
                if not vals:
                    raise ZeroDivisionError
                return sum(vals) / len(vals)
            if len(vals) < 2:
                raise ZeroDivisionError
            if max(vals) == min(vals):     # одинаковые значения — ровно 0, как в Excel
                return 0.0
            m = math.fsum(vals) / len(vals)
            return math.sqrt(sum((x - m) ** 2 for x in vals) / (len(vals) - 1))
        if n == "SQRT":
            x = _num(args[0])
            if x < 0:
                raise XlErr("#NUM!")
            return math.sqrt(x)
        if n in ("T.INV.2T", "TINV"):
            return t_inv_2t(_num(args[0]), _num(args[1]))
        if n == "NORM.DIST":
            x, mu, sd = _num(args[0]), _num(args[1]), _num(args[2])
            if sd <= 0:
                raise XlErr("#NUM!")
            z = (x - mu) / sd
            if args[3] in (True, 1, 1.0):
                return 0.5 * math.erfc(-z / math.sqrt(2.0))
            return math.exp(-0.5 * z * z) / (sd * math.sqrt(2.0 * math.pi))
        raise XlErr("#NAME?")


def load_cells(xlsx_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), read_only=True, data_only=False)
    cells = {}
    for ws in wb.worksheets:
        try:
            ws.reset_dimensions()
        except Exception:  # noqa: BLE001
            pass
        d = {}
        for row in ws.iter_rows():
            for c in row:
                v = getattr(c, "value", None)
                if v is not None:
                    d[c.coordinate] = v
        cells[ws.title] = d
    wb.close()
    return cells


# --------------------------------------------------------------------------
# Работа с xlsx как с zip/XML (картинки и стили остаются нетронутыми)
# --------------------------------------------------------------------------
def cell_re(ref):
    return re.compile(r'<c r="%s"((?: [A-Za-z:]+="[^"]*")*)\s*(?:/>|>.*?</c>)' % ref, re.S)


T_ATTR_RE = re.compile(r' t="[^"]*"')


def set_num(xml, ref, val):
    m = cell_re(ref).search(xml)
    if not m:
        return xml, False
    attrs = T_ATTR_RE.sub("", m.group(1))
    return xml[:m.start()] + '<c r="%s"%s><v>%s</v></c>' % (ref, attrs, val) + xml[m.end():], True


def set_str(xml, ref, text):
    m = cell_re(ref).search(xml)
    if not m:
        return xml, False
    attrs = T_ATTR_RE.sub("", m.group(1))
    new = '<c r="%s"%s t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (ref, attrs, escape(text))
    return xml[:m.start()] + new + xml[m.end():], True


def clear_cell(xml, ref):
    m = cell_re(ref).search(xml)
    if not m or "<f" in m.group(0):
        return xml, False
    attrs = T_ATTR_RE.sub("", m.group(1))
    return xml[:m.start()] + '<c r="%s"%s/>' % (ref, attrs) + xml[m.end():], True


FORMULA_RE = re.compile(
    r'<c r="([A-Z]+\d+)"((?: [A-Za-z:]+="[^"]*")*)>(<f[^>]*?(?:/>|>[^<]*</f>))(?:<v(?: [^>]*)?>[^<]*</v>|<v/>)?</c>')


def strip_cache(xml):
    return FORMULA_RE.sub(
        lambda m: '<c r="%s"%s>%s</c>' % (m.group(1), T_ATTR_RE.sub("", m.group(2)), m.group(3)), xml)


def inject_values(xml, sheet_name, book):
    def repl(m):
        ref, attrs, f = m.group(1), T_ATTR_RE.sub("", m.group(2)), m.group(3)
        try:
            r = book.result(sheet_name, ref)
        except KeyError:
            return '<c r="%s"%s>%s</c>' % (ref, attrs, f)
        if r[0] == "e":
            return '<c r="%s"%s t="e">%s<v>%s</v></c>' % (ref, attrs, f, r[1])
        v = r[1]
        if isinstance(v, bool):
            return '<c r="%s"%s t="b">%s<v>%d</v></c>' % (ref, attrs, f, int(v))
        if isinstance(v, (int, float)):
            return '<c r="%s"%s>%s<v>%s</v></c>' % (ref, attrs, f, repr(float(v)))
        return '<c r="%s"%s t="str">%s<v>%s</v></c>' % (ref, attrs, f, escape(str(v)))
    return FORMULA_RE.sub(repl, xml)


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


def read_seeds(cells):
    """Значения троек из шаблона (образец разброса)."""
    j = cells.get(JOURNAL_SHEET, {})
    seeds = {}
    for key, (label, row, col) in all_blocks().items():
        vals = []
        for k in range(3):
            v = j.get(col + str(row + k))
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                vals.append(D(repr(v)))
            else:
                vals.append(None)
        seeds[key] = vals
    return seeds


def pack(zin, out):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as zout:
        for info in zin.infolist():
            if info.filename.endswith("/"):
                continue
            data = out.get(info.filename)
            if data is None:
                data = zin.read(info.filename)
            zout.writestr(info.filename, data)
    return buf.getvalue()


ID_REF = "K2"                  # «ID №…» (название PDF-файла) — справа от даты в шапке журнала
UB_STEPS = ("0.3", "0.2", "0.1", "0.05", "0.03", "0.02", "0.01", "0.005", "0.003", "0.002", "0.001",
            "0.0005", "0.0002", "0.0001")


def set_formula(xml, ref, formula):
    m = cell_re(ref).search(xml)
    if not m:
        return xml, False
    attrs = T_ATTR_RE.sub("", m.group(1))
    return xml[:m.start()] + '<c r="%s"%s><f>%s</f></c>' % (ref, attrs, escape(formula)) + xml[m.end():], True


def pack_out(zin, out):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as zout:
        for info in zin.infolist():
            if info.filename.endswith("/"):
                continue
            if info.filename in out and out[info.filename] is None:
                continue
            data = out.get(info.filename)
            if data is None:
                data = zin.read(info.filename)
            zout.writestr(info.filename, data)
    return buf.getvalue()


def drop_calc_chain(zin, out):
    """calcChain.xml удаляется (Excel построит его сам) — часть формул в файле изменена."""
    if "xl/calcChain.xml" not in set(zin.namelist()):
        return
    out["xl/calcChain.xml"] = None
    ct = zin.read("[Content_Types].xml").decode("utf8")
    out["[Content_Types].xml"] = re.sub(r'<Override PartName="/xl/calcChain\.xml"[^>]*/>', "", ct).encode("utf8")
    rels = zin.read("xl/_rels/workbook.xml.rels").decode("utf8")
    out["xl/_rels/workbook.xml.rels"] = re.sub(
        r'<Relationship\b[^>]*Target="[^"]*calcChain\.xml"[^>]*/>', "", rels).encode("utf8")


def _nf(f):
    return re.sub(r"[()\s$]", "", f) if isinstance(f, str) else ""


def block_ub(cells, book, row, col):
    """UB блока по цепочке: журнал (col, r+4) -> лист!G22 -> =D21 (=t*u) -> u=√(UA²+UB²) -> UB.
    -> (UB, лист, ячейка UB) или None, если цепочка нестандартная."""
    try:
        f = cells[JOURNAL_SHEET].get("%s%d" % (col, row + 4))
        m = re.fullmatch(r"='?((?:[^']|'')+?)'?!\$?([A-Z]+)\$?(\d+)", f or "")
        if not m:
            return None
        sh = m.group(1).replace("''", "'")
        cell = m.group(2) + m.group(3)
        f = cells[sh].get(cell)
        for _ in range(5):
            if isinstance(f, str) and re.fullmatch(r"=\$?[A-Z]+\$?\d+", f):
                cell = f[1:].replace("$", "")
                f = cells[sh].get(cell)
        mm = re.fullmatch(r"=([A-Z]+\d+)\*([A-Z]+\d+)", _nf(f))
        if not mm:
            return None
        a, b = mm.groups()
        t, u = (a, b) if "INV" in _nf(cells[sh].get(a)) else (b, a)
        if not re.fullmatch(r"=_xlfn\.T\.INV\.2T1-0\.9545,[A-Z]+\d+", _nf(cells[sh].get(t))):
            return None
        m3 = re.fullmatch(r"=SQRT([A-Z]+\d+)\^2\+([A-Z]+\d+)\^2", _nf(cells[sh].get(u)))
        if not m3:
            return None
        x, y = m3.groups()
        ua, ub = (x, y) if "STDEV" in _nf(cells[sh].get(x)) else (y, x)
        if not re.fullmatch(r"=STDEVA[A-Z]+\d+:[A-Z]+\d+/SQRT(?:3|[A-Z]+\d+)", _nf(cells[sh].get(ua))):
            return None
        v = book.value(sh, ub)
        return (float(v), sh, ub) if isinstance(v, (int, float)) and not isinstance(v, bool) else None
    except Exception:  # noqa: BLE001
        return None


class Styles:
    """Добавление форматов чисел в styles.xml (копия стиля ячейки с другим числовым форматом)."""
    BUILTIN = {"General": 0, "0": 1, "0.00": 2}

    def __init__(self, xml):
        self.xml = xml
        m = re.search(r"<cellXfs\b[^>]*>(.*?)</cellXfs>", xml, re.S)
        self.xfs = re.findall(r"<xf\b[^>]*/>|<xf\b[^>]*>.*?</xf>", m.group(1), re.S) if m else []
        self.fmts = {code: int(i) for i, code in re.findall(r'<numFmt numFmtId="(\d+)" formatCode="([^"]*)"\s*/>', xml)}
        self.new_fmts, self.new_xfs, self.cache = [], [], {}

    def fmt_id(self, code):
        if code in self.BUILTIN:
            return self.BUILTIN[code]
        if code not in self.fmts:
            nid = max([163] + list(self.fmts.values())) + 1
            self.fmts[code] = nid
            self.new_fmts.append('<numFmt numFmtId="%d" formatCode="%s"/>' % (nid, code))
        return self.fmts[code]

    def with_format(self, s, code):
        key = (s, code)
        if key in self.cache:
            return self.cache[key]
        if s >= len(self.xfs):
            return s
        fid = self.fmt_id(code)
        xf = self.xfs[s]
        if 'numFmtId="' in xf:
            xf = re.sub(r'numFmtId="\d+"', 'numFmtId="%d"' % fid, xf, count=1)
        else:
            xf = xf.replace("<xf ", '<xf numFmtId="%d" ' % fid, 1)
        if "applyNumberFormat=" in xf:
            xf = re.sub(r'applyNumberFormat="\d"', 'applyNumberFormat="1"', xf, count=1)
        else:
            xf = xf.replace("<xf ", '<xf applyNumberFormat="1" ', 1)
        idx = len(self.xfs) + len(self.new_xfs)
        self.new_xfs.append(xf)
        self.cache[key] = idx
        return idx

    def render(self):
        x = self.xml
        if self.new_fmts:
            if "<numFmts" in x:
                x = re.sub(r"</numFmts>", "".join(self.new_fmts) + "</numFmts>", x, count=1)
                x = re.sub(r'<numFmts count="\d+"', '<numFmts count="%d"' % len(self.fmts), x, count=1)
            else:
                x = re.sub(r"(<styleSheet\b[^>]*>)", r'\1<numFmts count="%d">%s</numFmts>'
                           % (len(self.new_fmts), "".join(self.new_fmts)), x, count=1)
        if self.new_xfs:
            x = re.sub(r"</cellXfs>", "".join(self.new_xfs) + "</cellXfs>", x, count=1)
            x = re.sub(r'<cellXfs count="\d+"', '<cellXfs count="%d"' % (len(self.xfs) + len(self.new_xfs)),
                       x, count=1)
        return x


def num_code(dp):
    return "0" if dp <= 0 else "0." + "0" * dp


def restyle(xml, ref, styles, code):
    """Поменять числовой формат ячейки (остальное оформление — как было)."""
    m = cell_re(ref).search(xml)
    if not m:
        return xml
    attrs = m.group(1)
    ms = re.search(r' s="(\d+)"', attrs)
    s = int(ms.group(1)) if ms else 0
    ns = styles.with_format(s, code)
    new_attrs = re.sub(r' s="\d+"', ' s="%d"' % ns, attrs) if ms else attrs + ' s="%d"' % ns
    head = '<c r="%s"%s' % (ref, attrs)
    return xml[:m.start()] + '<c r="%s"%s' % (ref, new_attrs) + xml[m.start() + len(head):]


def dec_count(s):
    return len(s.split(".")[1]) if "." in s else 0


def fill_workbook(template_bytes, meta, values):
    tcells = load_cells(template_bytes)
    if JOURNAL_SHEET not in tcells:
        raise ValueError("В шаблоне нет листа «%s»." % JOURNAL_SHEET)
    seeds = read_seeds(tcells)
    blocks = all_blocks()
    tbook = Book(tcells)
    log = {"filled": [], "cleared": [], "inexact": [], "missing_cells": [], "verify_bad": [], "calc": True,
           "u_fit": [], "u_bad": [], "u_skip": [], "ub_changed": []}
    ub_cells = {}   # лист -> {ячейка: новая формула UB}

    zin = zipfile.ZipFile(io.BytesIO(template_bytes))
    paths = sheet_paths(zin)
    jpath = next(p for n, p in paths if n == JOURNAL_SHEET)
    j = zin.read(jpath).decode("utf8")
    out = {}
    jc = dict(tcells[JOURNAL_SHEET])  # копия значений журнала для расчёта формул

    # --- шапка журнала
    d = meta["date"]
    if d:
        txt = "     Дата проведения испытаний: «%02d» %s %d г.      " % (d.day, MONTHS[d.month - 1], d.year)
        j, ok = set_str(j, "C2", txt)
        if not ok:
            log["missing_cells"].append("C2")
    for ref, v in (("E11", meta["brand"] + " "), ("E12", meta["model"]), ("E15", meta["vin"] + " ")):
        if v.strip():
            j, ok = set_str(j, ref, v)
            if not ok:
                log["missing_cells"].append(ref)
    if meta.get("id"):
        j, ok = set_str(j, ID_REF, "ID № %s" % meta["id"])
        if not ok:
            log["missing_cells"].append(ID_REF)
    for ref, v in (("D6", meta["temp"]), ("D7", meta["hum"]), ("D8", meta["press"]), ("D9", meta["wind"])):
        if v is not None:
            j, ok = set_num(j, ref, fmt(D(v)))
            if not ok:
                log["missing_cells"].append(ref)

    # --- замеры
    for key, (label, row, col) in blocks.items():
        cells = [col + str(row + k) for k in range(3)]
        if key in values:
            p = values[key]
            new, exact = None, True
            if p.get("unc") and p["unc"] > 0:
                ubi = block_ub(tcells, tbook, row, col)
                if ubi is None:
                    log["u_skip"].append(label)
                else:
                    ub0, ush, ucell = ubi
                    new, extra, uok = fit_u_triplet(seeds[key], p["value"], p["dp"], p["unc"], p["unc_dp"], ub0)
                    if new is None or not uok:
                        # с UB шаблона так не получается — берём ближайшее меньшее «круглое» UB
                        for a in UB_STEPS:
                            ub1 = float(a) / 2
                            if ub1 >= ub0:
                                continue
                            n2, e2, ok2 = fit_u_triplet(seeds[key], p["value"], p["dp"], p["unc"],
                                                        p["unc_dp"], ub1)
                            if n2 is not None and ok2:
                                ub_cells.setdefault(ush, {})[ucell] = "=%s/2" % a
                                log["ub_changed"].append("%s: UB %s → =%s/2 ('%s'!%s)" % (
                                    label, tcells[ush].get(ucell), a, ush.strip(), ucell))
                                new, extra, uok = n2, e2, ok2
                                break
                    if new is not None:
                        log["u_fit"].append((label, p["unc"], p["unc_dp"], extra, uok, key))
                    else:
                        log["u_bad"].append("%s: протокол %s, а формулы шаблона дают не меньше %s" % (
                            label, fmt(p["unc"]), round(u_of_s(1e-9, ub0), p["unc_dp"] + 2)))
            if new is None:
                new, exact = gen_triplet(seeds[key], p["value"], p["dp"])
            for ref, x in zip(cells, new):
                j, ok = set_num(j, ref, fmt(x))
                if ok:
                    jc[ref] = float(x)
                else:
                    log["missing_cells"].append(ref)
            if p.get("unc") is not None and p["unc"] == 0 and len(set(new)) == 1:
                # в протоколе U = 0 и все замеры одинаковые: в журнале тоже 0 (а не #DIV/0!)
                uref = "%s%d" % (col, row + 4)
                f = jc.get(uref)
                if isinstance(f, str) and f.startswith("="):
                    nf = "IF(STDEVA(%s%d:%s%d)=0,0,%s)" % (col, row, col, row + 2, f[1:])
                    j, okf = set_formula(j, uref, nf)
                    if okf:
                        jc[uref] = "=" + nf
                        log["u_fit"].append((label, p["unc"], p["unc_dp"], 0, True, key))
            log["filled"].append((label, p["value"], [fmt(x) for x in new], exact, p["ok"], key, p["dp"]))
            if not exact:
                log["inexact"].append(label)
        else:
            had = any(v is not None for v in seeds[key])
            for ref in cells:
                j, done = clear_cell(j, ref)
                if done:
                    jc.pop(ref, None)
            if had:
                log["cleared"].append(label)
            # среднее и неопределённость пустого блока — пустые (без #DIV/0!),
            # при ручном вводе замеров формулы снова считают
            guard = "COUNT(%s%d:%s%d)=0" % (col, row, col, row + 2)
            for ref in ("%s%d" % (col, row + 3), "%s%d" % (col, row + 4)):
                f = jc.get(ref)
                if isinstance(f, str) and f.startswith("=") and "COUNT(" not in f:
                    nf = 'IF(%s,"",%s)' % (guard, f[1:])
                    j, ok = set_formula(j, ref, nf)
                    if ok:
                        jc[ref] = "=" + nf
    # --- формат чисел: столько знаков после запятой, сколько в протоколе (среднее и U) и в замерах
    styles = Styles(zin.read("xl/styles.xml").decode("utf8"))
    for label_, val_, trip_, exact_, ok_, key_, dp_ in log["filled"]:
        b_ = blocks[key_]
        row, col = b_[1], (b_[2] if len(b_) > 2 else JCOL)
        for k_, t_ in enumerate(trip_):
            j = restyle(j, "%s%d" % (col, row + k_), styles, num_code(max(dec_count(t) for t in trip_)))
        j = restyle(j, "%s%d" % (col, row + 3), styles, num_code(dp_))
        pv = values[key_]
        if pv.get("unc") is not None:
            j = restyle(j, "%s%d" % (col, row + 4), styles, num_code(pv["unc_dp"]))
    out["xl/styles.xml"] = styles.render().encode("utf8")
    out[jpath] = j.encode("utf8")

    # --- расчётные листы: объект, дата, изменённые UB
    cal_name = ("%s %s" % (meta["brand"], meta["model"])).strip()
    serial = (d - date(1899, 12, 30)).days if d else None
    for n, p in paths:
        if p == jpath:
            continue
        x = zin.read(p).decode("utf8")
        if cal_name:
            x, _ = set_str(x, "D2", cal_name)
        if serial is not None:
            x, _ = set_num(x, "D4", str(serial))
        for ref, f in ub_cells.get(n, {}).items():
            x, _ = set_formula(x, ref, f[1:])
        out[p] = x.encode("utf8")

    wbx = zin.read("xl/workbook.xml").decode("utf8")
    if "<calcPr" in wbx:
        wbx = re.sub(r"<calcPr([^>]*?)/>", lambda m: "<calcPr" + re.sub(r' fullCalcOnLoad="[^"]*"', "", m.group(1))
                     + ' fullCalcOnLoad="1"/>', wbx, count=1)
    else:
        wbx = wbx.replace("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>')
    out["xl/workbook.xml"] = wbx.encode("utf8")
    drop_calc_chain(zin, out)

    # --- расчёт формул и запись готовых значений в файл
    try:
        allc = dict(tcells)
        allc[JOURNAL_SHEET] = jc
        for sh_, cc in ub_cells.items():
            allc[sh_] = dict(allc[sh_], **cc)
        book = Book(allc)
        for n, p in paths:
            out[p] = inject_values(out[p].decode("utf8"), n, book).encode("utf8")
        # контроль: среднее в журнале == значение протокола
        for label, val, trip, exact, ok, key, dp in log["filled"]:
            _, row, col = blocks[key]
            r = book.result(JOURNAL_SHEET, col + str(row + 3))
            if r[0] != "v" or not isinstance(r[1], (int, float)) or abs(D(repr(r[1])) - val) > D("1e-9"):
                log["verify_bad"].append(label)
        # контроль: неопределённость в журнале (строка +4) == протоколу после округления
        for label, unc, unc_dp, extra, uok, key in log["u_fit"]:
            _, row, col = blocks[key]
            r = book.result(JOURNAL_SHEET, "%s%d" % (col, row + 4))
            got = r[1] if r[0] == "v" and isinstance(r[1], (int, float)) else None
            if got is None or not _round_eq(got, unc, unc_dp):
                log["u_bad"].append("%s: протокол %s, в журнале %s" % (
                    label, fmt(unc), "ошибка" if got is None else round(got, unc_dp + 2)))
    except Exception:  # noqa: BLE001
        log["calc"] = False
        for n, p in paths:
            out[p] = strip_cache(out[p].decode("utf8")).encode("utf8")
    return pack_out(zin, out), log


def process(pdf_bytes, template_bytes, file_id=""):
    raw = pdf_text(pdf_bytes)
    if "Aniqlanayotgan parametr nomi" not in raw:
        raise ValueError("В PDF не найден текст протокола (возможно, это скан без текстового слоя).")
    meta = parse_meta(raw)
    meta["id"] = (file_id or "").strip()
    params = parse_params(raw)
    if not params:
        raise ValueError("В протоколе не удалось найти таблицы результатов.")
    values, unmatched, dup, trace = map_params(params)
    data, log = fill_workbook(template_bytes, meta, values)
    log["unmatched"] = [p["name"] for p in unmatched]
    log["trace"] = [(p["section"], p["name"], key, p["rows"][0]["final"]) for p, key in trace]
    log["suspect"] = [p["name"] for p in unmatched
                      if "освещающей поверхности" in p["name_f"] or "габарит" in p["name_f"]]
    log["dup"] = dup
    log["params"] = len(params)
    log["nonconf"] = [p["name"] for p in params if any(r["concl"] != "M" for r in p["rows"])]
    log["notes"] = [(v["name"], n) for v in values.values() for n in v["notes"]]
    return data, meta, log


def id_from_filename(fname):
    """ID = название PDF-файла: «2007212.pdf», «2007212 (1).pdf» -> «2007212»."""
    stem = Path(fname or "").stem
    m = re.search(r"\d{4,}", stem)
    return m.group(0) if m else stem.strip()


def safe_name(s):
    return re.sub(r"[^\w\-.]+", "_", s, flags=re.U).strip("_")


def out_name(meta):
    if meta.get("id"):
        return "%s.xlsx" % safe_name(meta["id"])
    return "Журнал_%s_%s_%s.xlsx" % (safe_name(meta["brand"]), safe_name(meta["model"]), meta["vin"])


def unique_ids(names):
    ids, seen = [], {}
    for nm in names:
        i = id_from_filename(nm) or "file"
        seen[i] = seen.get(i, 0) + 1
        ids.append(i if seen[i] == 1 else "%s_%d" % (i, seen[i]))
    return ids


def make_zip(results):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        used = set()
        for r in results:
            if r["data"] is None:
                continue
            nm = out_name(r["meta"])
            base, n = nm[:-5], 2
            while nm in used:
                nm, n = "%s_%d.xlsx" % (base, n), n + 1
            used.add(nm)
            z.writestr(nm, r["data"])
    return buf.getvalue()


def problems(meta, log):
    res = []
    if log["verify_bad"]:
        res.append("среднее не сошлось")
    if log.get("u_bad"):
        res.append("U не сошлась: %d" % len(log["u_bad"]))
    if not log["calc"]:
        res.append("формулы не посчитаны")
    if log["nonconf"]:
        res.append("есть N в протоколе")
    if log.get("suspect"):
        res.append("нераспознанные пункты")
    if not meta["brand"] or not meta["model"]:
        res.append("нет марки/модели")
    if log["notes"]:
        res.append("расхождения в протоколе: %d" % len(log["notes"]))
    if log.get("ub_changed"):
        res.append("UB изменена: %d" % len(log["ub_changed"]))
    return res


# --------------------------------------------------------------------------
# Интерфейс Streamlit
# --------------------------------------------------------------------------
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
MAX_FILES = 10


def main():
    st.set_page_config(page_title="Журнал M1", page_icon="📋", layout="centered")
    st.markdown(
        "<style>.stButton>button,.stDownloadButton>button{min-height:3rem;font-size:1.05rem}"
        ".block-container{padding-top:1.2rem;padding-bottom:2rem}</style>",
        unsafe_allow_html=True,
    )
    st.title("📋 Журнал испытаний M1")
    st.caption("Загрузите один или сразу несколько (до %d) PDF-протоколов — получите заполненные журналы."
               % MAX_FILES)

    tpl_path = Path(__file__).resolve().parent / TEMPLATE_NAME
    template_bytes = tpl_path.read_bytes() if tpl_path.exists() else None
    if template_bytes is None:
        up_t = st.file_uploader("Шаблон журнала (template.xlsx)", type=["xlsx"], key="tpl")
        if up_t is not None:
            template_bytes = up_t.getvalue()
    else:
        with st.expander("Шаблон"):
            st.write("Используется %s рядом с app.py." % TEMPLATE_NAME)
            up_t = st.file_uploader("Заменить шаблон", type=["xlsx"], key="tpl2")
            if up_t is not None:
                template_bytes = up_t.getvalue()

    pdfs = st.file_uploader("PDF-протоколы испытаний (можно несколько)", type=["pdf"],
                            accept_multiple_files=True, key="pdfs") or []
    if len(pdfs) > MAX_FILES:
        st.warning("Выбрано %d файлов — будут обработаны первые %d." % (len(pdfs), MAX_FILES))
        pdfs = pdfs[:MAX_FILES]
    ids = unique_ids([f.name for f in pdfs])
    if len(pdfs) == 1:
        ids[0] = st.text_input("ID (берётся из названия файла)", value=ids[0], key="id_" + pdfs[0].name)
    elif pdfs:
        st.caption("ID берутся из названий файлов: " + ", ".join(ids))

    label = "Заполнить журнал" if len(pdfs) <= 1 else "Заполнить журналы (%d)" % len(pdfs)
    if st.button(label, type="primary", use_container_width=True,
                 disabled=(not pdfs or template_bytes is None)):
        results = []
        bar = st.progress(0.0, text="Обработка…")
        for n, (f, fid) in enumerate(zip(pdfs, ids)):
            bar.progress(n / len(pdfs), text="Обработка %d из %d: %s" % (n + 1, len(pdfs), f.name))
            try:
                data, meta, log = process(f.getvalue(), template_bytes, fid)
                results.append({"name": f.name, "data": data, "meta": meta, "log": log, "err": None})
            except Exception as e:  # noqa: BLE001
                results.append({"name": f.name, "data": None, "meta": None, "log": None, "err": str(e)})
        bar.empty()
        st.session_state["batch"] = results

    if template_bytes is None:
        st.info("Положите template.xlsx рядом с app.py или загрузите шаблон выше.")
    elif not pdfs:
        st.info("Загрузите PDF-протокол(ы).")

    results = st.session_state.get("batch")
    if not results:
        return
    ok = [r for r in results if r["data"] is not None]

    if len(results) == 1:
        r = results[0]
        if r["err"]:
            st.error("Не удалось обработать файл: %s" % r["err"])
            return
        st.success("Журнал готов")
        st.download_button("⬇️ Скачать журнал (.xlsx)", data=r["data"], file_name=out_name(r["meta"]),
                           mime=XLSX_MIME, use_container_width=True)
        show_details(r["meta"], r["log"])
        return

    st.success("Готово: %d из %d" % (len(ok), len(results)))
    if ok:
        st.download_button("⬇️ Скачать все журналы (.zip, %d)" % len(ok), data=make_zip(results),
                           file_name="Журналы_M1.zip", mime="application/zip", use_container_width=True)
    st.subheader("Сводка")
    for n, r in enumerate(results):
        if r["err"]:
            st.error("❌ **%s** — не обработан: %s" % (r["name"], r["err"]))
            continue
        meta, log = r["meta"], r["log"]
        pr = problems(meta, log)
        line = "**%s** — %s %s · заполнено %d · U совпала %d/%d" % (
            meta.get("id") or r["name"], meta["brand"] or "?", meta["model"] or "?", len(log["filled"]),
            sum(1 for x in log.get("u_fit", []) if x[4]), len(log.get("u_fit", [])))
        hard = any(x.startswith(("среднее", "U не", "формулы", "нераспозн", "нет марки")) for x in pr)
        (st.warning if hard else st.success)(("⚠️ " if hard else "✅ ") + line +
                                             ("\n\n" + "; ".join(pr) if pr else ""))
        st.download_button("⬇️ %s" % out_name(meta), data=r["data"], file_name=out_name(meta),
                           mime=XLSX_MIME, use_container_width=True, key="dl_%d" % n)

    if ok:
        st.subheader("Подробности")
        names = [(r["meta"].get("id") or r["name"]) for r in ok]
        pick = st.selectbox("Файл", list(range(len(ok))), format_func=lambda i: names[i])
        show_details(ok[pick]["meta"], ok[pick]["log"])


def show_details(meta, log):
    d = meta["date"]
    st.subheader("Реквизиты")
    st.markdown(
        "- **ID:** %s\n- **Протокол:** %s\n- **Марка / модель:** %s / %s\n- **VIN:** %s\n- **Дата:** %s\n"
        "- **Условия:** %s °C, %s %%, %s кПа, %s м/с" % (
            meta.get("id") or "—", meta["protocol"] or "—", meta["brand"] or "—", meta["model"] or "—",
            meta["vin"] or "—", d.strftime("%d.%m.%Y") if d else "—",
            meta["temp"], meta["hum"], meta["press"], meta["wind"]))
    if not meta["brand"] or not meta["model"]:
        st.warning("Марка/модель не распознаны — впишите вручную (E11, E12).")

    c1, c2, c3 = st.columns(3)
    c1.metric("Заполнено", len(log["filled"]))
    c2.metric("Очищено", len(log["cleared"]))
    c3.metric("Шаг изменён", len(log["inexact"]))

    if log["verify_bad"]:
        st.error("Среднее в журнале не сошлось с протоколом: " + "; ".join(log["verify_bad"]))
    elif log["calc"]:
        st.success("Проверка: среднее значение каждого заполненного блока совпадает с протоколом.")
    else:
        st.warning("Значения формул не посчитаны: они пересчитаются при открытии файла в Excel.")
    if log.get("u_bad"):
        st.error("Неопределённость не удалось сделать как в протоколе — поставлено ближайшее возможное:\n" +
                 "\n".join("- %s" % x for x in log["u_bad"]))
    elif log.get("u_fit") and log["calc"]:
        st.success("Проверка: неопределённость совпадает с протоколом у %d блоков." % len(log["u_fit"]))
    if log.get("ub_changed"):
        st.warning("Чтобы неопределённость совпала с протоколом, в этих блоках уменьшена погрешность "
                   "прибора UB (в шаблоне она больше, чем допускает протокол). Сверьте с паспортом прибора:\n" +
                   "\n".join("- %s" % x for x in log["ub_changed"]))
    if log.get("u_skip"):
        st.warning("Для этих блоков неопределённость не подгонялась (нестандартные формулы шаблона): "
                   + "; ".join(log["u_skip"]))
    if log["notes"]:
        st.warning("Ошибки/расхождения в протоколе (проверьте):\n" +
                   "\n".join("- %s: %s" % (n[:70], t) for n, t in log["notes"]))
    if log["nonconf"]:
        st.error("Несоответствия (N) в протоколе: " + "; ".join(log["nonconf"]))
    if log["missing_cells"]:
        st.warning("Не найдены ячейки в шаблоне: " + ", ".join(sorted(set(log["missing_cells"]))))
    if log["dup"]:
        st.warning("Повторяющиеся параметры в протоколе (взят первый): " + ", ".join(log["dup"]))
    if log.get("suspect"):
        st.error("Похоже на пункты журнала, но не распознаны — проверьте вручную:\n" +
                 "\n".join("- %s" % x[:140] for x in log["suspect"]))
    if log["unmatched"]:
        with st.expander("Параметры протокола без места в журнале (%d)" % len(log["unmatched"])):
            st.write("\n".join("- %s" % x[:120] for x in log["unmatched"]))

    with st.expander("Все параметры протокола и куда они попали (%d)" % len(log.get("trace", []))):
        for sec, nm, key, fin in log.get("trace", []):
            st.write("%s `%s` → %s  (%s)" % ("✅" if key else "⛔", key or "нет места в журнале", nm[:90], fin))
    with st.expander("Очищено — нет в протоколе (%d)" % len(log["cleared"])):
        st.write("\n".join("- %s" % x for x in log["cleared"]) or "—")
    with st.expander("Заполненные блоки (%d)" % len(log["filled"])):
        for label, val, trip, exact, ok, key, dp in log["filled"]:
            st.write("**%s**: %s → %s%s%s" % (label, fmt(val), " / ".join(trip),
                                              "" if exact else " ≈", "" if ok else " ⚠️N"))
    with st.expander("Неопределённость: протокол → журнал (%d)" % len(log.get("u_fit", []))):
        for label, unc, unc_dp, extra, uok, key in log.get("u_fit", []):
            st.write("%s **%s**: U = %s%s" % ("✅" if uok else "⚠️", label, fmt(unc),
                                              (" (добавлено знаков после запятой: %d)" % extra) if extra else ""))
    st.caption("≈ — у этих блоков шаг разброса из шаблона слегка изменён (на 1–2 единицы последнего "
               "знака), чтобы среднее было точно равно значению протокола.")


if __name__ == "__main__":
    main()
