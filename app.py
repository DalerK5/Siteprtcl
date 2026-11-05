# -*- coding: utf-8 -*-
"""
Заполнение рабочего журнала сертификационных испытаний ТС категории M1
по PDF-протоколу испытаний (шаблон: template.xlsx).

Запуск:  streamlit run app.py
Зависимости: streamlit, openpyxl, pypdf
"""
import io
import itertools
import math
import re
import zipfile
import warnings
from datetime import date
from decimal import Decimal as D
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
        m = re.search(r"марк[аиеу]?\s+([^,«»\"”]+?)\s+модел[ьиюе]?\s+([^,«»\"”]+?)\s*(?:,|VIN|идент|$)", flat, re.I)
        if m:
            meta["brand"] = meta["brand"] or m.group(1).strip()
            meta["model"] = meta["model"] or m.group(2).strip()
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
                rows.append({"meas": toks[0], "final": toks[-2], "unit": toks[-1], "concl": mm.group(3)})
        if not rows:
            continue
        params.append({"section": fold(section), "name": name, "raw_l": name.lower(),
                       "name_f": fold(name), "rows": rows, "req": req})
    return params


def row_value(row, req):
    """(Decimal, dp, note) — выбор между измеренным и итоговым значением."""
    m, f = to_dec(row["meas"]), to_dec(row["final"])
    if m == f:
        return f, decimals_of(row["final"]), None
    cm, cf = in_req(m, req), in_req(f, req)
    if cf is True and cm is False:
        return f, decimals_of(row["final"]), "измерено %s, в итоге %s — взято итоговое" % (row["meas"], row["final"])
    return m, decimals_of(row["meas"]), "измерено %s, в итоге %s — взято измеренное" % (row["meas"], row["final"])


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
                v, dp, notes = param_value({"rows": [r], "req": p["req"]})
                values.setdefault(k, {"value": v, "dp": dp, "ok": r["concl"] == "M",
                                      "name": p["name"], "notes": notes})
            trace.append((p, "plate_h"))
            continue
        if key in values:
            dup.append(key)
            trace.append((p, key + " (повтор)"))
            continue
        v, dp, notes = param_value(p)
        values[key] = {"value": v, "dp": dp, "ok": ok, "name": p["name"], "notes": notes}
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
   |(?P<op>[-+*/^(),])
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
            v = Evaluator(self, sheet, self.cells[sheet][coord][1:]).run()
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
    def __init__(self, book, sheet, text):
        self.b, self.sheet = book, sheet
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
        v = self.expr()
        if self.i < len(self.t) or isinstance(v, list):
            raise XlErr("#VALUE!")
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
        if kind == "op" and txt == "(":
            v = self.expr()
            if self.take() != ("op", ")"):
                raise XlErr("#NAME?")
            return v
        if kind == "fn":
            if self.take() != ("op", "("):
                raise XlErr("#NAME?")
            args = []
            if self.peek() != ("op", ")"):
                while True:
                    args.append(self.expr())
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

    def call(self, name, args):
        n = name.upper().replace("_XLFN.", "")
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
            m = sum(vals) / len(vals)
            return math.sqrt(sum((x - m) ** 2 for x in vals) / (len(vals) - 1))
        if n == "SQRT":
            x = _num(args[0])
            if x < 0:
                raise XlErr("#NUM!")
            return math.sqrt(x)
        if n == "T.INV.2T":
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


def fill_workbook(template_bytes, meta, values):
    tcells = load_cells(template_bytes)
    if JOURNAL_SHEET not in tcells:
        raise ValueError("В шаблоне нет листа «%s»." % JOURNAL_SHEET)
    seeds = read_seeds(tcells)
    blocks = all_blocks()
    log = {"filled": [], "cleared": [], "inexact": [], "missing_cells": [], "verify_bad": [], "calc": True}

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
            new, exact = gen_triplet(seeds[key], p["value"], p["dp"])
            for ref, x in zip(cells, new):
                j, ok = set_num(j, ref, fmt(x))
                if ok:
                    jc[ref] = float(x)
                else:
                    log["missing_cells"].append(ref)
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
    out[jpath] = j.encode("utf8")

    # --- расчётные листы: объект и дата
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
        out[p] = x.encode("utf8")

    wbx = zin.read("xl/workbook.xml").decode("utf8")
    if "<calcPr" in wbx:
        wbx = re.sub(r"<calcPr([^>]*?)/>", lambda m: "<calcPr" + re.sub(r' fullCalcOnLoad="[^"]*"', "", m.group(1))
                     + ' fullCalcOnLoad="1"/>', wbx, count=1)
    else:
        wbx = wbx.replace("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>')
    out["xl/workbook.xml"] = wbx.encode("utf8")

    # --- расчёт формул и запись готовых значений в файл
    try:
        allc = dict(tcells)
        allc[JOURNAL_SHEET] = jc
        book = Book(allc)
        for n, p in paths:
            out[p] = inject_values(out[p].decode("utf8"), n, book).encode("utf8")
        # контроль: среднее в журнале == значение протокола (с учётом его точности)
        for label, val, trip, exact, ok, key, dp in log["filled"]:
            _, row, col = blocks[key]
            r = book.result(JOURNAL_SHEET, col + str(row + 3))
            if r[0] != "v" or abs(D(repr(r[1])) - val) > D("1e-9"):
                log["verify_bad"].append(label)
    except Exception:  # noqa: BLE001
        log["calc"] = False
        for n, p in paths:
            out[p] = strip_cache(out[p].decode("utf8")).encode("utf8")
    return pack(zin, out), log


def process(pdf_bytes, template_bytes):
    raw = pdf_text(pdf_bytes)
    if "Aniqlanayotgan parametr nomi" not in raw:
        raise ValueError("В PDF не найден текст протокола (возможно, это скан без текстового слоя).")
    meta = parse_meta(raw)
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
            st.write("Используется %s рядом с app.py." % TEMPLATE_NAME)
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
            st.error("Не удалось обработать файл: %s" % e)

    if template_bytes is None:
        st.info("Положите template.xlsx рядом с app.py или загрузите шаблон выше.")
    elif pdf is None:
        st.info("Загрузите PDF-протокол.")

    res = st.session_state.get("res")
    if not res:
        return
    data, meta, log, pdf_name = res

    st.success("Журнал готов")
    name = "Журнал_%s_%s_%s.xlsx" % (safe_name(meta["brand"]), safe_name(meta["model"]), meta["vin"])
    st.download_button(
        "⬇️ Скачать журнал (.xlsx)", data=data, file_name=name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    d = meta["date"]
    st.subheader("Реквизиты")
    st.markdown(
        "- **Протокол:** %s\n- **Марка / модель:** %s / %s\n- **VIN:** %s\n- **Дата:** %s\n"
        "- **Условия:** %s °C, %s %%, %s кПа, %s м/с" % (
            meta["protocol"] or "—", meta["brand"], meta["model"], meta["vin"] or "—",
            d.strftime("%d.%m.%Y") if d else "—", meta["temp"], meta["hum"], meta["press"], meta["wind"]))

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
        st.info("Параметры протокола без места в журнале (не перенесены):\n" +
                "\n".join("- %s" % x[:120] for x in log["unmatched"]))

    with st.expander("Все параметры протокола и куда они попали (%d)" % len(log.get("trace", []))):
        for sec, nm, key, fin in log.get("trace", []):
            st.write("%s `%s` → %s  (%s)" % ("✅" if key else "⛔", key or "нет места в журнале", nm[:90], fin))
    with st.expander("Очищено — нет в протоколе (%d)" % len(log["cleared"])):
        st.write("\n".join("- %s" % x for x in log["cleared"]) or "—")
    with st.expander("Заполненные блоки (%d)" % len(log["filled"])):
        for label, val, trip, exact, ok, key, dp in log["filled"]:
            st.write("**%s**: %s → %s%s%s" % (label, fmt(val), " / ".join(trip),
                                              "" if exact else " ≈", "" if ok else " ⚠️N"))
    st.caption("≈ — у этих блоков шаг разброса из шаблона слегка изменён (на 1–2 единицы последнего "
               "знака), чтобы среднее было точно равно значению протокола.")


if __name__ == "__main__":
    main()
