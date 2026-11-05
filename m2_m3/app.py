# -*- coding: utf-8 -*-
"""
Заполнение журнала сертификационных испытаний ТС категорий M2–M3
(автобусы, микроавтобусы) по PDF-протоколу испытаний (шаблон: template.xlsx).

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
from decimal import Decimal as D, InvalidOperation
from fractions import Fraction as Fr
from pathlib import Path
from xml.sax.saxutils import escape

import openpyxl
from openpyxl.utils.cell import column_index_from_string, coordinate_from_string, get_column_letter
import streamlit as st
from pypdf import PdfReader

warnings.filterwarnings("ignore")

JOURNAL_SHEET = "НЕОПРЕДЕЛ"
JCOL = "E"                     # колонка замеров на листе журнала
TEMPLATE_NAME = "template.xlsx"
PROTO_REF = ("48-04", "F1")    # «Приложение к протоколу №…» (на него ссылаются все расчётные листы)

# --------------------------------------------------------------------------
# Раскладка журнала (лист «НЕОПРЕДЕЛ»): ключ -> (подпись, первая строка тройки)
# Тройка: E[r], E[r+1], E[r+2]; среднее «СР.ЗАНЧ» — E[r+3] (считается на расчётном листе).
# --------------------------------------------------------------------------
SIMPLE = {
    "heat": ("Температура частей системы отопления (Правило 122)", 219),
    "length": ("Максимальная длина", 224),
    "width": ("Максимальная ширина", 229),
    "height": ("Максимальная высота", 234),
    "horn": ("Звуковой сигнал", 239),
    "glass_w": ("Ветровые стёкла", 245),
    "glass_s": ("Стёкла, не являющиеся ветровыми", 250),
    "vent_speed": ("Скорость воздуха на выходе вентиляции", 255),
    "vent_mob": ("Подвижность воздуха в зоне головы", 260),
    "ac_cool": ("Охлаждение воздуха кондиционером", 266),
    "ac_speed": ("Скорость воздуха на выходе кондиционера", 271),
    "ac_temp": ("Температура воздуха", 276),
    "ac_head": ("Скорость воздуха в зоне головы", 281),
    "humidity": ("Относительная влажность", 286),
    "duct_temp": ("Температура поверхностей воздуховодов", 292),
    "plate_a1": ("Номерной знак: угол (б)", 297),
    "plate_a2": ("Номерной знак: угол (в)", 302),
    "plate_hmax": ("Номерной знак: высота (max)", 307),
    "plate_hmin": ("Номерной знак: высота (min)", 312),
    "pict": ("Надпись/пиктограмма у передней двери", 317),
    "digits": ("Цифры высотой", 322),
    "first_aid": ("Место для аптечки", 327),
    "srv_door_h": ("Служебная дверь: высота", 332),
    "srv_door_w": ("Служебная дверь: ширина", 337),
    "em_door_h": ("Аварийная дверь: высота", 342),
    "em_door_w": ("Аварийная дверь: ширина", 347),
    "em_win1_w": ("Аварийное окно 1: ширина", 352),
    "em_win1_l": ("Аварийное окно 1: длина", 356),
    "em_win2_w": ("Аварийное окно 2: ширина", 361),
    "em_win2_l": ("Аварийное окно 2: длина", 365),
    "door_ctrl": ("Органы открывания дверей снаружи: высота", 370),
    "em_win_floor": ("Нижний край аварийного окна над полом", 375),
    "step1_maxh": ("Первая ступенька: макс. высота", 381),
    "step1_mind": ("Первая ступенька: мин. глубина", 386),
    "step_maxh": ("Другие ступеньки: макс. высота", 391),
    "step_minh": ("Другие ступеньки: мин. высота", 396),
    "step_mind": ("Другие ступеньки: мин. глубина", 401),
    "seat_f": ("Ширина сиденья F", 406),
    "seat_g": ("Ширина свободного пространства G", 411),
    "cushion_depth": ("Глубина подушки сиденья", 416),
    "seat_space_h": ("Пространство над сиденьем: высота", 421),
    "handhold": ("Поручни у служебных дверей", 426),
    "guard": ("Ограждение проёмов для ступенек", 431),
}
# Световые приборы (лист «48-04»): огонь -> (подпись, {позиция: строка})
LAMPS = {
    "dipped": ("Фары ближнего света", {"top": 19, "bot": 24, "betw": 29, "width": 34}),
    "reverse": ("Огонь заднего хода", {"top": 39, "bot": 44}),
    "pos_front": ("Габаритные огни передние", {"top": 49, "bot": 54, "betw": 59, "width": 64}),
    "pos_rear": ("Габаритные огни задние", {"top": 69, "bot": 74, "betw": 79, "width": 84}),
    "brake": ("Сигнал торможения", {"top": 89, "bot": 94, "betw": 99, "width": 104}),
    "turn_front": ("Указатели поворота передние", {"top": 109, "bot": 114, "betw": 119, "width": 124}),
    "turn_rear": ("Указатели поворота задние", {"top": 129, "bot": 134, "betw": 139, "width": 144}),
    "side": ("Боковые повторители", {"top": 149, "bot": 154, "len": 159}),
    "rear_fog": ("Задний противотуманный фонарь", {"top": 164, "bot": 169, "brake": 174}),
    "refl_rear": ("Световозвращатели задние", {"top": 179, "bot": 184, "betw": 189, "width": 194}),
    "refl_front": ("Световозвращатели передние", {"top": 199, "bot": 204, "betw": 209, "width": 214}),
    # в журнале M2–M3 для них места нет — распознаём, чтобы не путать со следующими пунктами
    "fog_front": ("Передняя противотуманная фара", {}),
    "side_pos": ("Боковые габаритные огни", {}),
    "refl_side": ("Световозвращатели боковые", {}),
}
POS_LABEL = {"top": "верхняя точка", "bot": "нижняя точка", "betw": "между внутренними точками",
             "width": "до габарита по ширине", "len": "от передней части АТС",
             "brake": "до сигнала торможения"}
DIM_KEYS = ("length", "width", "height")   # в протоколе — m, в журнале — mm

MONTHS = ["января", "февраля", "марта", "апреля", "мая", "июня", "июля",
          "августа", "сентября", "октября", "ноября", "декабря"]


def all_blocks():
    """Все блоки замеров журнала: ключ -> (подпись, первая строка)."""
    res = dict(SIMPLE)
    for k, (label, rows) in LAMPS.items():
        for pos, row in rows.items():
            res["%s|%s" % (k, pos)] = ("%s: %s" % (label, POS_LABEL[pos]), row)
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
            # «markasi YUTONG, model ZK6890HQG,»
            r"markasi\s+([^,]+?)\s*,?\s+model[i]?\s+([^,\s]+(?:\s[^,\s]+)?)\s*(?:,|VIN)",
            # «FOTON markali BJ6931 modelli»
            r"([A-Za-zА-Яа-я0-9\-]+)\s+markali\s+(.+?)\s+modelli",
            # «марки MAN шасси MAN TGM-13.290 (»
            r"марк[аиеу]?\s+([A-Za-zА-Яа-я0-9\-]+)\s+шасси\s+(?:\1\s+)?([^\s(,]+)",
        ]
        for pat in pats:
            m = re.search(pat, flat, re.I)
            if m:
                meta["brand"] = meta["brand"] or m.group(1).strip()
                meta["model"] = meta["model"] or m.group(2).strip()
                break
    m = (re.search(r"\b[MМ]([123])\s*toifa", flat)
         or re.search(r"категори[яи][^)]{0,40}?\b[MМ]([123])\b", flat))
    meta["category"] = "M" + m.group(1) if m else ""
    vins = []
    for mm in re.finditer(r"(?:VIN|идент\.?|идентификационный)[^A-Z0-9]{0,25}((?:[A-HJ-NPR-Z0-9]{17}[\s,;]*)+)", flat):
        for v in re.findall(r"[A-HJ-NPR-Z0-9]{17}", mm.group(1)):
            if v not in vins:
                vins.append(v)
    if not vins:
        vins = list(dict.fromkeys(re.findall(r"\b([A-HJ-NPR-Z0-9]{17})\b", flat)))
    meta["vins"] = vins
    meta["vin"] = vins[0] if vins else ""
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
# Сопоставление параметров протокола с блоками журнала M2–M3
# --------------------------------------------------------------------------
def detect_lamp(n):
    if "между сигналом торможения" in n:
        return "rear_fog"
    if "от передней части атс" in n:
        return "side"
    if "ближнего" in n and "фар" in n:
        return "dipped"
    if "заднего хода" in n:
        return "reverse"
    if "противотуман" in n:
        return "rear_fog" if ("задн" in n or "фонар" in n) else "fog_front"
    if "габаритн" in n and "огн" in n:
        if "боков" in n:
            return "side_pos"
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
        if "боков" in n:
            return "refl_side"
        return "refl_front" if "передн" in n else ("refl_rear" if "задн" in n else None)
    return None


def lamp_pos(n):
    if "между сигналом торможения" in n:
        return "brake"
    if "от передней части атс" in n:
        return "len"
    if "верхней точки" in n:
        return "top"
    if "нижней точки" in n:
        return "bot"
    if "между крайними внутренними" in n:
        return "betw"
    if "габарита кузова по ширине" in n:
        return "width"
    return None


LAMP_ORDER = ["dipped", "fog_front", "reverse", "pos_front", "pos_rear", "brake", "turn_front",
              "turn_rear", "side", "rear_fog", "refl_rear", "refl_front", "side_pos", "refl_side"]


def is_short(n):
    """Короткое продолжение предыдущего пункта («высота», «ширина», «минимальная глубина»…)."""
    return len(n) <= 40 and not re.search(r"сиден|подушк|поручн|огнетуш|аптеч|окн|ступен", n)


def classify(p, ctx):
    """Возвращает ключ блока (или None). ctx — контекст предыдущих пунктов
    (текущий огонь, дверь, группа ступенек), меняется по ходу разбора."""
    section, n, req = p["section"], p["name_f"], p["req"]

    # ---------------- световые приборы
    in_light = ("освещен" in section or "световой сигнал" in section
                or ("светов" in section and "прибор" in section))
    lamp_here = detect_lamp(n)
    if lamp_here or in_light:
        lamp = lamp_here or (ctx.get("lamp") if in_light else None)
        pos = lamp_pos(n)
        seen = ctx.setdefault("seen", set())
        if not lamp_here and lamp and pos and (lamp, pos) in seen:
            # в протоколе пропущено название огня: позиция повторилась — берём следующий огонь
            # по обычному порядку пунктов (с предупреждением)
            used = {l for l, _ in seen}
            nxt = [x for x in LAMP_ORDER[LAMP_ORDER.index(lamp) + 1:] if x not in used]
            lamp = nxt[0] if nxt else None
            p["_guess"] = lamp
        ctx["lamp"] = lamp
        p["_lamp"] = lamp
        if lamp and pos:
            seen.add((lamp, pos))
        if lamp and pos and pos in LAMPS[lamp][1]:
            return "%s|%s" % (lamp, pos)
        return None
    ctx["lamp"] = None

    # ---------------- контекст дверей / ступенек / огнетушителя
    short = is_short(n)
    if "двер" in n and "служебн" in n:
        ctx["door"] = "srv"
    elif "двер" in n and "аварийн" in n:
        ctx["door"] = "em"
    elif not short:
        ctx["door"] = None
    if "ступен" in n:
        ctx["step"] = "first" if "перв" in n else ("other" if "друг" in n else ctx.get("step"))
        if "перв" not in n and "друг" not in n and ("ступенек" in n or "ступеньки" in n):
            ctx["step"] = "any"
    elif not short:
        ctx["step"] = None
    if "огнетуш" in n:
        ctx["ext"] = True
    elif not short or "двер" in n:
        ctx["ext"] = False

    # ---------------- отопление, габариты, сигнал, стёкла
    if "температура поверхности без покрытия" in n or "температура любой такой части" in n:
        return "heat"
    if "ступен" not in n and "ступен" not in section:
        if "максимальная длина" in n:
            return "length"
        if "максимальная ширина" in n and "транспортного средства" in n:
            return "width"
        if "максимальная высота" in n and "транспортного средства" in n:
            return "height"
    if "звуков" in n and ("сигнал" in n or "уровень" in n or "звука" in n):
        return "horn"
    if "светопропускан" in section or "светопропускан" in n or "стекл" in n:
        if "не являющ" in n or "боков" in n:
            return "glass_s"
        if "ветров" in n:
            return "glass_w"

    # ---------------- вентиляция / кондиционирование (по названию пункта)
    if "потоков на выходе из системы вентиляции" in n:
        return "vent_speed"
    if "подвижность воздуха" in n or "система вентиляции должна" in n:
        return "vent_mob"
    if "исключать возможность охлаждения" in n or ("охлаждения воздуха" in n and "голов" in n):
        return "ac_cool"
    if "потока на выходе из системы кондиционирования" in n:
        return "ac_speed"
    if n.startswith("температура воздуха должна"):
        return "ac_temp"
    if "скорость воздуха в зоне головы" in n:
        return "ac_head"
    if "относительная влажность" in n:
        return "humidity"
    if "температура наружных поверхностей воздуховодов" in n:
        return "duct_temp"

    # ---------------- номерной знак
    if "номерн" in section or "номерн" in n:
        if "продольн" in n:
            return "plate_a1"
        if "опорн" in n and "высот" not in n:
            return "plate_a2"
        if "верхнего края" in n:
            return "plate_hmax"
        if "нижнего края" in n:
            return "plate_hmin"
        lo, hi = req
        if lo is not None and hi is None:
            return "plate_hmin"
        if hi is not None and lo is None:
            return "plate_hmax"
        return None

    # ---------------- автобусы: надписи, аптечка, двери, окна
    if "пиктограмм" in n or ("надпись" in n and "высот" in n):
        return "pict"
    if "цифры высотой" in n:
        return "digits"
    if "аптеч" in n:
        return "first_aid"
    if "орган управления" in n and "открывания дверей" in n:
        return "door_ctrl"
    if "аварийного окна" in n and "над уровнем пола" in n:
        return "em_win_floor"
    if "аварийн" in n and "окн" in n:
        if "ширин" in n:
            return "em_win_w"
        if "длин" in n or "дилин" in n or "высот" in n:
            return "em_win_l"
    if "ручк" in n and "двер" in n:
        return None                     # высота наружной ручки — места в журнале нет
    if ctx.get("door") and not ctx.get("ext"):
        if "высот" in n:
            return ctx["door"] + "_door_h"
        if "ширин" in n:
            return ctx["door"] + "_door_w"

    # ---------------- сиденья
    if "сплошное сиденье" in n or "ширина сиденья" in n:
        return "seat_f"
    if "ширина свободного пространства" in n:
        return "seat_g"
    if "глубина подушки" in n:
        return "cushion_depth"
    if "над каждым местом для сиденья" in n:
        return "seat_space_h"

    # ---------------- поручни, ограждение
    if "рукоятка располагается по вертикали" in n or ("поручн" in n and "служебн" in n and "двер" in n):
        return "handhold"
    if "огражден" in n:
        return "guard"

    # ---------------- ступеньки
    step = ctx.get("step")
    if step and ("ступен" in n or short) and "наклон" not in n:
        if "максимальн" in n and "высот" in n:
            return "step_maxh" if step == "other" else "step1_maxh"
        if "минимальн" in n and "высот" in n:
            return "step_minh"
        if "минимальн" in n and "глубин" in n:
            return "step_mind" if step == "other" else "step1_mind"
    return None


OVERFLOW = {"em_win_w": ("em_win1_w", "em_win2_w"), "em_win_l": ("em_win1_l", "em_win2_l")}


def map_params(params):
    """-> values {ключ: {...}}, не найденные, дубли, трасса сопоставления."""
    values, unmatched, dup, trace = {}, [], [], []
    ctx = {}
    for p in params:
        key = classify(p, ctx)
        if key in OVERFLOW:
            free = [k for k in OVERFLOW[key] if k not in values]
            key = free[0] if free else OVERFLOW[key][-1]
        if key is None:
            unmatched.append(p)
            trace.append((p, None))
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
        unit = p["rows"][0]["unit"].lower()
        if key in DIM_KEYS and (unit in ("m", "м") or v < 100):
            v, dp = v * 1000, max(dp - 3, 0)          # m -> mm (в журнале габариты в mm)
            v = v.quantize(D(1).scaleb(-dp))
            notes = notes + ["переведено из m в mm: %s" % fmt(v)]
        if key == "first_aid" and unit not in ("mm", "мм"):
            notes = notes + ["в протоколе единица «%s», в журнале — mm" % p["rows"][0]["unit"]]
        if p.get("_guess"):
            notes = notes + ["в протоколе нет названия огня — по порядку пунктов отнесено к «%s»"
                             % LAMPS[p["_guess"]][0]]
        ok = all(r["concl"] == "M" for r in p["rows"])
        values[key] = {"value": v, "dp": dp, "ok": ok, "name": p["name"], "notes": notes}
        trace.append((p, key))
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
        v = self.expr()
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


def drop_calc_chain(zin, out):
    """calcChain.xml удаляется (Excel построит его сам): часть «формул-констант»
    в тройках заменяется значениями, и старая цепочка стала бы некорректной."""
    names = set(zin.namelist())
    if "xl/calcChain.xml" not in names:
        return
    out["xl/calcChain.xml"] = None
    ct = zin.read("[Content_Types].xml").decode("utf8")
    out["[Content_Types].xml"] = re.sub(r'<Override PartName="/xl/calcChain\.xml"[^>]*/>', "", ct).encode("utf8")
    rels = zin.read("xl/_rels/workbook.xml.rels").decode("utf8")
    out["xl/_rels/workbook.xml.rels"] = re.sub(
        r'<Relationship\b[^>]*Target="[^"]*calcChain\.xml"[^>]*/>', "", rels).encode("utf8")


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


def set_input(xml, ref, val):
    """Записать число во входную ячейку тройки (формула-константа «=619» тоже заменяется)."""
    m = cell_re(ref).search(xml)
    if not m:
        return xml, False
    attrs = T_ATTR_RE.sub("", m.group(1))
    return xml[:m.start()] + '<c r="%s"%s><v>%s</v></c>' % (ref, attrs, val) + xml[m.end():], True


def clear_input(xml, ref):
    m = cell_re(ref).search(xml)
    if not m:
        return xml, False
    attrs = T_ATTR_RE.sub("", m.group(1))
    return xml[:m.start()] + '<c r="%s"%s/>' % (ref, attrs) + xml[m.end():], True


def block_bias(book, row):
    """Поправка «Og'ish» блока: среднее тройки − «СР.ЗАНЧ» (E[r+3]); None — нельзя определить."""
    try:
        vals = [book.value(JOURNAL_SHEET, JCOL + str(row + k)) for k in range(3)]
        if not all(isinstance(v, (int, float)) and not isinstance(v, bool) for v in vals):
            return None
        r = book.result(JOURNAL_SHEET, JCOL + str(row + 3))
        if r[0] != "v" or not isinstance(r[1], (int, float)):
            return None
        return D(repr(sum(vals) / 3.0 - r[1])).quantize(D("1e-9"))
    except Exception:  # noqa: BLE001
        return None


def num_or_text(v):
    return fmt(D(v)) if v is not None else ""


def fill_workbook(template_bytes, meta, values, mode="meas"):
    """mode: "meas" — среднее тройки = протоколу (как в шаблоне);
             "corr" — «СР.ЗАНЧ» (с поправкой Og'ish) = протоколу."""
    tcells = load_cells(template_bytes)
    if JOURNAL_SHEET not in tcells:
        raise ValueError("В шаблоне нет листа «%s» — это не шаблон M2–M3." % JOURNAL_SHEET)
    tbook = Book(tcells)
    blocks = all_blocks()
    seeds, bias = {}, {}
    for key, (label, row) in blocks.items():
        vals = []
        for k in range(3):
            try:
                v = tbook.value(JOURNAL_SHEET, JCOL + str(row + k))
            except XlErr:
                v = None
            vals.append(D(repr(v)) if isinstance(v, (int, float)) and not isinstance(v, bool) else None)
        seeds[key] = vals
        bias[key] = block_bias(tbook, row)

    log = {"filled": [], "cleared": [], "inexact": [], "missing_cells": [], "verify_bad": [],
           "calc": True, "unlinked": []}
    zin = zipfile.ZipFile(io.BytesIO(template_bytes))
    paths = sheet_paths(zin)
    jpath = next(p for n, p in paths if n == JOURNAL_SHEET)
    j = zin.read(jpath).decode("utf8")
    out = {}
    jc = dict(tcells[JOURNAL_SHEET])

    # --- шапка
    head = [("D5", "%s ˚С" % num_or_text(meta["temp"]) if meta["temp"] is not None else ""),
            ("D6", "%s %%" % num_or_text(meta["hum"]) if meta["hum"] is not None else ""),
            ("D7", "%s kPa" % num_or_text(meta["press"]) if meta["press"] is not None else ""),
            ("D8", "%s m/s" % num_or_text(meta["wind"]) if meta["wind"] is not None else ""),
            ("D10", meta["brand"]), ("D11", meta["model"]), ("D12", meta["category"]),
            ("D14", meta["vin"])]
    for ref, v in head:
        if v and v.strip():
            j, ok = set_str(j, ref, v)
            jc[ref] = v
            if not ok:
                log["missing_cells"].append(ref)

    # --- замеры
    for key, (label, row) in blocks.items():
        cells = [JCOL + str(row + k) for k in range(3)]
        if key in values:
            p = values[key]
            b = bias[key] if bias[key] is not None else D(0)
            dp = p["dp"]
            if mode == "corr" and b:
                target = p["value"] + b
                dp = max(dp, -b.normalize().as_tuple().exponent)
            else:
                target = p["value"]
            new, exact = gen_triplet(seeds[key], target, dp)
            for ref, x in zip(cells, new):
                j, ok = set_input(j, ref, fmt(x))
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
                j, done = clear_input(j, ref)
                if done:
                    jc.pop(ref, None)
            if had:
                log["cleared"].append(label)
    out[jpath] = j.encode("utf8")

    # --- номер протокола на расчётном листе «48-04» (остальные листы ссылаются на него)
    pr_sheet, pr_ref = PROTO_REF
    pr_path = next((p for n, p in paths if n == pr_sheet), None)
    pr_cells = {}
    if pr_path and meta["protocol"]:
        x = zin.read(pr_path).decode("utf8")
        txt = "Приложение к протоколу №%s" % meta["protocol"]
        x, ok = set_str(x, pr_ref, txt)
        if ok:
            out[pr_path] = x.encode("utf8")
            pr_cells[pr_ref] = txt
    for n, p in paths:
        if p not in out:
            out[p] = zin.read(p)

    wbx = zin.read("xl/workbook.xml").decode("utf8")
    if "<calcPr" in wbx:
        wbx = re.sub(r"<calcPr([^>]*?)/>", lambda m: "<calcPr" + re.sub(r' fullCalcOnLoad="[^"]*"', "", m.group(1))
                     + ' fullCalcOnLoad="1"/>', wbx, count=1)
    else:
        wbx = wbx.replace("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>')
    out["xl/workbook.xml"] = wbx.encode("utf8")
    drop_calc_chain(zin, out)

    # --- расчёт формул и запись готовых значений
    try:
        allc = dict(tcells)
        allc[JOURNAL_SHEET] = jc
        if pr_cells:
            allc[pr_sheet] = dict(allc[pr_sheet], **pr_cells)
        book = Book(allc)
        for n, p in paths:
            out[p] = inject_values(out[p].decode("utf8"), n, book).encode("utf8")
        # контроль: среднее тройки и «СР.ЗАНЧ» на расчётном листе
        for label, val, trip, exact, ok, key, dp in log["filled"]:
            row = blocks[key][1]
            avg = sum(Fr(D(t)) for t in trip) / 3
            b = bias[key]
            want_avg = Fr(val + (b or 0)) if mode == "corr" else Fr(val)
            if avg != want_avg:
                log["verify_bad"].append(label)
                continue
            r = book.result(JOURNAL_SHEET, JCOL + str(row + 3))
            if b is None or r[0] != "v" or not isinstance(r[1], (int, float)):
                log["unlinked"].append(label)
                continue
            if abs(D(repr(r[1])) - (D(avg.numerator) / D(avg.denominator) - b)) > D("1e-6"):
                log["unlinked"].append(label)
    except Exception:  # noqa: BLE001
        log["calc"] = False
        for n, p in paths:
            out[p] = strip_cache(out[p].decode("utf8")).encode("utf8")
    return pack_out(zin, out), log


def process(pdf_bytes, template_bytes, mode="meas"):
    raw = pdf_text(pdf_bytes)
    if "Aniqlanayotgan parametr nomi" not in raw:
        raise ValueError("В PDF не найден текст протокола (возможно, это скан без текстового слоя).")
    meta = parse_meta(raw)
    params = parse_params(raw)
    if not params:
        raise ValueError("В протоколе не удалось найти таблицы результатов.")
    values, unmatched, dup, trace = map_params(params)
    data, log = fill_workbook(template_bytes, meta, values, mode)
    log["unmatched"] = [p["name"] for p in unmatched]
    log["trace"] = [(p["section"], p["name"], key, p["rows"][0]["final"]) for p, key in trace]
    log["suspect"] = [p["name"] for p in unmatched
                      if "освещающей поверхности" in p["name_f"] and not p.get("_lamp")]
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
MODES = {"По замерам: среднее тройки = протоколу (как в шаблоне)": "meas",
         "По «СР.ЗАНЧ»: среднее с поправкой Og'ish = протоколу": "corr"}


def main():
    st.set_page_config(page_title="Журнал M2–M3", page_icon="🚌", layout="centered")
    st.markdown(
        "<style>.stButton>button,.stDownloadButton>button{min-height:3rem;font-size:1.05rem}"
        ".block-container{padding-top:1.2rem;padding-bottom:2rem}</style>",
        unsafe_allow_html=True,
    )
    st.title("🚌 Журнал испытаний M2–M3")
    st.caption("Автобусы и микроавтобусы. Загрузите PDF-протокол — получите заполненный журнал по шаблону.")

    tpl_path = Path(__file__).resolve().parent / TEMPLATE_NAME
    template_bytes = tpl_path.read_bytes() if tpl_path.exists() else None
    if template_bytes is None:
        up_t = st.file_uploader("Шаблон журнала M2–M3 (template.xlsx)", type=["xlsx"], key="tpl")
        if up_t is not None:
            template_bytes = up_t.getvalue()
    else:
        with st.expander("Шаблон и настройки"):
            st.write("Используется %s рядом с app.py." % TEMPLATE_NAME)
            up_t = st.file_uploader("Заменить шаблон", type=["xlsx"], key="tpl2")
            if up_t is not None:
                template_bytes = up_t.getvalue()
            st.radio("С протоколом совпадает:", list(MODES), key="mode")

    pdf = st.file_uploader("PDF-протокол испытаний", type=["pdf"], key="pdf")

    if st.button("Заполнить журнал", type="primary", use_container_width=True,
                 disabled=(pdf is None or template_bytes is None)):
        try:
            mode = MODES.get(st.session_state.get("mode"), "meas")
            with st.spinner("Обработка протокола…"):
                data, meta, log = process(pdf.getvalue(), template_bytes, mode)
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
    name = "Журнал_M2M3_%s_%s_%s.xlsx" % (safe_name(meta["brand"]), safe_name(meta["model"]), meta["vin"])
    st.download_button(
        "⬇️ Скачать журнал (.xlsx)", data=data, file_name=name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    d = meta["date"]
    st.subheader("Реквизиты")
    st.markdown(
        "- **Протокол:** %s\n- **Марка / модель:** %s / %s\n- **Категория:** %s\n- **VIN:** %s\n"
        "- **Дата испытаний:** %s\n- **Условия:** %s °C, %s %%, %s кПа, %s м/с" % (
            meta["protocol"] or "—", meta["brand"] or "—", meta["model"] or "—", meta["category"] or "—",
            meta["vin"] or "—", d.strftime("%d.%m.%Y") if d else "—",
            meta["temp"], meta["hum"], meta["press"], meta["wind"]))
    if len(meta.get("vins", [])) > 1:
        st.warning("В протоколе несколько VIN, в журнал записан первый. Все: " + ", ".join(meta["vins"]))
    if not meta["brand"] or not meta["model"]:
        st.warning("Марка/модель не распознаны — впишите вручную (НЕОПРЕДЕЛ!D10, D11).")

    c1, c2, c3 = st.columns(3)
    c1.metric("Заполнено", len(log["filled"]))
    c2.metric("Очищено", len(log["cleared"]))
    c3.metric("Шаг изменён", len(log["inexact"]))

    if log["verify_bad"]:
        st.error("Среднее тройки не сошлось с протоколом: " + "; ".join(log["verify_bad"]))
    elif log["calc"]:
        st.success("Проверка: среднее каждого заполненного блока совпадает с протоколом.")
    else:
        st.warning("Значения формул не посчитаны: они пересчитаются при открытии файла в Excel.")
    if log.get("unlinked"):
        st.warning("«СР.ЗАНЧ» этих блоков в шаблоне не связан с тройкой (проверьте формулы расчётного листа): "
                   + "; ".join(log["unlinked"]))
    if log["notes"]:
        st.warning("Проверьте (расхождения в протоколе, пересчёт единиц):\n" +
                   "\n".join("- %s: %s" % (n[:70], t) for n, t in log["notes"]))
    if log["nonconf"]:
        st.error("Несоответствия (N) в протоколе: " + "; ".join(log["nonconf"]))
    if log["missing_cells"]:
        st.warning("Не найдены ячейки в шаблоне: " + ", ".join(sorted(set(log["missing_cells"]))))
    if log["dup"]:
        st.warning("Повторяющиеся параметры в протоколе (взят первый): " + ", ".join(log["dup"]))
    if log.get("suspect"):
        st.error("Похоже на световые приборы, но огонь не распознан — проверьте вручную:\n" +
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
    st.caption("≈ — у этих блоков шаг разброса из шаблона слегка изменён (на 1–2 единицы последнего "
               "знака), чтобы среднее было точно равно значению протокола.")


if __name__ == "__main__":
    main()
