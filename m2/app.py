# -*- coding: utf-8 -*-
"""
Заполнение журнала сертификационных испытаний ТС категории M2 (автобусы малой вместимости)
по PDF-протоколу испытаний (шаблон: template_m2.xlsx).
Код совпадает с m2_m3/app.py (M3) — отличаются только CAT и TEMPLATE_NAME.

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

JOURNAL_SHEET = "НЕОПРЕДЕЛ"
JCOL = "E"                     # колонка замеров на листе журнала
CAT = "M2"                     # категория ТС этого приложения
TEMPLATE_NAME = "template_%s.xlsx" % CAT.lower()
CAT_NOTE = ""   # уточнение категории для подписи на сайте
PROTO_REF = ("48-04", "F1")
ID_REF = "E1"                  # «ID №…» на листе журнала = название PDF-файла    # «Приложение к протоколу №…» (на него ссылаются все расчётные листы)

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
    "seat_g": ("Ширина подушки сиденья G", 411),
    "cushion_depth": ("Глубина подушки сиденья", 416),
    "seat_space_h": ("Пространство над сиденьем: высота", 421),
    "handhold": ("Поручни у служебных дверей", 426),
    "guard": ("Ограждение проёмов для ступенек", 431),
    # --- дополнительные блоки шаблона (с строки 493)
    "mirror": ("Зеркала класса 5 и 6: высота от дороги", 493),
    "mass_total": ("Технически допустимая максимальная масса", 498),
    "mass_axle_driven": ("Масса на одиночную ведомую ось", 503),
    "mass_axle_drive": ("Масса на одиночную ведущую ось", 508),
    "g_entry": ("Угол въезда", 513),
    "g_exit": ("Угол съезда", 518),
    "g_ramp": ("Угол продольной проходимости", 523),
    "g_clear_f": ("Дорожный просвет под передней осью", 528),
    "g_clear_r": ("Дорожный просвет под задней осью", 533),
    "g_clear_mid": ("Межосевой дорожный просвет", 538),
    "seat_free": ("Ширина свободного пространства места сиденья", 543),
    "handrail_len": ("Длина поручня", 548),
    "handrail_gap": ("Зазор между поручнем и потолком/стенкой", 553),
    "handrail_h": ("Горизонтальные поручни: высота от пола", 558),
    "ext_len": ("Место для огнетушителя: длина", 563),
    "ext_h": ("Место для огнетушителя: высота", 568),
    "ext_w": ("Место для огнетушителя: ширина", 573),
    "em_handle": ("Высота наружной ручки аварийной двери", 578),
    "first_aid_vol": ("Объём места для аптечки", 583),
    "aisle_long": ("Наклон прохода: продольный", 588),
    "aisle_trans": ("Наклон прохода: поперечный", 593),
    "step_slope": ("Наклон ступеньки", 598),
    "cushion_h": ("Высота подушки сиденья", 603),
    "seat_dist": ("Расстояние между сиденьями", 608),
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
    "fog_front": ("Передняя противотуманная фара", {"top": 438, "bot": 443, "width": 448}),
    "side_pos": ("Боковые габаритные огни", {"top": 453, "bot": 458, "flen": 463, "rlen": 468}),
    "refl_side": ("Световозвращатели боковые", {"top": 473, "bot": 478, "flen": 483, "rlen": 488}),
}
POS_LABEL = {"top": "верхняя точка", "bot": "нижняя точка", "betw": "между внутренними точками",
             "width": "до габарита по ширине", "len": "от передней части АТС",
             "brake": "до сигнала торможения", "flen": "от передней оконечности ТС",
             "rlen": "до задней оконечности ТС"}
DIM_KEYS = ("length", "width", "height")   # в журнале — m (как в протоколе)

MONTHS = ["января", "февраля", "марта", "апреля", "мая", "июня", "июля",
          "августа", "сентября", "октября", "ноября", "декабря"]


LAYOUT = {}   # раскладка, найденная по подписям текущего шаблона (build_layout)


def all_blocks():
    """Все блоки замеров журнала: ключ -> (подпись, первая строка).
    Если шаблон уже прочитан — раскладка по его подписям, иначе — раскладка по умолчанию."""
    if LAYOUT:
        return dict(LAYOUT)
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


def clean_brand(b):
    """Сократить марку: без пробега/запаса хода («530 км»), года («2026») и «Model …» после них."""
    b = re.split(r"\s*(?:\b\d+[.,]?\d*\s*(?:км|km)\b|\b(?:19|20)\d{2}\b|\b(?:model|модел\w*)\b)", b,
                 maxsplit=1, flags=re.I)[0]
    return b.strip(" ,.;-") or b


def parse_meta(raw):
    flat = norm(raw)
    meta = {}
    m = re.search(r"№\s*(SB-[\w\-/]+)", flat)
    meta["protocol"] = m.group(1) if m else ""
    m = re.search(r"марк[аиеу]?\s*[«\"“]([^»\"”]+)[»\"”]", flat)
    meta["brand"] = m.group(1).strip() if m else ""
    m = re.search(r"модел[ьиюе]\s*[«\"“]([^»\"”]+)[»\"”]", flat)
    meta["model"] = m.group(1).strip() if m else ""
    # «Мобильная химико-бактериологическая лаборатория (марка транспортного средства KRANTAS (категория …»
    m = re.search(r"Sinov obyekti:\s*(.{3,120}?)\s*\(\s*марк[аи] транспортного средства\s+([A-Za-zА-Яа-я0-9\-]+)", flat)
    if m and not meta["brand"]:
        meta["brand"] = m.group(2).strip()
        meta["model"] = meta["model"] or m.group(1).strip(" ,.;")
    if not meta["brand"] or not meta["model"]:
        pats = [
            # «марки DONGFENG модели EII008,»
            r"марк[аиеу]?\s+([^,«»\"”]+?)\s+модел[ьиюе]?\s+([^,«»\"”]+?)\s*(?:,|VIN|идент|$)",
            # «markasi YUTONG, model ZK6890HQG,»
            r"markasi\s+([^,]+?)\s*,?\s+model[i]?\s*:?\s*([^,\s]+(?:\s[^,\s:]+)?)\s*(?:,|VIN)",
            # «Транспортное средство Chevrolet, Тип транспортного средства 7T12Y, …, Коммерческое наименование Damas DLX ,»
            r"Транспортное средство\s+([^,]+?)\s*,.*?Коммерческое наименование\s+([^,(]+?)\s*[,(]",
            r"Транспортное средство\s+([^,]+?)\s*,\s*Тип транспортного средства\s+([^,]+?)\s*,",
            # «ZEEKR tovar belgisi ostidagi 9X Ultra rusumli»
            r"([A-Za-zА-Яа-я0-9\-]+)\s+tovar\s+belgisi\s+ostidagi\s+(.+?)\s+rusumli",
            # «X markali Y rusumli»
            r"([A-Za-zА-Яа-я0-9\-]+)\s+markali\s+(.+?)\s+rusumli",
            # «FOTON markali BJ6931 modelli»
            r"([A-Za-zА-Яа-я0-9\-]+)\s+markali\s+(.+?)\s+modelli",
            # «Самосвал 4м3 на шасси FOTON AUMARK. Модель BJ1088VEJEA-F. Категория N2.»
            r"на\s+шасси\s+([^.,(]+?)\s*[.,]\s*Модел[ьи]\s+([A-Za-zА-Яа-я0-9\-]+)",
            # «марки MAN шасси MAN TGM-13.290 (»
            r"марк[аиеу]?\s+([A-Za-zА-Яа-я0-9\-]+)\s+шасси\s+(?:\1\s+)?([^\s(,]+)",
        ]
        for pat in pats:
            m = re.search(pat, flat, re.I)
            if m:
                meta["brand"] = meta["brand"] or m.group(1).strip()
                meta["model"] = meta["model"] or m.group(2).strip()
                break
    # общий случай без кавычек: «марка GAC AION I60 530 км, модель «…»», «Марки X.Модель Y. VIN»
    if not meta["brand"]:
        m = re.search(r"\bмарк[аиеу]?\s*:?\s+([^«»\"”]{1,80}?)\s*[,.;]?\s*модел[ьиюе]", flat, re.I)
        if m:
            meta["brand"] = m.group(1).strip(" ,.;")
    if not meta["model"]:
        m = re.search(r"\bмодел[ьиюе]?\s*:?\s*[«\"“]?\s*([^,«»\"”]{1,60}?)\s*[»\"”]?\s*[,.;]?\s*(?:VIN|идент|$)",
                      flat, re.I)
        if m:
            meta["model"] = m.group(1).strip(" ,.;")
    m = (re.search(r"\b([MМNН])([123])\s*(?:toifa|kategoriya)", flat)
         or re.search(r"категори[яи][^)]{0,40}?\b([MМNН])([123])\b", flat, re.I))
    meta["category"] = ("M" if m.group(1) in "MМ" else "N") + m.group(2) if m else ""
    meta["brand"] = clean_brand(meta["brand"]) if meta["brand"] else ""
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
            if (len(toks) >= 3 and all(re.fullmatch(r"-?\d+(?:[.,]+\d+)?,?", t) for t in toks)):
                # строка без единиц: «1 0,448 0,005 0,448 M» — измерено, U, итог
                rows.append({"meas": toks[0].rstrip(","), "final": toks[-1].rstrip(","), "unit": "",
                             "concl": mm.group(3), "unc": toks[1].rstrip(",")})
                continue
            if len(toks) >= 3 and re.fullmatch(r"-?[\d.,]+", toks[0]) and re.fullmatch(r"-?[\d.,]+", toks[-2]):
                nums = [t.rstrip(",") for t in toks if re.fullmatch(r"-?\d+(?:[.,]+\d+)?,?", t)]
                unc = nums[1] if len(nums) >= 3 else None
                rows.append({"meas": toks[0], "final": toks[-2], "unit": toks[-1], "concl": mm.group(3),
                             "unc": unc})
        if not rows:
            # запасной разбор: строка результата разорвана переносами (например, «децибел\n(dB)»)
            toks = body.split()
            numre = r"-?\d+(?:[.,]+\d+)?,?"
            for i in range(len(toks) - 1):
                if re.fullmatch(r"\d+", toks[i]) and re.fullmatch(numre, toks[i + 1]):
                    j = next((k for k in range(i + 2, min(len(toks), i + 40)) if toks[k] in ("M", "N")), None)
                    if j is None:
                        break
                    seg = toks[i + 1:j]
                    nums = [(k, t.rstrip(",")) for k, t in enumerate(seg) if re.fullmatch(numre, t)]
                    if len(nums) >= 2:
                        kf, fin = nums[-1]
                        unit = seg[kf + 1] if kf + 1 < len(seg) else ""
                        rows.append({"meas": nums[0][1], "final": fin, "unit": unit, "concl": toks[j],
                                     "unc": nums[1][1] if len(nums) >= 3 else None})
                    break
        if not rows:
            continue
        params.append({"section": fold(section), "name": name, "raw_l": name.lower(),
                       "name_f": fold(name), "rows": rows, "req": req,
                       "req_raw": norm(mr.group(1)) if mr else ""})
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
    # «до задней/от передней оконечности ТС» — это про длину, а не про передний/задний огонь
    n = re.sub(r"(передн|задн)\w*\s+оконечност\w*", "оконечность", n)
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
    if "между сигналом торможения" in n or "до сигнала торможения" in n:
        return "brake"
    if "от передней части атс" in n:
        return "len"
    if "верхней точки" in n:
        return "top"
    if "нижней точки" in n:
        return "bot"
    if "между крайними внутренними" in n:
        return "betw"
    if "габарита кузова по ширине" in n or ("по ширине" in n and "габарит" in n):
        return "width"
    if "по длине" in n and "спереди" in n:
        return "flen"
    if "задней оконечност" in n:
        return "rlen"
    if "по длине" in n:
        return "len"
    return None


LAMP_ORDER = ["dipped", "fog_front", "reverse", "pos_front", "pos_rear", "brake", "turn_front",
              "turn_rear", "side", "rear_fog", "refl_rear", "refl_front", "side_pos", "refl_side"]


def is_short(n):
    """Короткое продолжение предыдущего пункта («высота», «ширина», «минимальная глубина»…)."""
    return len(n) <= 40 and not re.search(r"сиден|подушк|поручн|огнетуш|аптеч|окн|ступен|двер|прохода", n)


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
        if lamp and pos:
            return "%s|%s" % (lamp, pos)
        return None
    ctx["lamp"] = None

    # ---------------- контекст дверей / ступенек / огнетушителя
    short = is_short(n)
    if "двер" in n and "служебн" in n and "ступен" not in n:
        ctx["door"] = "srv"
    elif "двер" in n and ("аварийн" in n or "запасн" in n) and "ступен" not in n:
        ctx["door"] = "em"
    elif "двер" in n and ("одиночн" in n or "двойн" in n) and len(n) <= 60:
        pass                                  # «Ширина, одиночная дверь» — продолжение предыдущей двери
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
    if "наклон прохода" in n:
        ctx["aisle"] = True
    elif not short:
        ctx["aisle"] = False

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

    # ---------------- зеркала, масса, категория G
    if "зеркал" in n and ("класса 5" in n or "класса 6" in n or "обзор" in section):
        return "mirror"
    if "масс" in n and ("допустим" in n or "масс" in section):
        if "25" in n and "ведущ" in n:
            return "mass_drive25"
        if "передн" in n and "ос" in n:
            return "mass_front"
        if "задн" in n and "ос" in n:
            return "mass_rear"
        if "ведом" in n:
            return "mass_axle_driven"
        if "ведущ" in n:
            return "mass_axle_drive"
        if "ось" not in n and "оси" not in n:
            return "mass_total"
    if "угол въезда" in n:
        return "g_entry"
    if "угол съезда" in n:
        return "g_exit"
    if "продольной проходимости" in n:
        return "g_ramp"
    if "просвет" in n:
        if "межосев" in n:
            return "g_clear_mid"
        if "передней ос" in n:
            return "g_clear_f"
        if "задней ос" in n:
            return "g_clear_r"

    # ---------------- огнетушитель, поручни, проход, сиденья (дополнительные блоки)
    if ctx.get("ext"):
        if "длин" in n:
            return "ext_len"
        if "высот" in n:
            return "ext_h"
        if "ширин" in n:
            return "ext_w"
    if ("наружн" in n and "ручк" in n) and ("аварийн" in n or "запасн" in n):
        return "em_handle"
    if "ступен" in n and "высота от земли" in n and "максимум" in n:
        return "step1_maxh"
    if "длина любого поручня" in n or ("длина" in n and "поручн" in n):
        return "handrail_len"
    if "поручн" in n and ("потолк" in n or "стенок" in n) and "свободное пространство" in n:
        return "handrail_gap"
    if "горизонтальные поручни" in n:
        return "handrail_h"
    if "наклон" in n and "ступен" in n:
        return "step_slope"
    if "наклон прохода" in n or (ctx.get("aisle") and "направлени" in n):
        return "aisle_trans" if "поперечн" in n else "aisle_long"
    if "высота подушки" in n:
        return "cushion_h"
    if "расстояние между сиденьями" in n:
        return "seat_dist"

    # ---------------- автобусы: надписи, аптечка, двери, окна
    if "пиктограмм" in n or ("надпись" in n and "высот" in n):
        return "pict"
    if "цифры высотой" in n:
        return "digits"
    if "аптеч" in n:
        unit = p["rows"][0]["unit"].lower()
        return "first_aid" if unit in ("mm", "мм", "cm", "см") else "first_aid_vol"
    if "орган управления" in n and ("открывания дверей" in n or "открытия дверей" in n):
        return "door_ctrl"
    if ("аварийного окна" in n or "запасного окна" in n) and "нижнего края" in n:
        return "em_win_floor"
    if "топливн" in n and "бак" in n:
        ctx["tank"] = True
        return "fuel_tank"
    if ctx.pop("tank", False) and "задней оконечност" in n:
        return "fuel_tank_rear"             # «… и менее 30 cm от задней оконечности ТС» — продолжение про бак
    if "предназначенном для ног" in n or ("для ног" in n and "уровня пола" in n):
        return "legroom_h"
    if "дымност" in n:
        return "smoke"
    if ("аварийн" in n or "запасн" in n) and "окн" in n and "площад" in n:
        return "em_win_area"
    if ("аварийн" in n or "запасн" in n) and "окн" in n:
        if "ширин" in n:
            return "em_win_w"
        if "длин" in n or "дилин" in n or "высот" in n:
            return "em_win_l"
    if ctx.get("door") and not ctx.get("ext"):
        if "двойн" in n:
            return "srv_door_w2"
        if "высот" in n:
            if ctx["door"] == "srv" and re.search(r"класс\w*\s*в\b", n):
                return "srv_door_h_b"
            return ctx["door"] + "_door_h"
        if "ширин" in n:
            return ctx["door"] + "_door_w"

    # ---------------- сиденья
    if "сплошное сиденье" in n or "ширина сиденья" in n:
        return "seat_f"
    if "ширина свободного пространства" in n:
        return "seat_free"
    if "ширина подушки" in n:
        return "seat_g"
    if "глубина подушки" in n:
        return "cushion_depth"
    if "над каждым местом для сиденья" in n or "над каждым сидячим местом" in n:
        return "seat_space_h"

    # ---------------- поручни, ограждение
    if "рукоятка располагается по вертикали" in n or ("поручн" in n and "служебн" in n and "двер" in n):
        return "handhold"
    if "огражден" in n and "вглубь" in n:
        return "guard_depth"
    if "огражден" in n:
        return "guard"

    # ---------------- огнетушитель без пункта «длина»: «высота» / «ширина», норма 200 (mm)
    if short and re.fullmatch(r"200\s*\(mm\)", p.get("req_raw") or "") and \
            p["rows"][0]["unit"].lower() in ("mm", "мм"):
        if "высот" in n:
            return "ext_h"
        if "ширин" in n:
            return "ext_w"

    # ---------------- ступеньки
    step = ctx.get("step")
    if step and ("ступен" in n or short) and "наклон" not in n:
        if "высота от земли" in n and "максимум" in n:
            return "step1_maxh"
        if ("максимальн" in n or "максимум" in n) and "высот" in n:
            return "step_maxh" if step == "other" else "step1_maxh"
        if ("минимальн" in n or "минимум" in n) and "высот" in n:
            return "step_minh"
        if ("минимальн" in n or "минимум" in n) and "глубин" in n:
            return "step_mind" if step == "other" else "step1_mind"
    return None


# --------------------------------------------------------------------------
# Раскладка по подписям шаблона (столбец D листа журнала): ключ -> правило
# Порядок важен: срабатывает первое подходящее правило. Каждый ключ — одна (первая) строка.
# --------------------------------------------------------------------------
LABEL_RULES = [
    ("heat", r"температура поверхности без покрытия|температура любой такой части"),
    ("length", r"^максимальная длина"), ("width", r"^максимальная ширина"), ("height", r"^максимальная высота"),
    ("horn", r"звукового сигнального"),
    ("glass_s", r"не являющиеся ветровыми"), ("glass_w", r"ветровые стекла"),
    ("vent_speed", r"скорость воздушных потоков"), ("vent_mob", r"подвижность воздуха"),
    ("ac_cool", r"охлаждение воздуха"),
    ("ac_speed", r"скорость воздушного потока на выходе из системы кондиционирования"),
    ("ac_temp", r"температура воздуха на выходе"), ("ac_head", r"скорость воздуха в зоне головы"),
    ("humidity", r"относительная влажность"), ("duct_temp", r"воздуховодов"),
    ("plate_a1", r"продольной плоскости"), ("plate_a2", r"опорной плоскости"),
    ("plate_hmax", r"верхнего края номерного"), ("plate_hmin", r"нижнего края номерного"),
    ("pict", r"надпись или пиктограмма"), ("digits", r"цифры"),
    ("fuel_tank_rear", r"топливный бак.*задней оконечност"), ("fuel_tank", r"топливный бак"),
    ("first_aid_vol", r"об[ъь][её]м места для аптечки"), ("first_aid", r"аптечки"),
    ("srv_door_h_b", r"служебная дверь: высота.*класс в"),
    ("srv_door_h", r"служебная дверь: высота"), ("srv_door_w2", r"служебная дверь \(двойная\).*ширина"),
    ("srv_door_w", r"служебная дверь.*ширина"),
    ("em_door_h", r"(аварийная|запасная) дверь: высота"), ("em_door_w", r"(аварийная|запасная) дверь.*ширина"),
    ("em_win_area", r"окно: площадь"),
    ("em_win2_w", r"(аварийное|запасное) окно \(второе\): ширина"),
    ("em_win2_l", r"(аварийное|запасное) окно \(второе\): длина"),
    ("em_win1_w", r"(аварийное|запасное) окно: ширина"), ("em_win1_l", r"(аварийное|запасное) окно: длина"),
    ("door_ctrl", r"органы открывания"),
    ("em_win_floor", r"нижнего края (аварийного|запасного) окна"),
    ("em_handle", r"наружной ручки (аварийной|запасной)"),
    ("step1_maxh", r"первая ступенька.*максимальная высота"), ("step1_mind", r"первая ступенька.*минимальная глубина"),
    ("step_maxh", r"другие ступеньки: максимальная высота"), ("step_minh", r"другие ступеньки: минимальная высота"),
    ("step_mind", r"другие ступеньки: минимальная глубина"),
    ("step_slope", r"наклон ступеньки"),
    ("seat_g", r"ширина подушки сиденья g"),
    ("seat_f", r"ширина сиденья \(сплошное|ширина подушки сиденья с каждой стороны"),
    ("seat_free", r"ширина свободного пространства|свободная ширина на высоте"),
    ("cushion_depth", r"глубина подушки"), ("cushion_h", r"высота подушки"),
    ("legroom_h", r"над местом для ног"),
    ("seat_space_h", r"свободное пространство над местом"), ("seat_dist", r"расстояние между сиденьями"),
    ("handhold", r"рукоятка у служебной двери|поручни у служебных дверей"),
    ("guard_depth", r"ограждение.*глубина за"),
    ("guard", r"ограждение (проёма|проема|лестничного)"),
    ("handrail_len", r"длина любого поручня"), ("handrail_gap", r"между поручнями"),
    ("handrail_h", r"горизонтальные поручни"),
    ("ext_len", r"огнетушителя: длина"), ("ext_h", r"огнетушителя: высота"), ("ext_w", r"огнетушителя: ширина"),
    ("aisle_long", r"наклон прохода в продольном"), ("aisle_trans", r"наклон прохода в поперечном"),
    ("mirror", r"зеркала класса"), ("smoke", r"дымность"),
    ("mass_drive25", r"не менее 25"),
    ("mass_front", r"на переднюю ось"), ("mass_rear", r"на заднюю ось"),
    ("mass_axle_driven", r"ведомую ось"), ("mass_axle_drive", r"ведущую ось"),
    ("mass_total", r"максимальная масса|допустимая максимальная масса"),
    ("g_entry", r"угол въезда"), ("g_exit", r"угол съезда"), ("g_ramp", r"продольной проходимости"),
    ("g_clear_f", r"просвет под передней осью"), ("g_clear_r", r"просвет под задней осью"),
    ("g_clear_mid", r"межосевой"),
]
_LABEL_RULES_C = [(k, re.compile(fold(p))) for k, p in LABEL_RULES]

# если блока с таким ключом в шаблоне нет — куда ещё можно записать значение
ALIASES = {
    "seat_g": ["seat_f"],                      # M2: «ширина подушки сиденья» — одна строка
    "first_aid_vol": ["first_aid"],            # M2: строки «объём аптечки» нет — в строку «аптечка»
    "srv_door_h": ["srv_door_h_b"],
    "mass_axle_driven": ["mass_front"], "mass_axle_drive": ["mass_rear"],
    "mass_front": ["mass_axle_driven"], "mass_rear": ["mass_axle_drive"],
}
COPY_TO = {"mass_axle_drive": "mass_drive25", "mass_rear": "mass_drive25"}   # ведущая ось → и в блок «не менее 25 %»


def build_layout(cells):
    """Найти блоки замеров по подписям шаблона: ключ -> (подпись, первая строка)."""
    J = cells.get(JOURNAL_SHEET, {})
    rows = sorted({int(re.sub(r"\D", "", k)) for k in J if k.startswith("D")})
    layout = {}
    for r in rows:
        d = J.get("D%d" % r)
        e = J.get("%s%d" % (JCOL, r))
        m = J.get("%s%d" % (JCOL, r + 3))
        if not isinstance(d, str) or fold(d) in ("ср.знач", "ср.занч") or "решени" in fold(d):
            continue
        if e is None or not (isinstance(m, str) and m.startswith("=")):
            continue
        n = fold(d)
        lamp, pos = detect_lamp(n), lamp_pos(n)
        if lamp and pos:
            key = "%s|%s" % (lamp, pos)
        else:
            key = next((k for k, rx in _LABEL_RULES_C if rx.search(n)), None)
        if key and key not in layout:
            layout[key] = (" ".join(d.split()), r)
    return layout


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
        unc, unc_dp = param_unc(p)
        if key in DIM_KEYS and (unit in ("mm", "мм") or v >= 100):
            # «8940 m, U 0,003 m» — опечатка в единице значения: значение в mm, а U уже в m
            if unc is not None and unit in ("mm", "мм"):
                unc, unc_dp = unc / 1000, unc_dp + 3
            v, dp = v / 1000, dp + 3                   # mm -> m (в журнале габариты в m)
            notes = notes + ["переведено из mm в m: %s" % fmt(v)]
        if p.get("_guess"):
            notes = notes + ["в протоколе нет названия огня — по порядку пунктов отнесено к «%s»"
                             % LAMPS[p["_guess"]][0]]
        ok = all(r["concl"] == "M" for r in p["rows"])
        values[key] = {"value": v, "dp": dp, "ok": ok, "name": p["name"], "notes": notes,
                       "unc": unc, "unc_dp": unc_dp, "req": p["req"],
                       "req_raw": p.get("req_raw", "")}
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


@functools.lru_cache(maxsize=None)
def _t95(df):
    return t_inv_2t(0.05, df)


def u_of_s(s, ub):
    """Расширенная неопределённость по формулам шаблона:
    UA=STDEVA/√3, u=√(UA²+UB²), veff=u⁴/(UA⁴/2), U=TINV(0.05;veff)·u."""
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
        if n == "MATCH":
            look, arr = args[0], args[1] if isinstance(args[1], list) else [args[1]]
            for i, x in enumerate(arr):
                if (isinstance(x, str) and isinstance(look, str) and x.strip().lower() == look.strip().lower()) or \
                        (not isinstance(x, str) and not isinstance(look, str) and x is not None and look is not None
                         and _num(x) == _num(look)):
                    return float(i + 1)
            raise XlErr("#N/A")
        if n == "INDEX":
            arr = args[0] if isinstance(args[0], list) else [args[0]]
            i = int(_num(args[1])) if len(args) > 1 else 1
            if i < 1 or i > len(arr):
                raise XlErr("#REF!")
            v = arr[i - 1]
            return 0.0 if v is None else v
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


def block_ub(cells, book, row):
    """UB блока: F[r] -> лист!U (=t*u) -> u (=SQRT(UA^2+UB^2)) -> UB. None — цепочка нестандартная."""
    try:
        f = cells[JOURNAL_SHEET].get("F%d" % row)
        m = re.fullmatch(r"='?((?:[^']|'')+?)'?!\$?([A-Z]+)\$?(\d+)", f or "")
        if not m:
            return None
        sh = m.group(1).replace("''", "'")
        uc = cells[sh].get(m.group(2) + m.group(3))
        mm = re.fullmatch(r"=([A-Z]+\d+)\*([A-Z]+\d+)", uc or "")
        if not mm:
            return None
        tf, uf = cells[sh].get(mm.group(1)), cells[sh].get(mm.group(2))
        m3 = re.fullmatch(r"=SQRT\(([A-Z]+\d+)\^2\+([A-Z]+\d+)\^2\)", uf or "")
        mt = re.fullmatch(r"=TINV\(0\.05,([A-Z]+\d+)\)", tf or "")
        if not (m3 and mt):
            return None
        if not re.fullmatch(r"=STDEVA\([A-Z]+\d+:[A-Z]+\d+\)/SQRT\(3\)", cells[sh].get(m3.group(1)) or ""):
            return None
        ub = book.value(sh, m3.group(2))
        return (float(ub), sh, m3.group(2)) if isinstance(ub, (int, float)) else None
    except Exception:  # noqa: BLE001
        return None


def proto_norm(v, kind, old):
    """Норма блока по требованию протокола: (норма в единицах шаблона, None) или (None, причина).
    kind: "ge" — «не менее», "le" — «не более»; old — норма шаблона.
    Кандидаты: разобранная граница, затем все числа требования («1800 допускается до 2500», «4 цилиндров: 400»);
    берётся первый, которому значение протокола удовлетворяет (так протокол с выводом «M» не противоречит журналу)."""
    lo, hi = v.get("req") or (None, None)
    if (kind == "ge" and lo is None and hi is not None) or (kind == "le" and hi is None and lo is not None):
        return None, ""                          # протокол задаёт другую границу («не более 100» для блока 65–100)
    raw = re.sub(r"\([^)]*\)", " ", v.get("req_raw") or "")
    cands = [x for x in ((lo if kind == "ge" else hi),) if x is not None]
    words = [w for w in re.findall(r"[^\W\d_]+", raw.lower())
             if w not in ("не", "менее", "более", "ниже", "выше", "должна", "должен", "должно", "составлять",
                          "превышать", "допускается", "до", "можно", "увел", "max", "min", "максимум", "минимум",
                          "mm", "cm", "m", "мм", "см", "м")]
    if len(words) <= 3:                          # короткое требование: «36», «1800 допускается до 2500», «4 цилиндров: 400»
        cands += [to_dec(x) for x in re.findall(r"\d+(?:[.,]\d+)?", raw)]
    if not cands:
        return None, ""
    val = v["value"]

    def scaled(c):                               # единицы требования (m) ↔ шаблона (mm) и наоборот
        if c and old and c != 0:
            r = abs(old / c)
            if 300 <= r <= 3000:
                return c * 1000
            if 300 <= 1 / r <= 3000 if r else False:
                return c / 1000
        return c

    for c in cands:
        c = scaled(c)
        if c is None or c == 0 and kind == "ge" and val < 0:
            continue
        if (kind == "ge" and val >= c) or (kind == "le" and val <= c):
            return c, ""
    explicit = lo is not None or hi is not None
    fits_old = (kind == "ge" and val >= old) or (kind == "le" and val <= old)
    if v.get("ok") and (explicit or not fits_old):
        # без «не более/не менее» и значение в норме шаблона — норма шаблона, не противоречие («100 (cm)» для 65–100)
        return None, "значение %s вне нормы протокола («%s»), а вывод протокола «M» — норма оставлена из шаблона" % (
            fmt(val), (v.get("req_raw") or "").strip())
    return None, ""


def block_norm_cell(cells, row):
    """Ячейка нормы блока по формуле ППР (E[r+4] -> лист!ППР = [1-]NORM.DIST(норма,…)):
    (лист, ячейка, "ge"|"le") или None."""
    m = re.fullmatch(r"='?((?:[^']|'')+?)'?!\$?([A-Z]+)\$?(\d+)", cells[JOURNAL_SHEET].get("E%d" % (row + 4)) or "")
    if not m:
        return None
    sh = m.group(1).replace("''", "'")
    f = cells.get(sh, {}).get(m.group(2) + m.group(3)) or ""
    mm = re.fullmatch(r"=(1-)?(?:_xlfn\.)?NORM\.DIST\(\$?([A-Z]+)\$?(\d+),[^()]*\(?\)?\)", f)
    if not mm:
        return None
    return sh, mm.group(2) + mm.group(3), ("ge" if mm.group(1) else "le")


# Запасные значения UB (вида «=a/2»), если с UB шаблона неопределённость протокола недостижима
UB_STEPS = ("0.3", "0.2", "0.1", "0.05", "0.03", "0.02", "0.01", "0.005", "0.003", "0.002", "0.001",
            "0.0005", "0.0002", "0.0001")


def set_formula(xml, ref, formula):
    m = cell_re(ref).search(xml)
    if not m:
        return xml, False
    attrs = T_ATTR_RE.sub("", m.group(1))
    return xml[:m.start()] + '<c r="%s"%s><f>%s</f></c>' % (ref, attrs, escape(formula)) + xml[m.end():], True


def num_or_text(v):
    return fmt(D(v)) if v is not None else ""


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


def fill_workbook(template_bytes, meta, values, mode="corr"):
    """mode: "meas" — среднее тройки = протоколу (как в шаблоне);
             "corr" — «СР.ЗАНЧ» (с поправкой Og'ish) = протоколу."""
    tcells = load_cells(template_bytes)
    if JOURNAL_SHEET not in tcells:
        raise ValueError("В шаблоне нет листа «%s» — это не шаблон %s." % (JOURNAL_SHEET, CAT))
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
    ubs = {key: block_ub(tcells, tbook, row) for key, (label, row) in blocks.items()}

    log = {"filled": [], "cleared": [], "inexact": [], "missing_cells": [], "verify_bad": [],
           "calc": True, "unlinked": [], "u_fit": [], "u_bad": [], "u_skip": [], "ub_changed": [],
           "norm_changed": [], "proto_contra": []}
    ub_cells = {}   # лист -> {ячейка: новая формула UB}
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
    if meta.get("id"):
        head.append((ID_REF, "ID № %s" % meta["id"]))
    for ref, v in head:
        if ref in ("D10", "D11", "D14") and not (v and v.strip()):
            v = ""                        # не найдено — пусто, а не значение из шаблона
        if v and v.strip() or ref in ("D10", "D11", "D14"):
            j, ok = set_str(j, ref, v)
            jc[ref] = v
            if not ok:
                log["missing_cells"].append(ref)

    # --- замеры
    block_rows = sorted({row for label, row in blocks.values()} | {436})
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
            new = None
            if p.get("unc") and p["unc"] > 0:
                if ubs[key] is None:
                    log["u_skip"].append(label)
                else:
                    ub0, ush, ucell = ubs[key]
                    new, extra, uok = fit_u_triplet(seeds[key], target, dp, p["unc"], p["unc_dp"], ub0)
                    if new is None or not uok:
                        # с UB шаблона так не получается — берём ближайшее меньшее «круглое» UB
                        for a in UB_STEPS:
                            ub1 = float(a) / 2
                            if ub1 >= ub0:
                                continue
                            n2, e2, ok2 = fit_u_triplet(seeds[key], target, dp, p["unc"], p["unc_dp"], ub1)
                            if n2 is not None and ok2:
                                old_f = tcells[ush].get(ucell)
                                ub_cells.setdefault(ush, {})[ucell] = "=%s/2" % a
                                log["ub_changed"].append("%s: UB %s → =%s/2 ('%s'!%s)" % (
                                    label, old_f, a, ush.strip(), ucell))
                                new, extra, uok = n2, e2, ok2
                                break
                    if new is not None:
                        exact = True
                        log["u_fit"].append((label, p["unc"], p["unc_dp"], extra, uok, key))
                    else:
                        log["u_bad"].append("%s: протокол %s, а формулы шаблона дают не меньше %s" % (
                            label, fmt(p["unc"]), round(u_of_s(1e-9, ub0), p["unc_dp"] + 2)))
            if new is None:
                new, exact = gen_triplet(seeds[key], target, dp)
            for ref, x in zip(cells, new):
                j, ok = set_input(j, ref, fmt(x))
                if ok:
                    jc[ref] = float(x)
                else:
                    log["missing_cells"].append(ref)
            if p.get("unc") is not None and p["unc"] == 0 and len(set(new)) == 1:
                # в протоколе U = 0 и все замеры одинаковые: в журнале тоже 0 (а не #DIV/0!)
                uref = "F%d" % row
                f = jc.get(uref)
                if isinstance(f, str) and f.startswith("="):
                    nf = "IF(STDEVA(%s%d:%s%d)=0,0,%s)" % (JCOL, row, JCOL, row + 2, f[1:])
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
                j, done = clear_input(j, ref)
                if done:
                    jc.pop(ref, None)
            if had:
                log["cleared"].append(label)
            # среднее, решение и неопределённость пустого блока — пустые (без #DIV/0! и чисел),
            # но при ручном вводе замеров формулы снова считают
            guard = 'COUNT(%s%d:%s%d)=0' % (JCOL, row, JCOL, row + 2)
            end = next((r for r in block_rows if r > row), row + 6)
            for rr in range(row, end):
                for col in (JCOL, "F"):
                    if col == JCOL and rr < row + 3:
                        continue
                    ref = "%s%d" % (col, rr)
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
        j = restyle(j, "%s%d" % (JCOL, row + 3), styles, num_code(dp_))
        pv = values[key_]
        if pv.get("unc") is not None:
            j = restyle(j, "F%d" % row, styles, num_code(pv["unc_dp"]))
    out["xl/styles.xml"] = styles.render().encode("utf8")
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
    # --- лист «Масса»: тип ТС для нормы (двухосные / трехосные / сочлененные / четырехосные) — по протоколу
    mass_cells = {}
    ms_path = next((p for n, p in paths if n == "Масса"), None)
    msc = tcells.get("Масса", {})
    if ms_path and "mass_total" in values and isinstance(msc.get("D6"), str):
        nm = fold(values["mass_total"]["name"])
        opt = {k: msc.get(k) for k in ("K7", "K8", "K9", "K10")}
        choice = (opt["K9"] if "сочлен" in nm else opt["K10"] if "четырехос" in nm else
                  opt["K8"] if "трехос" in nm else opt["K7"] if "двухос" in nm else None)
        if isinstance(choice, str) and choice != msc.get("D6"):
            x = (out[ms_path] if ms_path in out else zin.read(ms_path)).decode("utf8")
            x, ok = set_str(x, "D6", choice)
            if ok:
                out[ms_path] = x.encode("utf8")
                mass_cells["D6"] = choice
                log["notes_mass"] = "лист «Масса»: тип ТС «%s»" % choice
    # --- нормы из протокола: журнал повторяет протокол (норма блока = требованию протокола)
    norm_cells = {}
    for key in values:
        if key not in blocks:
            continue
        nc = block_norm_cell(tcells, blocks[key][1])
        if nc is None:
            continue
        old = tcells.get(nc[0], {}).get(nc[1])
        if not isinstance(old, (int, float)) or isinstance(old, bool):
            continue
        old = D(repr(old))
        want, why = proto_norm(values[key], nc[2], old)
        if want is None:
            if why:
                log["proto_contra"].append("%s: %s" % (blocks[key][0], why))
            continue
        if want == old:
            continue
        norm_cells.setdefault(nc[0], {})[nc[1]] = float(want)
        log["norm_changed"].append("%s: норма %s → %s (по протоколу, '%s'!%s)" % (
            blocks[key][0], fmt(old), fmt(want), nc[0].strip(), nc[1]))
    for n, p in paths:
        if p not in out:
            out[p] = zin.read(p)
        if n in norm_cells:
            x = out[p].decode("utf8")
            for ref, v in norm_cells[n].items():
                x, _ = set_num(x, ref, "%g" % v)
            out[p] = x.encode("utf8")
        if n in ub_cells:
            x = out[p].decode("utf8")
            for ref, f in ub_cells[n].items():
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

    # --- расчёт формул и запись готовых значений
    try:
        allc = dict(tcells)
        allc[JOURNAL_SHEET] = jc
        if pr_cells:
            allc[pr_sheet] = dict(allc[pr_sheet], **pr_cells)
        for sh_, cc in ub_cells.items():
            allc[sh_] = dict(allc[sh_], **cc)
        for sh_, cc in norm_cells.items():
            allc[sh_] = dict(allc[sh_], **cc)
        if mass_cells:
            allc["Масса"] = dict(allc["Масса"], **mass_cells)
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
        # контроль неопределённости: U в журнале (колонка F) = протоколу после округления
        for label, unc, unc_dp, extra, uok, key in log["u_fit"]:
            row = blocks[key][1]
            r = book.result(JOURNAL_SHEET, "F%d" % row)
            got = r[1] if r[0] == "v" and isinstance(r[1], (int, float)) else None
            if got is None or not _round_eq(got, unc, unc_dp):
                log["u_bad"].append("%s: протокол %s, в журнале %s" % (
                    label, fmt(unc), "ошибка" if got is None else round(got, unc_dp + 2)))
    except Exception:  # noqa: BLE001
        log["calc"] = False
        for n, p in paths:
            out[p] = strip_cache(out[p].decode("utf8")).encode("utf8")
    return pack_out(zin, out), log


def process(pdf_bytes, template_bytes, mode="corr", file_id=""):
    raw = pdf_text(pdf_bytes)
    if "Aniqlanayotgan parametr nomi" not in raw:
        raise ValueError("В PDF не найден текст протокола (возможно, это скан без текстового слоя).")
    meta = parse_meta(raw)
    meta["id"] = (file_id or "").strip()
    params = parse_params(raw)
    if not params:
        raise ValueError("В протоколе не удалось найти таблицы результатов.")
    global LAYOUT
    LAYOUT = {}
    LAYOUT = build_layout(load_cells(template_bytes))
    values, unmatched, dup, trace = map_params(params)
    placed = {}
    for key, v in values.items():
        target = key if key in LAYOUT else next((a for a in ALIASES.get(key, []) if a in LAYOUT and a not in values
                                                 and a not in placed), None)
        if target is None and COPY_TO.get(key) in LAYOUT and COPY_TO[key] not in values \
                and COPY_TO[key] not in placed:
            target = COPY_TO[key]
        if target is None:
            v["_nospace"] = True
            continue
        placed[target] = v
        cp = COPY_TO.get(key)
        if cp and cp in LAYOUT and cp not in values:
            placed.setdefault(cp, v)
    nospace_names = {v["name"] for v in values.values() if v.get("_nospace")}
    for p in params:
        if p["name"] in nospace_names and p not in unmatched:
            unmatched.append(p)
    trace = [(p, (k if not (k and values.get(k, {}).get("_nospace")) else None)) for p, k in trace]
    values = placed
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


def id_from_filename(fname):
    """ID = название PDF-файла: «1901093.pdf», «1901093 (1).pdf» -> «1901093»."""
    stem = Path(fname or "").stem
    m = re.search(r"\d{4,}", stem)
    return m.group(0) if m else stem.strip()


def safe_name(s):
    return re.sub(r"[^\w\-.]+", "_", s, flags=re.U).strip("_")


# --------------------------------------------------------------------------
# Интерфейс Streamlit
# --------------------------------------------------------------------------
MODES = {"«СР.ЗАНЧ» (среднее с поправкой Og'ish) = протоколу": "corr",
         "Среднее самих замеров = протоколу (без поправки)": "meas"}


XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def out_name(meta):
    if meta.get("id"):
        return "%s.xlsx" % safe_name(meta["id"])
    return "Журнал_%s_%s_%s_%s.xlsx" % (CAT, safe_name(meta["brand"]), safe_name(meta["model"]), meta["vin"])


def problems(meta, log):
    """Короткий список того, что требует внимания (для сводки)."""
    res = []
    if log["verify_bad"] or log.get("unlinked"):
        res.append("среднее не сошлось")
    if log.get("u_bad"):
        res.append("U не сошлась: %d" % len(log["u_bad"]))
    if not log["calc"]:
        res.append("формулы не посчитаны")
    if log["nonconf"]:
        res.append("есть N в протоколе")
    if log.get("suspect"):
        res.append("нераспознанные огни")
    if not log["filled"]:
        res.append("нет пунктов для журнала")
    if meta.get("category") and meta["category"] != CAT:
        res.append("категория в протоколе %s, а журнал %s" % (meta["category"], CAT))
    if not meta["brand"] or not meta["model"]:
        res.append("нет марки/модели")
    if not meta["vin"]:
        res.append("в протоколе нет VIN")
    if log["notes"]:
        res.append("расхождения в протоколе: %d" % len(log["notes"]))
    if log.get("ub_changed"):
        res.append("UB изменена: %d" % len(log["ub_changed"]))
    if log.get("norm_changed"):
        res.append("норма по протоколу: %d" % len(log["norm_changed"]))
    if log.get("proto_contra"):
        res.append("противоречия в протоколе: %d" % len(log["proto_contra"]))
    return res


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


def main():
    st.set_page_config(page_title="Журнал %s" % CAT, page_icon="🚌", layout="centered")
    st.markdown(
        "<style>.stButton>button,.stDownloadButton>button{min-height:3rem;font-size:1.05rem}"
        ".block-container{padding-top:1.2rem;padding-bottom:2rem}</style>",
        unsafe_allow_html=True,
    )
    st.title("🚌 Журнал испытаний %s" % CAT)
    st.caption("Автобусы категории %s%s. Загрузите один или сразу несколько PDF-протоколов — " % (CAT, CAT_NOTE) +
               "получите заполненные журналы по шаблону.")

    tpl_path = Path(__file__).resolve().parent / TEMPLATE_NAME
    template_bytes = tpl_path.read_bytes() if tpl_path.exists() else None
    if template_bytes is None:
        up_t = st.file_uploader("Шаблон журнала %s (%s)" % (CAT, TEMPLATE_NAME), type=["xlsx"], key="tpl")
        if up_t is not None:
            template_bytes = up_t.getvalue()
    else:
        with st.expander("Шаблон и настройки"):
            st.write("Используется %s рядом с app.py." % TEMPLATE_NAME)
            up_t = st.file_uploader("Заменить шаблон", type=["xlsx"], key="tpl2")
            if up_t is not None:
                template_bytes = up_t.getvalue()
            st.radio("С протоколом совпадает:", list(MODES), key="mode")

    pdfs = st.file_uploader("PDF-протоколы испытаний (можно несколько)", type=["pdf"],
                            accept_multiple_files=True, key="pdfs") or []
    ids = unique_ids([f.name for f in pdfs])
    if len(pdfs) == 1:
        ids[0] = st.text_input("ID (берётся из названия файла)", value=ids[0], key="id_" + pdfs[0].name)
    elif pdfs:
        st.caption("ID берутся из названий файлов: " + ", ".join(ids))

    label = "Заполнить журнал" if len(pdfs) <= 1 else "Заполнить журналы (%d)" % len(pdfs)
    if st.button(label, type="primary", use_container_width=True,
                 disabled=(not pdfs or template_bytes is None)):
        mode = MODES.get(st.session_state.get("mode"), "corr")
        results = []
        bar = st.progress(0.0, text="Обработка…")
        for n, (f, fid) in enumerate(zip(pdfs, ids)):
            bar.progress(n / len(pdfs), text="Обработка %d из %d: %s" % (n + 1, len(pdfs), f.name))
            try:
                data, meta, log = process(f.getvalue(), template_bytes, mode, fid)
                results.append({"name": f.name, "data": data, "meta": meta, "log": log, "err": None})
            except Exception as e:  # noqa: BLE001
                results.append({"name": f.name, "data": None, "meta": None, "log": None, "err": str(e)})
        bar.empty()
        st.session_state["batch"] = results

    if template_bytes is None:
        st.info("Положите %s рядом с app.py или загрузите шаблон выше." % TEMPLATE_NAME)
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
                           file_name="Журналы_%s.zip" % CAT, mime="application/zip", use_container_width=True)
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
        hard = any(x.startswith(("среднее", "U не", "формулы", "нераспозн", "нет марки", "нет пунктов", "категория"))
                   for x in pr)
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
        "- **ID:** %s\n- **Протокол:** %s\n- **Марка / модель:** %s / %s\n- **Категория:** %s\n- **VIN:** %s\n"
        "- **Дата испытаний:** %s\n- **Условия:** %s °C, %s %%, %s кПа, %s м/с" % (
            meta.get("id") or "—", meta["protocol"] or "—", meta["brand"] or "—", meta["model"] or "—", meta["category"] or "—",
            meta["vin"] or "—", d.strftime("%d.%m.%Y") if d else "—",
            meta["temp"], meta["hum"], meta["press"], meta["wind"]))
    if len(meta.get("vins", [])) > 1:
        st.warning("В протоколе несколько VIN, в журнал записан первый. Все: " + ", ".join(meta["vins"]))
    if not log["filled"]:
        st.error("В протоколе нет ни одного пункта, для которого есть место в журнале %s — заполнять нечего. " % CAT +
                 "Пункты протокола: " + "; ".join(x[:80] for x in log["unmatched"]))
    if meta.get("category") and meta["category"] != CAT:
        st.error("В протоколе категория %s, а это журнал %s — проверьте, тот ли шаблон (приложение) выбран."
                 % (meta["category"], CAT))
    if not meta["brand"] or not meta["model"]:
        st.warning("Марка/модель не распознаны — впишите вручную (НЕОПРЕДЕЛ!D10, D11).")
    if not meta["vin"]:
        st.info("В протоколе нет VIN — ячейка D14 оставлена пустой.")

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
    if log.get("u_bad"):
        st.error("Неопределённость не удалось сделать как в протоколе — формулы шаблона "
                 "(UB, TINV) не дают такого значения; поставлено ближайшее возможное:\n" +
                 "\n".join("- %s" % x for x in log["u_bad"]))
    elif log.get("u_fit") and log["calc"]:
        st.success("Проверка: неопределённость (колонка F) совпадает с протоколом у %d блоков."
                   % len(log["u_fit"]))
    if log.get("norm_changed"):
        st.info("Норма в журнале поставлена как в протоколе:\n" +
                "\n".join("- %s" % x for x in log["norm_changed"]))
    if log.get("proto_contra"):
        st.warning("Противоречие в самом протоколе — значение не укладывается в норму протокола, а вывод «M» "
                   "(значение в журнал записано как в протоколе):\n" +
                   "\n".join("- %s" % x for x in log["proto_contra"]))
    if log.get("ub_changed"):
        st.warning("Чтобы неопределённость совпала с протоколом, в этих блоках уменьшена погрешность "
                   "прибора UB (в шаблоне она больше, чем допускает протокол). Сверьте с паспортом прибора:\n" +
                   "\n".join("- %s" % x for x in log["ub_changed"]))
    if log.get("u_skip"):
        st.warning("Для этих блоков неопределённость не подгонялась (нестандартные формулы шаблона): "
                   + "; ".join(log["u_skip"]))
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
    with st.expander("Неопределённость: протокол → журнал (%d)" % len(log.get("u_fit", []))):
        for label, unc, unc_dp, extra, uok, key in log.get("u_fit", []):
            st.write("%s **%s**: U = %s%s" % ("✅" if uok else "⚠️", label, fmt(unc),
                                              (" (добавлено знаков после запятой: %d)" % extra) if extra else ""))
    st.caption("≈ — у этих блоков шаг разброса из шаблона слегка изменён (на 1–2 единицы последнего "
               "знака), чтобы среднее было точно равно значению протокола.")


if __name__ == "__main__":
    main()
