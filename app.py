# -*- coding: utf-8 -*-
import io
import re
import math
import zipfile
from pathlib import Path
from datetime import date
from decimal import Decimal, getcontext, ROUND_HALF_UP
from fractions import Fraction
from xml.sax.saxutils import escape, unescape

import streamlit as st
import openpyxl
from pypdf import PdfReader

getcontext().prec = 60

JOURNAL = 'Рабочий журнал'

MONTHS_GEN = ['января', 'февраля', 'марта', 'апреля', 'мая', 'июня',
              'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря']

SINGLE_ROWS = [21, 26, 31, 36, 41, 46, 51, 56, 61, 66, 71, 76, 81, 86, 91, 96,
               101, 106, 112, 118, 123, 128, 133, 138, 143, 149, 158, 163,
               169, 174, 179, 184, 191, 196, 201, 206]
ROW_LABEL = {21: 'Педаль E', 26: 'Педаль H', 31: 'Педаль J', 36: 'CO (отработавшие газы)',
             41: 'CH (отработавшие газы)', 46: 'Звуковой сигнал', 51: 'Длина',
             56: 'Ширина', 61: 'Высота', 66: 'Скорость воздуха на выходе вентиляции',
             71: 'Подвижность воздуха', 76: 'Кондиционер: охлаждение',
             81: 'Скорость на выходе кондиционера', 86: 'Температура воздуха',
             91: 'Скорость в зоне головы', 96: 'Влажность',
             101: 'Температура воздуховодов', 106: 'Отопление',
             112: 'Выступающие части', 118: 'Решётки и щели (зазоры)',
             123: 'Решётки и щели (радиус)', 128: 'Стеклоочиститель',
             133: 'Бамперы', 138: 'Двери', 143: 'Ручки двери',
             149: 'Обшивка кузова', 158: 'Светопропускание ветрового стекла',
             163: 'Светопропускание бокового стекла',
             169: 'Номерной знак: угол б', 174: 'Номерной знак: угол в',
             179: 'Номерной знак: высота min', 184: 'Номерной знак: высота max',
             191: 'CO в воздухе кабины', 196: 'NO₂ в воздухе кабины',
             201: 'NO в воздухе кабины', 206: 'Внешний шум (разгон)'}

LIGHT_ROWS = {'ближний': 214, 'передний противотуманный': 219, 'задний ход': 224,
              'габариты передние': 229, 'габариты задние': 234, 'торможение': 239,
              'указатели передние': 244, 'указатели задние': 249, 'повторители': 254,
              'задний противотуманный': 261, 'задние световозвращатели': 266,
              'передние световозвращатели': 271}
LIGHT_LABEL = {214: 'Ближний свет', 219: 'Передний противотуманный',
               224: 'Задний ход', 229: 'Габаритные огни передние',
               234: 'Габаритные огни задние', 239: 'Сигнал торможения',
               244: 'Указатели поворота передние', 249: 'Указатели поворота задние',
               254: 'Боковые повторители', 261: 'Задний противотуманный фонарь',
               266: 'Световозвращатели задние', 271: 'Световозвращатели передние'}
COLS = {'F': 6, 'G': 7, 'H': 8, 'I': 9}

LAT2CYR = {'a': 'а', 'b': 'в', 'c': 'с', 'e': 'е', 'h': 'н', 'k': 'к',
           'm': 'м', 'o': 'о', 'p': 'р', 't': 'т', 'x': 'х', 'y': 'у'}


class XLError(Exception):
    def __init__(self, code):
        self.code = code
        super().__init__(code)


def norm_txt(s):
    s = s.lower().replace('ё', 'е')
    s = ''.join(LAT2CYR.get(ch, ch) for ch in s)
    return re.sub(r'\s+', ' ', s).strip()


def dec_of_tok(tok):
    tok = tok.replace(',', '.')
    return Decimal(tok)


def prec_of_tok(tok):
    tok = tok.replace(',', '.')
    if '.' in tok:
        return len(tok.split('.')[1].rstrip('0')) or 1
    return 0


def dec_fr(fr, pr=50):
    return Decimal(fr.numerator) / Decimal(fr.denominator) if fr else Decimal(0)


def n_dec(d):
    return -d.as_tuple().exponent if d.as_tuple().exponent < 0 else 0


def excel_round(d, p):
    return d.quantize(Decimal(1).scaleb(-p), rounding=ROUND_HALF_UP)


HEADER_RE = re.compile(r'Sinov bayonnomasi № SB-.*?etiladi\.', re.S)
ROW_RE = re.compile(r'^\s*(\d+)\s+(-?\d+(?:[.,]\d+)?)\s+(\S+)\s+'
                    r'(-?\d+(?:[.,]\d+)?)\s+(\S+)\s+'
                    r'(-?\d+(?:[.,]\d+)?)\s+(\S+)\s+([MN])\s*$', re.M)
SEC_RE = re.compile(r"(\d+)\.\s*Sinalayotgan\s+ko[’'`]?rsatgich\s+(.*?)\s+Texnik\s+talablarni", re.S)
PAR_RE = re.compile(r"(\d+)\.\s*Aniqlanayotgan\s+parametr\s+nomi\s+(.*?)\s+Maxsulotga", re.S)
MIN_RE = re.compile(r'(?:не менее|не ниже|камида|not\s+less|минимал)\s*:?\s*(-?\d+(?:[.,]\d+)?)')
MAX_RE = re.compile(r'(?:не более|не превышает|ортик|not\s+more|максимал|should\s+not\s+exceed)\s*:?\s*(-?\d+(?:[.,]\d+)?)')
RANGE1_RE = re.compile(r'(?:от|from)\s+(-?\d+(?:[.,]\d+)?)\s+(?:до|to|–|—|-)\s*(-?\d+(?:[.,]\d+)?)')
RANGE2_RE = re.compile(r'(?<![\d/.,])(-?\d+(?:[.,]\d+)?)\s*[–—]\s*(-?\d+(?:[.,]\d+)?)')


def parse_req(text):
    m = RANGE1_RE.search(text) or RANGE2_RE.search(text)
    if m:
        a, b = sorted([dec_of_tok(m.group(1)), dec_of_tok(m.group(2))])
        return ('range', a, b)
    m = MIN_RE.search(text)
    if m:
        return ('min', dec_of_tok(m.group(1)), None)
    m = MAX_RE.search(text)
    if m:
        return ('max', dec_of_tok(m.group(1)), None)
    return ('none', None, None)


def req_fits(req, v):
    kind, a, b = req
    if kind == 'min':
        return v >= a
    if kind == 'max':
        return v <= a
    if kind == 'range':
        return a <= v <= b
    return False


def parse_pdf(pdf_bytes):
    if not pdf_bytes[:5] == b'%PDF-':
        raise ValueError('Файл не является PDF.')
    reader = PdfReader(io.BytesIO(pdf_bytes))
    raw = '\n'.join((pg.extract_text() or '') for pg in reader.pages)
    if not raw.strip():
        raise ValueError('В PDF нет текста: похоже на скан без распознавания.')
    text = HEADER_RE.sub('', raw)

    meta = {}
    m = re.search(r'\b(SB-\S+)', raw)
    meta['protocol'] = m.group(1) if m else ''
    m = re.search(r'марка\s*«([^»]+)»', raw, re.I)
    meta['marka'] = m.group(1).strip() if m else ''
    m = re.search(r'модель\s*«([^»]+)»', raw, re.I)
    meta['model'] = m.group(1).strip() if m else ''
    m = re.search(r'\b([A-HJ-NPR-Z0-9]{17})\b', raw)
    meta['vin'] = m.group(1) if m else ''
    m = re.search(r'Sinov boshlanish vaqti[^0-9]*(\d{1,2})[.](\d{1,2})[.](\d{4})', raw)
    if not m:
        m = re.search(r'Sanasi[^0-9]*(\d{1,2})[.](\d{1,2})[.](\d{4})', raw)
    if m:
        meta['date'] = date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    else:
        meta['date'] = None

    def cond(name):
        mm = re.search(name + r'[^0-9\-]*(-?\d+(?:[.,]\d+)?)\s*([A-Za-z/²³%°]*)', raw)
        if not mm:
            return None
        v = dec_of_tok(mm.group(1))
        unit = mm.group(2).lower()
        if unit in ('gpa', 'hpa', 'гпа'):
            v = v / Decimal(10)
        return v

    meta['temp'] = cond(r'Harorat')
    meta['hum'] = cond(r'Nisbiy namlik')
    meta['wind'] = cond(r'Havo tezligi')
    meta['press'] = cond(r'Atmosfera bosimi')

    secs = [(m.start(), m.group(2).strip()) for m in SEC_RE.finditer(text)]
    params = []
    for i, (spos, sname) in enumerate(secs):
        send = secs[i + 1][0] if i + 1 < len(secs) else len(text)
        chunk = text[spos:send]
        pars = list(PAR_RE.finditer(chunk))
        for j, pm in enumerate(pars):
            pend = pars[j + 1].start() if j + 1 < len(pars) else len(chunk)
            pchunk = chunk[pm.end() - 10:pend]
            name = re.sub(r'\s+', ' ', pm.group(2)).strip()
            name = re.sub(r'^\d+\s*[.)]?\s+(?=[А-Яа-ЯA-Za-z])', '', name).strip()
            req = parse_req(pchunk)
            results = []
            after = pchunk.split('Xulosa')[-1]
            for rm in ROW_RE.finditer(after):
                results.append({'meas_tok': rm.group(2), 'unc_tok': rm.group(4),
                                'res_tok': rm.group(6), 'flag': rm.group(8)})
            params.append({'section': sname, 'name': name, 'req': req, 'results': results})
    if not params:
        raise ValueError('Не удалось распознать разделы протокола: неподдерживаемый формат PDF.')
    return meta, params


def detect_light(nn):
    if 'повторител' in nn or ('боков' in nn and 'поворот' in nn):
        return 'повторители'
    if 'световозвращател' in nn:
        if 'задн' in nn:
            return 'задние световозвращатели'
        if 'передн' in nn:
            return 'передние световозвращатели'
        return None
    if 'противотуманн' in nn:
        return 'задний противотуманный' if 'задн' in nn else 'передний противотуманный'
    if 'ближн' in nn:
        return 'ближний'
    if 'заднего хода' in nn:
        return 'задний ход'
    if 'габаритн' in nn:
        return 'габариты задние' if 'задн' in nn else 'габариты передние'
    if 'торможен' in nn:
        return 'торможение'
    if 'поворот' in nn:
        return 'указатели задние' if 'задн' in nn else 'указатели передние'
    return None


def detect_pos(nn):
    if 'от передней части атс до края' in nn or 'до края бокового' in nn:
        return 'H'
    if 'внутренними точками' in nn or 'внутренних точек' in nn or 'между крайними внутренними' in nn:
        return 'H'
    if 'верхней точки' in nn or 'верхней точкой' in nn:
        return 'F'
    if 'нижней точки' in nn or 'нижней точкой' in nn:
        return 'G'
    if 'по ширине' in nn:
        return 'I'
    return None


def match_singles(nn, combo):
    rules = [
        (191, lambda: 'кабин' in combo and 'оксид углерода' in nn),
        (196, lambda: 'диоксид азота' in nn),
        (201, lambda: 'оксид азота' in nn),
        (169, lambda: 'номерной' in combo and 'высот' not in nn and ('продольн' in nn or 'вверх' in nn)),
        (174, lambda: 'номерной' in combo and 'высот' not in nn and ('опорн' in nn or 'вниз' in nn)),
        (179, lambda: 'номерной' in combo and 'высот' in nn and ('нижн' in nn or 'min' in nn)),
        (184, lambda: 'номерной' in combo and 'высот' in nn and ('верхн' in nn or 'max' in nn)),
        (206, lambda: 'шум' in nn and 'разгон' in nn),
        (21, lambda: 'педал' in nn and nn.endswith(' e')),
        (26, lambda: 'педал' in nn and nn.endswith(' h')),
        (31, lambda: 'педал' in nn and nn.endswith(' j')),
        (36, lambda: 'оксид углерода' in nn),
        (41, lambda: 'углеводород' in nn),
        (46, lambda: 'звуков' in nn and 'сигнал' in nn),
        (101, lambda: 'воздуховод' in nn),
        (106, lambda: 'отоплен' in nn),
        (96, lambda: 'влажност' in nn),
        (91, lambda: 'зоне головы' in nn),
        (71, lambda: 'подвижност' in nn or ('обеспечивать' in nn and 'вентиляц' in nn)),
        (76, lambda: 'кондицион' in nn and 'охлажд' in nn),
        (81, lambda: 'кондицион' in nn and ('скорост' in nn or 'поток' in nn)),
        (66, lambda: 'вентиляц' in nn and ('выход' in nn or 'поток' in nn)),
        (86, lambda: 'температур' in nn and 'воздух' in nn),
        (51, lambda: 'длин' in nn),
        (56, lambda: 'ширин' in nn),
        (61, lambda: 'высот' in nn),
        (112, lambda: 'выступающ' in nn or 'наружной поверхности' in nn),
        (123, lambda: 'решетк' in nn and 'радиус' in nn),
        (118, lambda: 'решетк' in nn),
        (128, lambda: 'стеклоочистител' in nn),
        (133, lambda: 'бампер' in nn),
        (143, lambda: 'ручк' in nn),
        (138, lambda: 'двер' in nn),
        (149, lambda: 'обшивк' in nn),
        (158, lambda: 'светопропускан' in nn and 'ветров' in nn),
        (163, lambda: 'светопропускан' in nn),
    ]
    for row, test in rules:
        try:
            if test():
                return row
        except Exception:
            pass
    return None


def _betacf(a, b, x):
    MAXIT, EPS, FPMIN = 200, 3e-14, 1e-300
    qab, qap, qam = a + b, a + 1.0, a - 1.0
    c, d = 1.0, 1.0 - qab * x / qap
    if abs(d) < FPMIN:
        d = FPMIN
    d = 1.0 / d
    h = d
    for m in range(1, MAXIT + 1):
        m2 = 2 * m
        aa = m * (b - m) * x / ((qam + m2) * (a + m2))
        d = 1.0 + aa * d
        if abs(d) < FPMIN:
            d = FPMIN
        c = 1.0 + aa / c
        if abs(c) < FPMIN:
            c = FPMIN
        d = 1.0 / d
        h *= d * c
        aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2))
        d = 1.0 + aa * d
        if abs(d) < FPMIN:
            d = FPMIN
        c = 1.0 + aa / c
        if abs(c) < FPMIN:
            c = FPMIN
        d = 1.0 / d
        de = d * c
        h *= de
        if abs(de - 1.0) < EPS:
            break
    return h


def betai(a, b, x):
    if x <= 0.0:
        return 0.0
    if x >= 1.0:
        return 1.0
    lbeta = math.lgamma(a + b) - math.lgamma(a) - math.lgamma(b)
    bt = math.exp(lbeta + a * math.log(x) + b * math.log(1.0 - x))
    if x < (a + 1.0) / (a + b + 2.0):
        return bt * _betacf(a, b, x) / a
    return 1.0 - bt * _betacf(b, a, 1.0 - x) / b


def norm_s_inv(p):
    if p <= 0.0 or p >= 1.0:
        raise XLError('#NUM!')
    a = [-3.969683028665376e+01, 2.209460984245205e+02, -2.759285104469687e+02,
         1.383577518672690e+02, -3.066479806614716e+01, 2.506628277459239e+00]
    b = [-5.447609879822406e+01, 1.615858368580409e+02, -1.556989798598866e+02,
         6.680131188771972e+01, -1.328068155288572e+01]
    c = [-7.784894002430293e-03, -3.223964580411365e-01, -2.400758277161838e+00,
         -2.549732539343734e+00, 4.374664141464968e+00, 2.938163982698783e+00]
    d = [7.784695709041462e-03, 3.224671290700398e-01, 2.445134137142996e+00,
         3.754408661907416e+00]
    p_low, p_high = 0.02425, 1 - 0.02425
    if p < p_low:
        q = math.sqrt(-2 * math.log(p))
        x = (((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / \
            ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1)
    elif p <= p_high:
        q = p - 0.5
        r = q * q
        x = (((((a[0] * r + a[1]) * r + a[2]) * r + a[3]) * r + a[4]) * r + a[5]) * q / \
            (((((b[0] * r + b[1]) * r + b[2]) * r + b[3]) * r + b[4]) * r + 1)
    else:
        q = math.sqrt(-2 * math.log(1 - p))
        x = -(((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / \
            ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1)
    e = 0.5 * math.erfc(-x / math.sqrt(2.0)) - p
    u = e * math.sqrt(2 * math.pi) * math.exp(x * x / 2)
    return x - u / (1 + x * u / 2)


def t_inv_2t(p, df):
    p = float(p)
    if not (0.0 < p < 1.0):
        raise XLError('#NUM!')
    df = int(df)
    if df < 1:
        raise XLError('#NUM!')
    if df > 5_000_000:
        return Decimal(repr(norm_s_inv(1.0 - p / 2.0)))

    def g(t):
        return betai(df / 2.0, 0.5, df / (df + t * t)) - p
    lo, hi = 0.0, 1.0
    while g(hi) > 0 and hi < 1e12:
        hi *= 2.0
    for _ in range(300):
        mid = (lo + hi) / 2.0
        if g(mid) > 0:
            lo = mid
        else:
            hi = mid
    return Decimal(repr((lo + hi) / 2.0))


TOK_RE = re.compile(r"""
    \s+
  | (?P<num>\d+(?:\.\d*)?(?:[eE][+-]?\d+)?|\.\d+(?:[eE][+-]?\d+)?)
  | (?P<str>"[^"]*")
  | (?P<sheetq>'(?:[^']|'')+'!)
  | (?P<sheetb>[A-Za-zА-Яа-яЁё0-9_.]+!)
  | (?P<func>[A-Za-z_][A-Za-z0-9_.]*(?=\())
  | (?P<cell>\$?[A-Z]{1,3}\$?\d{1,7})
  | (?P<bool>TRUE|FALSE)
  | (?P<op>[+\-*/^(),:])
""", re.X | re.I)


def tokenize(f):
    toks = []
    pos = 0
    while pos < len(f):
        m = TOK_RE.match(f, pos)
        if not m:
            raise ValueError('tokenize: %r' % f[pos:pos + 10])
        pos = m.end()
        kind = m.lastgroup
        if kind is None:
            continue
        v = m.group()
        if kind == 'num':
            toks.append(('num', Decimal(v)))
        elif kind == 'str':
            toks.append(('str', v[1:-1]))
        elif kind == 'sheetq':
            toks.append(('sheet', v[1:-2].replace("''", "'")))
        elif kind == 'sheetb':
            toks.append(('sheet', v[:-1]))
        elif kind == 'func':
            toks.append(('func', v.upper()))
        elif kind == 'cell':
            toks.append(('cell', v.replace('$', '').upper()))
        elif kind == 'bool':
            toks.append(('bool', v.upper() == 'TRUE'))
        else:
            toks.append((v, v))
    toks.append(('end', ''))
    return toks


class Parser:
    def __init__(self, toks):
        self.t = toks
        self.i = 0

    def peek(self):
        return self.t[self.i]

    def next(self):
        tk = self.t[self.i]
        self.i += 1
        return tk

    def expect(self, op):
        if self.next()[0] != op:
            raise ValueError('expected %s' % op)

    def parse(self):
        node = self.add()
        if self.peek()[0] != 'end':
            raise ValueError('trailing tokens')
        return node

    def add(self):
        node = self.mul()
        while self.peek()[0] in ('+', '-'):
            op = self.next()[0]
            node = ('bin', op, node, self.mul())
        return node

    def mul(self):
        node = self.pow()
        while self.peek()[0] in ('*', '/'):
            op = self.next()[0]
            node = ('bin', op, node, self.pow())
        return node

    def pow(self):
        node = self.unary()
        if self.peek()[0] == '^':
            self.next()
            node = ('bin', '^', node, self.pow())
        return node

    def unary(self):
        if self.peek()[0] == '-':
            self.next()
            return ('neg', self.unary())
        if self.peek()[0] == '+':
            self.next()
            return self.unary()
        return self.primary()

    def primary(self):
        tk = self.peek()
        k = tk[0]
        if k in ('num', 'str', 'bool'):
            self.next()
            return (k, tk[1])
        if k == 'func':
            self.next()
            self.expect('(')
            args = []
            if self.peek()[0] != ')':
                args.append(self.add())
                while self.peek()[0] == ',':
                    self.next()
                    args.append(self.add())
            self.expect(')')
            return ('call', tk[1], args)
        if k == '(':
            self.next()
            node = self.add()
            self.expect(')')
            return node
        sheet = None
        if k == 'sheet':
            self.next()
            sheet = tk[1]
        if self.peek()[0] == 'cell':
            c1 = self.next()[1]
            if self.peek()[0] == ':':
                self.next()
                c2tk = self.next()
                if c2tk[0] == 'cell':
                    c2 = c2tk[1]
                else:
                    raise ValueError('bad range')
                return ('range', sheet, c1, c2)
            return ('ref', sheet, c1)
        raise ValueError('unexpected %r' % (k,))


def col_to_num(col):
    n = 0
    for ch in col:
        n = n * 26 + ord(ch) - 64
    return n


def num_to_col(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def split_ref(ref):
    m = re.match(r'([A-Z]+)(\d+)', ref)
    return m.group(1), int(m.group(2))


class Evaluator:
    def __init__(self, sheet_names):
        self.sheet_lc = {s.lower(): s for s in sheet_names}
        self.cache = {}
        self.busy = set()
        self.ast_cache = {}
        self.formulas = {}
        self.base_fn = None

    def resolve_sheet(self, s, cur):
        if s is None:
            return cur
        return self.sheet_lc.get(s.lower(), s)

    def get_value(self, sheet, ref):
        sheet = self.resolve_sheet(sheet, sheet)
        key = (sheet, ref)
        if key in self.cache:
            v = self.cache[key]
            if isinstance(v, XLError):
                raise v
            return v
        if key in self.busy:
            raise XLError('#REF!')
        if key in self.formulas.get(sheet, {}):
            self.busy.add(key)
            try:
                ast = self.ast_cache.get(key)
                if ast is None:
                    ast = Parser(tokenize(self.formulas[sheet][ref])).parse()
                    self.ast_cache[key] = ast
                v = self.ev(ast, sheet)
            except XLError as e:
                self.cache[key] = e
                raise
            finally:
                self.busy.discard(key)
            self.cache[key] = v
            return v
        v = self.base_fn(sheet, ref)
        if isinstance(v, XLError):
            raise v
        return v

    def ev(self, node, sheet):
        k = node[0]
        if k == 'num':
            return node[1]
        if k == 'str':
            return node[1]
        if k == 'bool':
            return Decimal(1) if node[1] else Decimal(0)
        if k == 'ref':
            return self.get_value(node[1] if node[1] else sheet, node[2])
        if k == 'range':
            sh = self.resolve_sheet(node[1], sheet)
            c1, r1 = split_ref(node[2])
            c2, r2 = split_ref(node[3])
            vals = []
            for rr in range(min(r1, r2), max(r1, r2) + 1):
                for cc in range(min(col_to_num(c1), col_to_num(c2)),
                                max(col_to_num(c1), col_to_num(c2)) + 1):
                    vals.append(self.get_value(sh, '%s%d' % (num_to_col(cc), rr)))
            return vals
        if k == 'neg':
            return -self.num(self.ev(node[1], sheet))
        if k == 'bin':
            op = node[1]
            if op == '+':
                return self.num(self.ev(node[2], sheet)) + self.num(self.ev(node[3], sheet))
            if op == '-':
                return self.num(self.ev(node[2], sheet)) - self.num(self.ev(node[3], sheet))
            if op == '*':
                return self.num(self.ev(node[2], sheet)) * self.num(self.ev(node[3], sheet))
            if op == '/':
                d = self.num(self.ev(node[3], sheet))
                if d == 0:
                    raise XLError('#DIV/0!')
                return self.num(self.ev(node[2], sheet)) / d
            if op == '^':
                b = self.num(self.ev(node[2], sheet))
                e = self.num(self.ev(node[3], sheet))
                if e == e.to_integral_value() and abs(int(e)) < 1000:
                    return b ** int(e)
                return Decimal(repr(float(b) ** float(e)))
        if k == 'call':
            return self.call(node[1], node[2], sheet)
        raise ValueError('bad node')

    @staticmethod
    def num(v):
        if isinstance(v, XLError):
            raise v
        if isinstance(v, Decimal):
            return v
        if isinstance(v, bool):
            return Decimal(1) if v else Decimal(0)
        if isinstance(v, str):
            try:
                return Decimal(v)
            except Exception:
                raise XLError('#VALUE!')
        if isinstance(v, list):
            raise XLError('#VALUE!')
        return Decimal(0)

    def flatten(self, args, sheet, stdeva=False):
        out = []
        for a in args:
            v = self.ev(a, sheet)
            if isinstance(v, list):
                for x in v:
                    if isinstance(x, XLError):
                        raise x
                    if isinstance(x, Decimal):
                        out.append(x)
                    elif isinstance(x, bool):
                        out.append(Decimal(1) if x else Decimal(0))
                    elif isinstance(x, str):
                        if stdeva:
                            out.append(Decimal(0))
            else:
                out.append(self.num(v))
        return out

    def call(self, name, args, sheet):
        name = name.upper()
        if name.startswith('_XLFN.'):
            name = name[6:]
        if name == 'SQRT':
            v = self.num(self.ev(args[0], sheet))
            if v < 0:
                raise XLError('#NUM!')
            return v.sqrt()
        if name == 'AVERAGE':
            vals = self.flatten(args, sheet)
            if not vals:
                raise XLError('#DIV/0!')
            return sum(vals) / Decimal(len(vals))
        if name == 'STDEVA':
            vals = self.flatten(args, sheet, stdeva=True)
            if len(vals) < 2:
                raise XLError('#DIV/0!')
            mean = sum(vals) / Decimal(len(vals))
            var = sum((x - mean) ** 2 for x in vals) / Decimal(len(vals) - 1)
            return var.sqrt()
        if name == 'T.INV.2T':
            return t_inv_2t(self.num(self.ev(args[0], sheet)),
                            self.num(self.ev(args[1], sheet)))
        if name == 'NORM.DIST':
            x = float(self.num(self.ev(args[0], sheet)))
            mu = float(self.num(self.ev(args[1], sheet)))
            sd = float(self.num(self.ev(args[2], sheet)))
            if sd <= 0:
                raise XLError('#NUM!')
            cum = self.ev(args[3], sheet)
            cum = cum if isinstance(cum, bool) else (self.num(cum) != 0)
            z = (x - mu) / sd
            if cum:
                return Decimal(repr(0.5 * (1.0 + math.erf(z / math.sqrt(2.0)))))
            return Decimal(repr(math.exp(-0.5 * z * z) / (sd * math.sqrt(2.0 * math.pi))))
        raise XLError('#NAME?')


CELL_RE_TMPL = r'<c r="%s"((?: [^>]*?)?)(?:/>|>(.*?)</c>)'


def fmt_num(d):
    if isinstance(d, Decimal):
        if d == d.to_integral_value():
            return str(int(d))
        return str(d)
    return repr(d)


def build_cell(ref, attrs, kind, value):
    s_m = re.search(r' s="(\d+)"', attrs)
    s_attr = ' s="%s"' % s_m.group(1) if s_m else ''
    if kind == 'clear':
        return '<c r="%s"%s/>' % (ref, s_attr)
    if kind == 'e':
        return '<c r="%s"%s t="e"><v>%s</v></c>' % (ref, s_attr, value)
    if kind == 'str':
        return '<c r="%s"%s t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % \
               (ref, s_attr, escape(str(value)))
    return '<c r="%s"%s><v>%s</v></c>' % (ref, s_attr, fmt_num(Decimal(value)))


def set_cell(xml, ref, kind, value=None):
    pat = re.compile(CELL_RE_TMPL % re.escape(ref), re.S)
    m = pat.search(xml)
    if m:
        new = build_cell(ref, m.group(1) or '', kind, value)
        return xml[:m.start()] + new + xml[m.end():]
    row_m = re.search(r'(\d+)$', ref)
    rnum = int(row_m.group(1))
    col = ref[:len(ref) - len(row_m.group(1))]
    cell = build_cell(ref, '', kind, value)
    row_pat = re.compile(r'<row r="%d"(?![0-9])[^>]*>(.*?)</row>' % rnum, re.S)
    rm = row_pat.search(xml)
    target_col = col_to_num(col)
    if rm:
        inner = rm.group(1)
        insert_at = len(inner)
        for cm in re.finditer(r'<c r="([A-Z]+)%d"' % rnum, inner):
            if col_to_num(cm.group(1)) > target_col:
                insert_at = cm.start()
                break
        abs_pos = rm.start(1) + insert_at
        return xml[:abs_pos] + cell + xml[abs_pos:]
    prev = None
    for rm2 in re.finditer(r'<row r="(\d+)"(?![0-9])', xml):
        if int(rm2.group(1)) < rnum:
            prev = rm2
    if prev is None:
        dm = re.search(r'<sheetData>', xml)
        return xml[:dm.end()] + '<row r="%d">%s</row>' % (rnum, cell) + xml[dm.end():]
    end_m = re.search(r'</row>', xml[prev.end():])
    abs_pos = prev.end() + end_m.end()
    return xml[:abs_pos] + '<row r="%d">%s</row>' % (rnum, cell) + xml[abs_pos:]


def parse_shared(zf):
    try:
        data = zf.read('xl/sharedStrings.xml').decode('utf-8')
    except KeyError:
        return []
    out = []
    for si in re.findall(r'<si>(.*?)</si>', data, re.S):
        txt = ''.join(re.findall(r'<t[^>]*>(.*?)</t>', si, re.S))
        out.append(unescape(txt))
    return out


def parse_sheet_cells(xml, shared):
    values, formulas = {}, {}
    for m in re.finditer(r'<c r="([A-Z]+\d+)"((?: [^>]*?)?)(?:/>|>(.*?)</c>)', xml, re.S):
        ref, attrs, content = m.group(1), m.group(2) or '', m.group(3) or ''
        t_m = re.search(r' t="(\w+)"', attrs)
        t = t_m.group(1) if t_m else None
        f_m = re.search(r'<f[^>]*>(.*?)</f>', content, re.S)
        v_m = re.search(r'<v>(.*?)</v>', content, re.S)
        if f_m:
            formulas[ref] = unescape(f_m.group(1))
            if v_m:
                raw = v_m.group(1)
                if t == 'str':
                    values[ref] = unescape(raw)
                elif t == 'e':
                    values[ref] = XLError(raw)
                else:
                    try:
                        values[ref] = Decimal(raw)
                    except Exception:
                        values[ref] = Decimal(0)
            continue
        if t == 's' and v_m:
            try:
                values[ref] = shared[int(v_m.group(1))]
            except Exception:
                pass
        elif t == 'inlineStr':
            t_m2 = re.search(r'<t[^>]*>(.*?)</t>', content, re.S)
            if t_m2:
                values[ref] = unescape(t_m2.group(1))
        elif v_m:
            if t == 'e':
                values[ref] = XLError(v_m.group(1))
            else:
                try:
                    values[ref] = Decimal(v_m.group(1))
                except Exception:
                    values[ref] = Decimal(0)
    return values, formulas


def gen_triple(samples, v, p):
    v = Decimal(v)
    if samples and all(s is not None for s in samples):
        samp = [Decimal(s) for s in samples]
        samp_f = [Fraction(s) for s in samp]
        mean_f = sum(samp_f) / 3
        v_f = Fraction(v)
        shift = v_f - mean_f
        cand = [s + shift for s in samp_f]
        max_s = max(n_dec(s) for s in samp)
        scale = 10 ** max_s
        exact = all((c * scale).denominator == 1 for c in cand)
        if exact:
            out = [Decimal((c * scale).numerator) / Decimal(scale) for c in cand]
            if sum(out) / 3 == v and all(o >= 0 for o in out):
                return out, False
        rounded_mean = excel_round(dec_fr(mean_f, p + 10), p)
        shift_d = v - rounded_mean
        out = [s + shift_d for s in samp]
        if any(o < 0 for o in out):
            return [v, v, v], False
        return out, True
    step = Decimal(1).scaleb(-max(p, 1))
    out = [v + step, v, v - step]
    if any(o < 0 for o in out):
        return [v, v, v], False
    return out, False


def build(pdf_bytes, tpl_bytes):
    meta, params = parse_pdf(pdf_bytes)

    zf = zipfile.ZipFile(io.BytesIO(tpl_bytes))
    wb_xml = zf.read('xl/workbook.xml').decode('utf-8')
    rels_xml = zf.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    rid2target = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels_xml))
    sheet_map = {}
    for name, rid in re.findall(r'<sheet name="([^"]+)" sheetId="\d+" r:id="(rId\d+)"/>', wb_xml):
        sheet_map[name] = 'xl/' + rid2target[rid].lstrip('/')
    if JOURNAL not in sheet_map:
        zf.close()
        raise ValueError('В шаблоне нет листа «Рабочий журнал».')

    shared = parse_shared(zf)
    sheet_xml = {name: zf.read(path).decode('utf-8') for name, path in sheet_map.items()}
    base_values, formulas = {}, {}
    for name in sheet_map:
        vals, forms = parse_sheet_cells(sheet_xml[name], shared)
        base_values[name] = vals
        formulas[name] = forms
    zf.close()

    wbv = openpyxl.load_workbook(io.BytesIO(tpl_bytes), read_only=True, data_only=True)
    if JOURNAL not in wbv.sheetnames:
        wbv.close()
        raise ValueError('В шаблоне нет листа «Рабочий журнал».')
    wsv = wbv[JOURNAL]
    sample_single = {}
    for r in SINGLE_ROWS:
        sample_single[r] = [wsv.cell(row=r + i, column=6).value for i in range(3)]
    sample_light = {}
    for key, r in LIGHT_ROWS.items():
        for colL, cnum in COLS.items():
            sample_light[(key, colL)] = [wsv.cell(row=r + i, column=cnum).value for i in range(3)]
    wbv.close()

    warnings, not_transferred, nflag = [], [], []
    chosen = {}

    def add_value(target, v, p, pname):
        if target not in chosen:
            chosen[target] = {'vals': [], 'p': p, 'name': pname}
        chosen[target]['vals'].append((Fraction(v), p))
        chosen[target]['p'] = max(chosen[target]['p'], p)

    def pick_value(prm, res):
        mt, rt = dec_of_tok(res['meas_tok']), dec_of_tok(res['res_tok'])
        val = mt
        if mt != rt:
            if req_fits(prm['req'], mt):
                val = mt
            elif req_fits(prm['req'], rt):
                val = rt
            else:
                val = mt
            warnings.append('измерено %s, в итоге %s — взято %s (%s)'
                            % (mt, rt, val, prm['name']))
        return val, prec_of_tok(res['res_tok'])

    last_light = None
    for prm in params:
        nn = norm_txt(prm['name'])
        sn = norm_txt(prm['section'])
        combo = nn + ' | ' + sn
        has_res = bool(prm['results'])
        if not has_res:
            continue
        if 'разбрызгиван' in nn:
            not_transferred.append(prm['name'])
            continue
        if 'шум' in nn and 'разгон' not in nn:
            not_transferred.append(prm['name'])
            continue

        pos = detect_pos(nn)
        lkey = detect_light(nn)
        if lkey:
            last_light = lkey
        if pos or lkey:
            if pos and (lkey or last_light):
                use = lkey or last_light
                for res in prm['results']:
                    val, p = pick_value(prm, res)
                    add_value(('L', use, pos), val, p, prm['name'])
                    if res['flag'] == 'N':
                        nflag.append('%s: %s' % (prm['name'], val))
            continue

        row = match_singles(nn, combo)
        if row is None:
            not_transferred.append(prm['name'])
            continue
        vals = []
        for res in prm['results']:
            val, p = pick_value(prm, res)
            vals.append((val, p))
            if res['flag'] == 'N':
                nflag.append('%s: %s' % (prm['name'], val))
        if not vals:
            continue
        p = max(pp for _, pp in vals)
        if len(vals) == 1:
            v = vals[0][0]
        else:
            v = dec_fr(sum(Fraction(x) for x, _ in vals) / len(vals), p + 10)
        add_value(('S', row), v, p, prm['name'])

    new_vals = {}
    filled_info, approx_count = [], 0
    filled_units = set()
    for target, info in chosen.items():
        vals_f = [x for x, _ in info['vals']]
        p = info['p']
        if len(vals_f) == 1:
            v = dec_fr(vals_f[0], p + 10)
        else:
            v = dec_fr(sum(vals_f) / len(vals_f), p + 10)
        if target[0] == 'S':
            row = target[1]
            triple, approx = gen_triple(sample_single.get(row), v, p)
            for i, tv in enumerate(triple):
                new_vals[('F', row + i)] = tv
            filled_units.add(('S', row))
            approx_count += 1 if approx else 0
            filled_info.append((ROW_LABEL.get(row, str(row)), v, approx))
        else:
            key, colL = target[1], target[2]
            row = LIGHT_ROWS[key]
            triple, approx = gen_triple(sample_light.get((key, colL)), v, p)
            for i, tv in enumerate(triple):
                new_vals[(colL, row + i)] = tv
            filled_units.add(('L', key, colL))
            approx_count += 1 if approx else 0
            filled_info.append(('%s, кол. %s' % (LIGHT_LABEL[key], colL), v, approx))

    cleared = []
    for row in SINGLE_ROWS:
        if ('S', row) in filled_units:
            continue
        if any(s is not None for s in sample_single.get(row, [])):
            cleared.append(ROW_LABEL.get(row, str(row)))
        for i in range(3):
            new_vals[('F', row + i)] = 'CLEAR'
    for key, row in LIGHT_ROWS.items():
        for colL in COLS:
            if ('L', key, colL) in filled_units:
                continue
            if any(s is not None for s in sample_light.get((key, colL), [])):
                cleared.append('%s, кол. %s' % (LIGHT_LABEL[key], colL))
            for i in range(3):
                new_vals[(colL, row + i)] = 'CLEAR'

    jvals = {}
    for (colL, r), tv in new_vals.items():
        jvals['%s%d' % (colL, r)] = tv
    if meta.get('marka'):
        jvals['E11'] = meta['marka']
    if meta.get('model'):
        jvals['E12'] = meta['model']
    if meta.get('vin'):
        jvals['E15'] = meta['vin']
    if meta.get('date'):
        d = meta['date']
        jvals['C2'] = '     Дата проведения испытаний: «%02d» %s %d г.      ' % \
                      (d.day, MONTHS_GEN[d.month - 1], d.year)
    for ref, key in (('D6', 'temp'), ('D7', 'hum'), ('D8', 'press'), ('D9', 'wind')):
        if meta.get(key) is not None:
            jvals[ref] = meta[key]
    if meta['date'] is None:
        warnings.append('Дата испытаний в PDF не найдена — строка даты не заполнена.')

    overrides = {}
    for ref, tv in jvals.items():
        overrides[(JOURNAL, ref)] = None if tv == 'CLEAR' else tv
    calc_meta = {}
    if meta.get('marka') or meta.get('model'):
        calc_meta['D2'] = (meta.get('marka') or '') + ' ' + (meta.get('model') or '')
    if meta.get('date'):
        calc_meta['D4'] = Decimal((meta['date'] - date(1899, 12, 30)).days)
    for name in sheet_map:
        if name != JOURNAL:
            for ref, tv in calc_meta.items():
                overrides[(name, ref)] = tv

    merged = {}
    for name in sheet_map:
        merged[name] = dict(base_values.get(name, {}))
    for (name, ref), tv in overrides.items():
        if tv is None:
            merged[name].pop(ref, None)
        else:
            merged[name][ref] = tv

    ctrl_errors = []
    for target, info in chosen.items():
        vals_f = [x for x, _ in info['vals']]
        p = info['p']
        v = dec_fr(vals_f[0], p + 10) if len(vals_f) == 1 \
            else dec_fr(sum(vals_f) / len(vals_f), p + 10)
        if target[0] == 'S':
            row = target[1]
            trip = [merged[JOURNAL].get('F%d' % (row + i), Decimal(0)) for i in range(3)]
        else:
            row = LIGHT_ROWS[target[1]]
            trip = [merged[JOURNAL].get('%s%d' % (target[2], row + i), Decimal(0))
                    for i in range(3)]
        trip = [x if isinstance(x, Decimal) else Decimal(0) for x in trip]
        mean = sum(trip) / 3
        if excel_round(mean, p) != v:
            ctrl_errors.append('%s: среднее %s ≠ протоколу %s' % (info['name'], mean, v))

    evaluator = Evaluator(list(sheet_map.keys()))
    evaluator.formulas = formulas

    def base_fn(sheet, ref):
        if (sheet, ref) in overrides:
            tv = overrides[(sheet, ref)]
            if tv is None:
                return Decimal(0)
            if isinstance(tv, Decimal):
                return tv
            try:
                return Decimal(str(tv))
            except Exception:
                return tv if isinstance(tv, str) else Decimal(0)
        return merged.get(sheet, {}).get(ref, Decimal(0))

    evaluator.base_fn = base_fn

    cache_results = {}
    calc_warnings = 0
    for name in sheet_map:
        for ref, f in formulas.get(name, {}).items():
            try:
                v = evaluator.get_value(name, ref)
                cache_results[(name, ref)] = ('str' if isinstance(v, str) else 'n', v)
            except XLError as e:
                cache_results[(name, ref)] = ('e', e.code)
            except Exception:
                calc_warnings += 1
    if calc_warnings:
        warnings.append('Не удалось пересчитать %d формул — оставлен прежний кэш.' % calc_warnings)

    edited = {}
    jxml = sheet_xml[JOURNAL]
    for ref, tv in jvals.items():
        if tv == 'CLEAR':
            jxml = set_cell(jxml, ref, 'clear')
        elif isinstance(tv, Decimal):
            jxml = set_cell(jxml, ref, 'n', tv)
        else:
            jxml = set_cell(jxml, ref, 'str', tv)
    for (name, ref), (kind, v) in cache_results.items():
        if name == JOURNAL:
            jxml = set_cell(jxml, ref, kind, v)
    edited[sheet_map[JOURNAL]] = jxml.encode('utf-8')

    for name in sheet_map:
        if name == JOURNAL:
            continue
        xml = sheet_xml[name]
        if 'D2' in calc_meta:
            xml = set_cell(xml, 'D2', 'str', calc_meta['D2'])
        if 'D4' in calc_meta:
            xml = set_cell(xml, 'D4', 'n', calc_meta['D4'])
        for (nm, ref), (kind, v) in cache_results.items():
            if nm == name:
                xml = set_cell(xml, ref, kind, v)
        edited[sheet_map[name]] = xml.encode('utf-8')

    if 'fullCalcOnLoad' not in wb_xml:
        wb_xml = re.sub(r'<calcPr ', '<calcPr fullCalcOnLoad="1" ', wb_xml, count=1)
    edited['xl/workbook.xml'] = wb_xml.encode('utf-8')

    out = io.BytesIO()
    zin = zipfile.ZipFile(io.BytesIO(tpl_bytes))
    zout = zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED)
    for info in zin.infolist():
        data = edited.get(info.filename, zin.read(info.filename))
        zout.writestr(info, data)
    zout.close()
    zin.close()

    report = {
        'meta': meta,
        'filled': filled_info,
        'cleared': cleared,
        'approx': approx_count,
        'warnings': warnings,
        'not_transferred': not_transferred,
        'nflag': nflag,
        'ctrl_errors': ctrl_errors,
    }
    return out.getvalue(), report


def main():
    st.set_page_config(page_title='Рабочий журнал испытаний', layout='centered')
    st.title('Заполнение рабочего журнала по PDF-протоколу')
    st.caption('Загрузите PDF-протокол испытаний — приложение заполнит шаблон template.xlsx '
               'и вернёт готовый файл с пересчитанными формулами.')

    pdf = st.file_uploader('PDF-протокол испытаний', type=['pdf'])

    tpl_path = Path(__file__).resolve().parent / 'template.xlsx'
    tpl_bytes = None
    if tpl_path.exists():
        tpl_bytes = tpl_path.read_bytes()
        st.success('Шаблон template.xlsx найден рядом с приложением.')
    else:
        tpl_up = st.file_uploader('Шаблон template.xlsx (если не найден рядом с app.py)',
                                  type=['xlsx'])
        if tpl_up is not None:
            tpl_bytes = tpl_up.read()

    if st.button('Заполнить журнал', disabled=pdf is None or tpl_bytes is None):
        with st.spinner('Разбор протокола и заполнение журнала…'):
            try:
                data, report = build(pdf.read(), tpl_bytes)
            except Exception as e:
                st.session_state.pop('res', None)
                st.error(str(e))
            else:
                st.session_state['res'] = (data, report, pdf.name)

    res = st.session_state.get('res')
    if res:
        data, report, pdf_name = res
        meta = report['meta']
        st.subheader('Реквизиты протокола')
        c1, c2 = st.columns(2)
        c1.write('**№ протокола:** %s' % (meta.get('protocol') or '—'))
        c1.write('**Марка:** %s' % (meta.get('marka') or '—'))
        c1.write('**Модель:** %s' % (meta.get('model') or '—'))
        c2.write('**VIN:** %s' % (meta.get('vin') or '—'))
        c2.write('**Дата:** %s' % (meta['date'].strftime('%d.%m.%Y') if meta.get('date') else '—'))
        c2.write('**t °C / влажность % / ветер м/с / давление кПа:** %s / %s / %s / %s' % (
            meta.get('temp') or '—', meta.get('hum') or '—',
            meta.get('wind') or '—', meta.get('press') or '—'))

        c1, c2, c3 = st.columns(3)
        c1.metric('Заполнено блоков', len(report['filled']))
        c2.metric('Очищено блоков', len(report['cleared']))
        c3.metric('Округлено («≈»)', report['approx'])

        if report['ctrl_errors']:
            st.error('Контроль средних: несовпадения — ' + '; '.join(report['ctrl_errors']))
        if report['nflag']:
            st.error('Несоответствия (N):\n' + '\n'.join('• ' + x for x in report['nflag']))
        if report['warnings']:
            st.warning('\n'.join('• ' + w for w in report['warnings']))

        with st.expander('Заполненные блоки (%d)' % len(report['filled'])):
            st.write('\n'.join('• %s = %s%s' % (n, v, ' ≈' if a else '')
                               for n, v, a in report['filled']))
        with st.expander('Очищено (%d)' % len(report['cleared'])):
            st.write('\n'.join('• ' + x for x in report['cleared']) or '—')
        with st.expander('Не перенесено в журнал (%d)' % len(report['not_transferred'])):
            st.write('\n'.join('• ' + x for x in report['not_transferred']) or '—')

        vin = meta.get('vin') or 'journal'
        st.download_button('Скачать заполненный журнал (.xlsx)', data=data,
                           file_name='zhurnal_%s.xlsx' % vin,
                           mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


if __name__ == '__main__':
    main()
