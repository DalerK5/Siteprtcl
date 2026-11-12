# -*- coding: utf-8 -*-
"""
Общее приложение: все журналы испытаний ТС на одном сайте (M1, M2, M3, N1, N2, N3, O4).
Сверху — выбор категории, ниже — приложение этой категории (код и шаблоны берутся из папок
категорий: app.py + template.xlsx, m2/, m2_m3/, n1/, n2/, n3/, o4/ — ничего не дублируется).

Три варианта оформления (для выбора): ?style=a — светлый минимализм, ?style=b — строгий корпоративный,
?style=c — мягкий тёплый. По умолчанию — STYLE.

Streamlit Cloud: Main file path  all/app.py
Ссылка на категорию: …/?cat=N3
"""
import importlib.util
import sys
from pathlib import Path

import streamlit as st

ROOT = Path(__file__).resolve().parent.parent
STYLE = "a"

# код, значок, название, файл приложения
CATS = [
    ("M1", "🚗", "Легковые автомобили", "app.py"),
    ("M2", "🚐", "Автобусы малой вместимости", "m2/app.py"),
    ("M3", "🚌", "Автобусы", "m2_m3/app.py"),
    ("N1", "🛻", "Грузовые автомобили", "n1/app.py"),
    ("N2", "🚚", "Грузовые автомобили", "n2/app.py"),
    ("N3", "🚛", "Грузовые автомобили", "n3/app.py"),
    ("O4", "🛞", "Прицепы и полуприцепы", "o4/app.py"),
]
BY_CODE = {c[0]: c for c in CATS}
# состояние, которое нельзя переносить между категориями: результаты и заменённый шаблон
RESET_KEYS = ("batch", "tpl", "tpl2")


def load_app(code):
    """Модуль приложения категории (загружается один раз за сессию сервера)."""
    name = "journal_app_%s" % code.lower()
    mod = sys.modules.get(name)
    if mod is None:
        spec = importlib.util.spec_from_file_location(name, ROOT / BY_CODE[code][3])
        mod = importlib.util.module_from_spec(spec)
        sys.modules[name] = mod
        spec.loader.exec_module(mod)
    return mod


def choose(code):
    """Выбор категории: сбросить результаты и заменённый шаблон прошлой категории."""
    if code != st.session_state.get("cat"):
        for k in RESET_KEYS:
            st.session_state.pop(k, None)
        st.session_state["cat"] = code


# ---------------------------------------------------------------- общий CSS (сетка плиток категорий)
BASE = """
[data-testid="stHeader"]{background:transparent}
[data-testid="stToolbar"],#MainMenu,footer{visibility:hidden}
.st-key-cats [data-testid="stHorizontalBlock"]{flex-wrap:wrap;gap:.6rem}
.st-key-cats [data-testid="stColumn"]{flex:1 1 calc(50% - .6rem)!important;min-width:calc(50% - .6rem)!important;width:auto!important}
@media (min-width:700px){.st-key-cats [data-testid="stColumn"]{flex:1 1 calc(@COLS@ - .6rem)!important;min-width:calc(@COLS@ - .6rem)!important}}
.st-key-cats button{width:100%;justify-content:flex-start;text-align:left;transition:all .15s ease}
.st-key-cats button p{text-align:left;margin:0;line-height:1.25;white-space:normal}
.st-key-cats button [data-testid="stMarkdownContainer"]{display:flex;flex-direction:column;gap:.15rem;align-items:flex-start}
.st-key-cats button p + p{font-size:.85rem;opacity:.75;font-weight:500}
.st-key-cats button{padding:.6rem .85rem;height:100%}
.st-key-cats button > div{width:100%;justify-content:flex-start!important;text-align:left}
.st-key-cats [data-testid="stColumn"] [data-testid="stVerticalBlock"],.st-key-cats [data-testid="stElementContainer"],
.st-key-cats [data-testid="stButton"]{height:100%}
[data-testid="stIconMaterial"]{font-family:'Material Symbols Rounded'!important}
.stButton>button,.stDownloadButton>button{min-height:3rem;font-size:1rem}
"""

# ---------------------------------------------------------------- А: светлый минимализм
CSS_A = """
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
.stApp{background:#ffffff}
.stApp,.stApp p,.stApp li,.stApp label,.stApp button,.stApp input,.stApp h1,.stApp summary{font-family:'Inter',system-ui,sans-serif!important}
.block-container{max-width:860px;padding-top:2.4rem;padding-bottom:4rem}
.hd{display:flex;align-items:center;gap:.6rem;font-weight:700;font-size:.95rem;color:#111;margin-bottom:2.4rem}
.hd .dot{width:.7rem;height:.7rem;border-radius:50%;background:#2563eb}
.stApp h1.ttl{font-size:clamp(2rem,6vw,3.2rem)!important;padding:0!important;font-weight:800;letter-spacing:-.035em;line-height:1.05;color:#0b0b0c;margin:0}
.sub{color:#6b7280;font-size:1.05rem;line-height:1.55;margin:.9rem 0 2rem;max-width:36rem}
.lbl{font-size:.78rem;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:#9ca3af;margin:1.6rem 0 .6rem}
.st-key-cats button{min-height:3.6rem;border-radius:14px;background:#fff;border:1px solid #e5e7eb;color:#111;font-weight:500}
.st-key-cats button:hover{border-color:#2563eb;color:#111;background:#f8faff}
.st-key-cats [data-testid="stBaseButton-primary"]{background:#eff4ff!important;border:1.5px solid #2563eb!important;color:#1d4ed8!important}
.st-key-cats [data-testid="stBaseButton-primary"] *{color:#1d4ed8!important}
[data-testid="stBaseButton-primary"]{background:#111!important;border:1px solid #111!important;color:#fff!important;border-radius:12px}
[data-testid="stBaseButton-primary"] *{color:#fff!important}
[data-testid="stBaseButton-primary"]:disabled{background:#f3f4f6!important;border-color:#e5e7eb!important}
[data-testid="stBaseButton-primary"]:disabled *{color:#9ca3af!important}
[data-testid="stBaseButton-secondary"]{border-radius:12px}
[data-testid="stFileUploaderDropzone"]{background:#fafafa;border:1.5px dashed #d1d5db;border-radius:14px}
[data-testid="stExpander"] details{border-radius:12px;border-color:#e5e7eb}
[data-testid="stAlert"]{border-radius:12px}
.sec{display:flex;align-items:center;gap:.6rem;margin:2.6rem 0 1rem;padding-top:1.6rem;border-top:1px solid #f0f0f0}
.sec .t{font-size:1.25rem;font-weight:700;color:#111}
.sec .b{font-size:.8rem;font-weight:700;color:#1d4ed8;background:#eff4ff;padding:.15rem .55rem;border-radius:999px}
"""

# ---------------------------------------------------------------- Б: строгий корпоративный
CSS_B = """
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700&display=swap');
.stApp{background:#eef1f5}
.stApp,.stApp p,.stApp li,.stApp label,.stApp button,.stApp input,.stApp h1,.stApp summary{font-family:'IBM Plex Sans',system-ui,sans-serif!important}
.block-container{max-width:920px;padding-top:0;padding-bottom:4rem}
.bar{margin:0 -100vw 0;padding:1rem 100vw 1rem;background:#0f2a4a;color:#fff;display:flex;justify-content:space-between;
     align-items:center;font-weight:600;letter-spacing:.02em}
.bar .r{opacity:.7;font-weight:400;font-size:.9rem}
@media (max-width:640px){.bar .r{display:none}}
.bar .logo{display:flex;align-items:center;gap:.55rem}
.bar .logo i{font-style:normal;background:#f5b700;color:#0f2a4a;font-weight:800;padding:.05rem .4rem;border-radius:4px}
.st-key-top{background:#fff;border:1px solid #dde3ea;border-radius:10px;padding:1.4rem 1.4rem 1.3rem;margin:1.4rem 0 1rem;
       box-shadow:0 1px 2px rgba(15,42,74,.05)}
.stApp h1.ttl{font-size:clamp(1.6rem,4.8vw,2.2rem)!important;padding:0!important;font-weight:700;color:#0f2a4a;margin:0;line-height:1.15}
.sub{color:#5b6b7f;font-size:1rem;margin:.5rem 0 0}
.lbl{font-size:.82rem;font-weight:600;color:#5b6b7f;margin:1.2rem 0 .5rem;text-transform:uppercase;letter-spacing:.06em}
.st-key-cats button{min-height:3.3rem;border-radius:8px;background:#fff;border:1px solid #cfd8e3;color:#0f2a4a;font-weight:500}
.st-key-cats button:hover{border-color:#0f2a4a;color:#0f2a4a}
.st-key-cats [data-testid="stBaseButton-primary"]{background:#0f2a4a!important;border-color:#0f2a4a!important;color:#fff!important}
.st-key-cats [data-testid="stBaseButton-primary"] *{color:#fff!important}
.st-key-body{background:#fff;border:1px solid #dde3ea;border-radius:10px;padding:1.2rem 1.3rem 1.4rem}
[data-testid="stBaseButton-primary"]{background:#0f2a4a!important;border:1px solid #0f2a4a!important;color:#fff!important;border-radius:8px}
[data-testid="stBaseButton-primary"] *{color:#fff!important}
[data-testid="stBaseButton-primary"]:disabled{background:#e6eaf0!important;border-color:#dde3ea!important}
[data-testid="stBaseButton-primary"]:disabled *{color:#8b97a6!important}
[data-testid="stFileUploaderDropzone"]{background:#f6f8fb;border:1px dashed #b7c3d1;border-radius:8px}
.sec{display:flex;align-items:center;gap:.7rem;margin:.2rem 0 1rem}
.sec .t{font-size:1.15rem;font-weight:700;color:#0f2a4a}
.sec .b{font-size:.8rem;font-weight:700;color:#0f2a4a;background:#f5b700;padding:.12rem .5rem;border-radius:4px}
"""

# ---------------------------------------------------------------- В: мягкий тёплый
CSS_C = """
@import url('https://fonts.googleapis.com/css2?family=Unbounded:wght@500;700&family=Nunito+Sans:wght@400;600;700;800&display=swap');
.stApp{background:#faf6ef}
.stApp,.stApp p,.stApp li,.stApp label,.stApp button,.stApp input,.stApp summary{font-family:'Nunito Sans',system-ui,sans-serif!important}
.block-container{max-width:880px;padding-top:2rem;padding-bottom:4rem}
.hero{background:#fff;border-radius:28px;padding:1.6rem 1.5rem;box-shadow:0 18px 40px -26px rgba(120,72,30,.45);
      display:flex;gap:1.2rem;align-items:center;margin-bottom:1.4rem}
@media (max-width:640px){.hero{flex-direction:column;align-items:flex-start;gap:.8rem}.hero .em{font-size:2.4rem}}
.hero .em{font-size:3.2rem;line-height:1;background:#ffe9dc;border-radius:22px;padding:.7rem .8rem}
.stApp h1.ttl{font-family:'Unbounded',sans-serif!important;font-size:clamp(1.3rem,5.2vw,1.9rem)!important;padding:0!important;font-weight:700;color:#2b211a;
     line-height:1.2;margin:0}
.sub{color:#7a6a5c;font-size:1rem;margin:.45rem 0 0}
.lbl{font-family:'Unbounded',sans-serif;font-size:.95rem;font-weight:500;color:#2b211a;margin:1.4rem 0 .7rem}
.st-key-cats button{min-height:4rem;border-radius:20px;background:#fff;border:0;color:#2b211a;font-weight:700;
       box-shadow:0 6px 18px -12px rgba(120,72,30,.5)}
.st-key-cats button:hover{transform:translateY(-2px);color:#e2602f}
.st-key-cats [data-testid="stBaseButton-primary"]{background:#e2602f!important;color:#fff!important}
.st-key-cats [data-testid="stBaseButton-primary"] *{color:#fff!important}
.st-key-body{background:#fff;border-radius:28px;padding:1.3rem 1.4rem 1.5rem;box-shadow:0 18px 40px -28px rgba(120,72,30,.45)}
[data-testid="stBaseButton-primary"]{background:#e2602f!important;border:0!important;color:#fff!important;border-radius:16px}
[data-testid="stBaseButton-primary"] *{color:#fff!important}
[data-testid="stBaseButton-primary"]:disabled{background:#f1e8de!important}
[data-testid="stBaseButton-primary"]:disabled *{color:#b3a394!important}
[data-testid="stBaseButton-secondary"]{border-radius:16px}
[data-testid="stFileUploaderDropzone"]{background:#fff8f2;border:2px dashed #f0b896;border-radius:20px}
[data-testid="stExpander"] details{border-radius:16px;border-color:#efe4d8}
[data-testid="stAlert"]{border-radius:16px}
.sec{display:flex;align-items:center;gap:.6rem;margin:1.6rem 0 .9rem}
.sec .t{font-family:'Unbounded',sans-serif;font-size:1.05rem;font-weight:700;color:#2b211a}
.sec .b{font-size:.8rem;font-weight:800;color:#fff;background:#e2602f;padding:.15rem .6rem;border-radius:999px}
"""

STYLES = {"a": (CSS_A, "25%"), "b": (CSS_B, "25%"), "c": (CSS_C, "25%")}


def header(style):
    if style == "a":
        st.markdown('<div class="hd"><span class="dot"></span>Журналы испытаний</div>'
                    '<h1 class="ttl">Заполнение журналов<br>по протоколам испытаний</h1>'
                    '<p class="sub">Выберите категорию ТС и загрузите PDF-протоколы — журналы заполнятся '
                    'по шаблону автоматически.</p><div class="lbl">Категория ТС</div>', unsafe_allow_html=True)
    elif style == "b":
        st.markdown('<div class="bar"><div class="logo"><i>ЖИ</i>Журналы испытаний ТС</div>'
                    '<div class="r">M1 · M2 · M3 · N1 · N2 · N3 · O4</div></div>'
                    , unsafe_allow_html=True)

    else:
        st.markdown('<div class="hero"><div class="em">📋</div><div><h1 class="ttl">Журналы испытаний ТС</h1>'
                    '<p class="sub">Выберите категорию и загрузите PDF-протоколы — остальное сделаем сами.</p>'
                    '</div></div><div class="lbl">Категория</div>', unsafe_allow_html=True)


def header_b():
    st.markdown('<h1 class="ttl">Заполнение рабочих журналов</h1>'
                '<p class="sub">Выберите категорию ТС и загрузите PDF-протоколы испытаний.</p>'
                '<div class="lbl">Категория ТС</div>', unsafe_allow_html=True)


def main():
    st.set_page_config(page_title="Журналы испытаний ТС", page_icon="📋", layout="centered")
    style = str(st.query_params.get("style", STYLE)).lower()
    style = style if style in STYLES else STYLE
    css, cols = STYLES[style]
    st.markdown("<style>%s%s</style>" % (BASE.replace("@COLS@", cols), css), unsafe_allow_html=True)

    # выбранная категория: из ссылки (?cat=N3) или прошлый выбор
    if st.session_state.get("cat") not in BY_CODE:
        q = str(st.query_params.get("cat", "")).upper()
        st.session_state["cat"] = q if q in BY_CODE else "M1"
    cur = st.session_state["cat"]

    header(style)
    with st.container(key="top"):
        if style == "b":
            header_b()
        cats(cur)
    if st.query_params.get("cat") != cur:
        st.query_params["cat"] = cur

    code, icon, title, _ = BY_CODE[cur]
    app = load_app(code)
    with st.container(key="body"):
        st.markdown('<div class="sec"><span class="b">%s</span><span class="t">%s</span></div>' % (code, title),
                    unsafe_allow_html=True)
        app.main(embedded=True)
        st.write("")
        with st.expander("Установить на телефон или компьютер"):
            st.markdown(
                "- **iPhone:** открыть сайт в Safari → «Поделиться» → «На экран Домой».\n"
                "- **Android:** Chrome → меню ⋮ → «Добавить на главный экран».\n"
                "- **Компьютер:** Chrome/Edge → значок установки в адресной строке.\n\n"
                "Ссылка сразу на категорию: в конце адреса `?cat=N3` (M1, M2, M3, N1, N2, N3, O4).")


def cats(cur):
    with st.container(key="cats"):
        cols_ = st.columns(len(CATS))
        for col, (code, icon, title, _) in zip(cols_, CATS):
            col.button("%s **%s**\n\n%s" % (icon, code, title), key="cat_" + code, on_click=choose, args=(code,),
                       type="primary" if code == cur else "secondary", use_container_width=True)


main()
