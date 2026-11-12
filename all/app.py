# -*- coding: utf-8 -*-
"""
Общее приложение: все журналы испытаний ТС на одном сайте (M1, M2, M3, N1, N2, N3, O4).
Сверху — выбор категории (плитки), ниже — приложение этой категории (код и шаблоны берутся из папок
категорий: app.py + template.xlsx, m2/, m2_m3/, n1/, n2/, n3/, o4/ — ничего не дублируется).

Streamlit Cloud: Main file path  all/app.py
Ссылка на категорию: …/?cat=N3
"""
import importlib.util
import sys
from pathlib import Path

import streamlit as st

ROOT = Path(__file__).resolve().parent.parent

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


CSS = """
<style>
:root{--a1:#4f46e5;--a2:#7c3aed}
[data-testid="stHeader"]{background:transparent}
[data-testid="stToolbar"],#MainMenu,footer{visibility:hidden}
[data-testid="stAppViewContainer"]{
  background:radial-gradient(1100px 520px at 0% -10%,rgba(79,70,229,.10),transparent 60%),
             radial-gradient(900px 480px at 110% 0%,rgba(6,182,212,.09),transparent 60%)}
.block-container{padding-top:2.2rem;padding-bottom:3rem;max-width:780px}

/* шапка */
.hero{position:relative;overflow:hidden;border-radius:24px;padding:1.5rem 1.4rem 1.35rem;margin:0 0 1.2rem;color:#fff;
      background:linear-gradient(135deg,var(--a1) 0%,var(--a2) 55%,#db2777 130%);
      box-shadow:0 14px 34px -14px rgba(79,70,229,.55)}
.hero:after{content:"";position:absolute;right:-60px;top:-60px;width:220px;height:220px;border-radius:50%;
      background:rgba(255,255,255,.10)}
.hero .k{font-size:.78rem;letter-spacing:.14em;text-transform:uppercase;opacity:.8;font-weight:700}
.hero h1{color:#fff;font-size:clamp(1.45rem,6vw,2rem);line-height:1.15;margin:.25rem 0 0;padding:0;font-weight:800}
.hero p{margin:.5rem 0 0;opacity:.9;font-size:.98rem;max-width:34rem}
.hero .row{margin-top:.9rem;font-size:1.45rem;letter-spacing:.35rem;opacity:.95}

/* шаги */
.step{display:flex;align-items:center;gap:.55rem;margin:1rem 0 .5rem;font-weight:700;font-size:1.02rem}
.step .n{display:inline-flex;align-items:center;justify-content:center;width:1.7rem;height:1.7rem;border-radius:50%;
         color:#fff;font-size:.9rem;background:linear-gradient(135deg,var(--a1),var(--a2))}

/* плитки категорий */
.st-key-tiles [data-testid="stHorizontalBlock"]{flex-wrap:wrap;gap:.55rem}
.st-key-tiles [data-testid="stColumn"]{flex:1 1 calc(25% - .55rem)!important;min-width:calc(25% - .55rem)!important;width:auto!important}
@media (min-width:700px){.st-key-tiles [data-testid="stColumn"]{flex:1 1 0!important;min-width:0!important}}
.st-key-tiles button{width:100%;min-height:4.1rem;border-radius:16px;font-weight:700;font-size:1.02rem;
       border:1px solid rgba(127,127,127,.22);background:rgba(127,127,127,.06);transition:all .15s ease}
.st-key-tiles button:hover{border-color:var(--a1);transform:translateY(-1px)}

/* выбранная категория */
.cat{display:flex;align-items:center;gap:.8rem;border-radius:18px;padding:.8rem 1rem;margin:.9rem 0 .2rem;
     border:1px solid rgba(127,127,127,.2);background:rgba(127,127,127,.05)}
.cat .ic{font-size:2.1rem;line-height:1}
.cat .ic svg{width:3.4rem;height:auto}
.cat .code{display:inline-block;font-weight:800;font-size:.85rem;padding:.12rem .55rem;border-radius:999px;color:#fff;
     background:linear-gradient(135deg,var(--a1),var(--a2));margin-right:.45rem}
.cat .t{font-weight:700;font-size:1.12rem}

/* элементы приложения категории */
.stButton>button,.stDownloadButton>button{min-height:3.1rem;font-size:1.04rem;border-radius:14px}
[data-testid="stBaseButton-primary"]{border:0!important;color:#fff!important;
     background:linear-gradient(135deg,var(--a1),var(--a2))!important;box-shadow:0 10px 22px -12px rgba(79,70,229,.9)}
[data-testid="stBaseButton-primary"]:hover{filter:brightness(1.08)}
[data-testid="stBaseButton-primary"]:disabled{background:rgba(127,127,127,.18)!important;box-shadow:none;
     color:rgba(127,127,127,.9)!important}
[data-testid="stFileUploaderDropzone"]{border:2px dashed rgba(79,70,229,.45);border-radius:18px;background:rgba(79,70,229,.05)}
[data-testid="stExpander"] details{border-radius:14px}
[data-testid="stAlert"]{border-radius:14px}
</style>
"""


def main():
    st.set_page_config(page_title="Журналы испытаний ТС", page_icon="📋", layout="centered")
    st.markdown(CSS, unsafe_allow_html=True)

    # выбранная категория: из ссылки (?cat=N3) или прошлый выбор
    if st.session_state.get("cat") not in BY_CODE:
        q = str(st.query_params.get("cat", "")).upper()
        st.session_state["cat"] = q if q in BY_CODE else "M1"

    st.markdown(
        '<div class="hero"><div class="k">Сертификационные испытания</div>'
        "<h1>Журналы испытаний ТС</h1>"
        "<p>Выберите категорию, загрузите PDF-протоколы — журналы заполнятся по шаблону автоматически.</p>"
        '<div class="row">%s</div></div>' % "".join(c[1] for c in CATS),
        unsafe_allow_html=True)

    st.markdown('<div class="step"><span class="n">1</span>Категория ТС</div>', unsafe_allow_html=True)
    cur = st.session_state["cat"]
    with st.container(key="tiles"):
        cols = st.columns(len(CATS))
        for col, (code, icon, _, _) in zip(cols, CATS):
            col.button("%s %s" % (icon, code), key="tile_" + code, on_click=choose, args=(code,),
                       type="primary" if code == cur else "secondary", use_container_width=True)
    if st.query_params.get("cat") != cur:
        st.query_params["cat"] = cur

    code, icon, title, _ = BY_CODE[cur]
    app = load_app(code)
    pic = getattr(app, "TRAILER_SVG", None) if code == "O4" else None
    st.markdown('<div class="cat"><div class="ic">%s</div><div class="t"><span class="code">%s</span>%s</div></div>'
                % (pic or icon, code, title), unsafe_allow_html=True)

    st.markdown('<div class="step"><span class="n">2</span>Протоколы и журналы</div>', unsafe_allow_html=True)
    app.main(embedded=True)

    st.write("")
    with st.expander("📱 Установить на телефон или компьютер"):
        st.markdown(
            "- **iPhone:** открыть сайт в Safari → «Поделиться» → «На экран Домой».\n"
            "- **Android:** Chrome → меню ⋮ → «Добавить на главный экран».\n"
            "- **Компьютер:** Chrome/Edge → значок установки в адресной строке.\n\n"
            "Ссылка сразу на категорию: в конце адреса `?cat=N3` (M1, M2, M3, N1, N2, N3, O4).")


main()
