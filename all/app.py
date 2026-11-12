# -*- coding: utf-8 -*-
"""
Общее приложение: все журналы испытаний ТС на одном сайте (M1, M2, M3, N1, N2, N3, O4).
Сверху — маленькие вкладки категорий в одну строку, ниже — приложение этой категории (код и шаблоны
берутся из папок категорий: app.py + template.xlsx, m2/, m2_m3/, n1/, n2/, n3/, o4/ — ничего не дублируется).

Streamlit Cloud: Main file path  all/app.py
Ссылка на категорию: …/?cat=N3
"""
import importlib.util
import sys
from pathlib import Path

import streamlit as st

ROOT = Path(__file__).resolve().parent.parent

# код, название, файл приложения
CATS = [
    ("M1", "Легковые автомобили", "app.py"),
    ("M2", "Автобусы малой вместимости", "m2/app.py"),
    ("M3", "Автобусы", "m2_m3/app.py"),
    ("N1", "Грузовые автомобили", "n1/app.py"),
    ("N2", "Грузовые автомобили", "n2/app.py"),
    ("N3", "Грузовые автомобили", "n3/app.py"),
    ("O4", "Прицепы и полуприцепы", "o4/app.py"),
]
BY_CODE = {c[0]: c for c in CATS}
# состояние, которое нельзя переносить между категориями: результаты и заменённый шаблон
RESET_KEYS = ("batch", "tpl", "tpl2")


def load_app(code):
    """Модуль приложения категории (загружается один раз за сессию сервера)."""
    name = "journal_app_%s" % code.lower()
    mod = sys.modules.get(name)
    if mod is None:
        spec = importlib.util.spec_from_file_location(name, ROOT / BY_CODE[code][2])
        mod = importlib.util.module_from_spec(spec)
        sys.modules[name] = mod
        spec.loader.exec_module(mod)
    return mod


def on_tab():
    """Смена вкладки: сбросить результаты и заменённый шаблон прошлой категории."""
    pick = st.session_state.get("tab")
    if pick is None:                      # повторное нажатие снимает выбор — оставляем прежнюю категорию
        st.session_state["tab"] = st.session_state.get("cat", "M1")
        return
    if pick != st.session_state.get("cat"):
        for k in RESET_KEYS:
            st.session_state.pop(k, None)
        st.session_state["cat"] = pick


CSS = """
<style>
[data-testid="stHeader"]{background:transparent}
[data-testid="stToolbar"],#MainMenu,footer{visibility:hidden}
.block-container{max-width:820px;padding-top:2.2rem;padding-bottom:3rem}
.stButton>button,.stDownloadButton>button{min-height:2.9rem;font-size:1rem;border-radius:10px}
.top{display:flex;align-items:baseline;justify-content:space-between;gap:.8rem;margin-bottom:.7rem}
.top .n{font-weight:800;font-size:1.25rem;letter-spacing:-.01em}
.top .c{font-size:.9rem;opacity:.6}
/* маленькие вкладки категорий в одну строку */
.st-key-tabs button{min-height:2.2rem!important;height:2.2rem;padding:0 .3rem!important}
.st-key-tabs button p{font-size:.9rem;font-weight:700;overflow:visible;text-overflow:clip}
.st-key-tabs button[aria-checked="true"]{background:#111!important;border-color:#111!important}
.st-key-tabs button[aria-checked="true"] *{color:#fff!important}
[data-testid="stBaseButton-primary"]{background:#111!important;border-color:#111!important;color:#fff!important}
[data-testid="stBaseButton-primary"] *{color:#fff!important}
[data-testid="stBaseButton-primary"]:disabled{background:transparent!important;border-color:rgba(127,127,127,.3)!important}
[data-testid="stBaseButton-primary"]:disabled *{color:rgba(127,127,127,.8)!important}
@media (max-width:640px){.top .c{display:none}}
.cur{margin:.55rem 0 1.1rem;font-size:.92rem;opacity:.75}
</style>
"""


def main():
    st.set_page_config(page_title="Журналы испытаний ТС", page_icon="📋", layout="centered")
    st.markdown(CSS, unsafe_allow_html=True)

    # выбранная категория: из ссылки (?cat=N3) или прошлый выбор
    if st.session_state.get("cat") not in BY_CODE:
        q = str(st.query_params.get("cat", "")).upper()
        st.session_state["cat"] = q if q in BY_CODE else "M1"
        st.session_state["tab"] = st.session_state["cat"]

    st.markdown('<div class="top"><span class="n">📋 Журналы испытаний ТС</span>'
                '<span class="c">по PDF-протоколам</span></div>', unsafe_allow_html=True)
    with st.container(key="tabs"):
        st.segmented_control("Категория", [c[0] for c in CATS], key="tab", on_change=on_tab,
                             label_visibility="collapsed", width="stretch")
    cur = st.session_state["cat"]
    if st.query_params.get("cat") != cur:
        st.query_params["cat"] = cur
    code, title, _ = BY_CODE[cur]
    st.markdown('<div class="cur"><b>%s</b> · %s</div>' % (code, title), unsafe_allow_html=True)

    load_app(code).main(embedded=True)

    st.write("")
    with st.expander("Установить на телефон или компьютер"):
        st.markdown(
            "- **iPhone:** открыть сайт в Safari → «Поделиться» → «На экран Домой».\n"
            "- **Android:** Chrome → меню ⋮ → «Добавить на главный экран».\n"
            "- **Компьютер:** Chrome/Edge → значок установки в адресной строке.\n\n"
            "Ссылка сразу на категорию: в конце адреса `?cat=N3` (M1, M2, M3, N1, N2, N3, O4).")


main()
