# -*- coding: utf-8 -*-
"""
Общее приложение: все журналы испытаний ТС на одном сайте (M1, M2, M3, N1, N2, N3, O4).
Сверху — выбор категории, ниже — приложение этой категории (код и шаблоны берутся из папок
категорий: app.py + template.xlsx, m2/, m2_m3/, n1/, n2/, n3/, o4/ — ничего не дублируется).

Streamlit Cloud: Main file path  all/app.py
Ссылка на категорию: …/?cat=N3
"""
import importlib.util
import sys
from pathlib import Path

import streamlit as st

ROOT = Path(__file__).resolve().parent.parent

# код, значок, название, описание, файл приложения
CATS = [
    ("M1", "🚗", "Легковые", "Легковые автомобили категории M1", "app.py"),
    ("M2", "🚐", "Микроавтобусы", "Автобусы малой вместимости категории M2 (до 5 т)", "m2/app.py"),
    ("M3", "🚌", "Автобусы", "Автобусы категории M3 (городской, межгородской)", "m2_m3/app.py"),
    ("N1", "🛻", "Грузовые до 3,5 т", "Грузовые автомобили категории N1 (до 3,5 т)", "n1/app.py"),
    ("N2", "🚚", "Грузовые 3,5–12 т", "Грузовые автомобили категории N2 (от 3,5 до 12 т)", "n2/app.py"),
    ("N3", "🚛", "Грузовые свыше 12 т", "Грузовые автомобили категории N3 (свыше 12 т)", "n3/app.py"),
    ("O4", "🛞", "Прицепы свыше 10 т", "Прицепы и полуприцепы категории O4 (свыше 10 т)", "o4/app.py"),
]
BY_CODE = {c[0]: c for c in CATS}
# состояние, которое нельзя переносить между категориями: результаты и заменённый шаблон
RESET_KEYS = ("batch", "tpl", "tpl2")


def load_app(code):
    """Модуль приложения категории (загружается один раз за сессию сервера)."""
    name = "journal_app_%s" % code.lower()
    mod = sys.modules.get(name)
    if mod is None:
        spec = importlib.util.spec_from_file_location(name, ROOT / BY_CODE[code][4])
        mod = importlib.util.module_from_spec(spec)
        sys.modules[name] = mod
        spec.loader.exec_module(mod)
    return mod


def on_pick():
    """Смена категории: сбросить результаты и заменённый шаблон прошлой категории."""
    pick = st.session_state.get("cat_pick")
    if pick is None:                      # повторное нажатие снимает выбор — оставляем прежнюю категорию
        st.session_state["cat_pick"] = st.session_state.get("cat", "M1")
        return
    if pick != st.session_state.get("cat"):
        for k in RESET_KEYS:
            st.session_state.pop(k, None)
        st.session_state["cat"] = pick


CSS = """
<style>
.block-container{padding-top:3.2rem;padding-bottom:2.5rem;max-width:760px}
.stButton>button,.stDownloadButton>button{min-height:3rem;font-size:1.05rem;border-radius:12px}
.hero{border-radius:18px;padding:1.1rem 1.2rem;margin:.2rem 0 1rem;
      background:linear-gradient(135deg,rgba(79,110,247,.16),rgba(16,185,129,.12));
      border:1px solid rgba(127,127,127,.18)}
.hero h1{font-size:clamp(1.3rem,5.4vw,1.75rem);line-height:1.2;margin:0;padding:0}
.hero p{margin:.35rem 0 0;opacity:.8;font-size:.98rem}
.cat-card{display:flex;gap:.9rem;align-items:center;border-radius:16px;padding:.85rem 1rem;margin:.4rem 0 1.1rem;
          border:1px solid rgba(127,127,127,.22);background:rgba(127,127,127,.06)}
.cat-card .ic{font-size:2.3rem;line-height:1;flex:0 0 auto}
.cat-card .ic svg{width:3.2rem;height:auto}
.cat-card .t{font-weight:700;font-size:1.15rem}
.cat-card .d{opacity:.75;font-size:.92rem;margin-top:.1rem}
.badge{display:inline-block;font-size:.78rem;font-weight:700;padding:.1rem .5rem;border-radius:999px;
       margin-right:.35rem;background:rgba(79,110,247,.18)}
button[data-variant="pills"]{min-height:2.6rem;font-size:1rem;padding:0 .9rem}
[data-testid="stBaseButton-primary"]{background:#4f6ef7;border-color:#4f6ef7;color:#fff}
[data-testid="stBaseButton-primary"]:hover{background:#3f5ee6;border-color:#3f5ee6;color:#fff}
button[data-variant="pills"][aria-checked="true"]{background:rgba(79,110,247,.16)!important;border-color:#4f6ef7!important}
button[data-variant="pills"][aria-checked="true"] *{color:#4f6ef7!important}
</style>
"""


def main():
    st.set_page_config(page_title="Журналы испытаний ТС", page_icon="📋", layout="centered")
    st.markdown(CSS, unsafe_allow_html=True)
    st.markdown(
        '<div class="hero"><h1>📋 Журналы испытаний ТС</h1>'
        "<p>Выберите категорию, загрузите PDF-протоколы — получите заполненные журналы по шаблону.</p></div>",
        unsafe_allow_html=True)

    # выбранная категория: из ссылки (?cat=N3) или прошлый выбор
    cur = st.session_state.get("cat")
    if cur not in BY_CODE:
        q = str(st.query_params.get("cat", "")).upper()
        cur = q if q in BY_CODE else "M1"
        st.session_state["cat"] = cur
        st.session_state["cat_pick"] = cur

    st.pills("Категория ТС", [c[0] for c in CATS], key="cat_pick", on_change=on_pick,
             format_func=lambda c: "%s %s" % (BY_CODE[c][1], c), selection_mode="single")
    cur = st.session_state["cat"]
    if st.query_params.get("cat") != cur:
        st.query_params["cat"] = cur

    code, icon, title, desc, _ = BY_CODE[cur]
    app = load_app(code)
    pic = getattr(app, "TRAILER_SVG", None) if code == "O4" else None
    st.markdown(
        '<div class="cat-card"><div class="ic">%s</div><div><div class="t"><span class="badge">%s</span>%s</div>'
        '<div class="d">%s</div></div></div>' % (pic or icon, code, title, desc),
        unsafe_allow_html=True)

    app.main(embedded=True)

    st.divider()
    with st.expander("📱 Как установить на телефон или компьютер"):
        st.markdown(
            "- **iPhone:** открыть сайт в Safari → «Поделиться» → «На экран Домой».\n"
            "- **Android:** Chrome → меню ⋮ → «Добавить на главный экран» (или «Установить приложение»).\n"
            "- **Компьютер:** Chrome/Edge → значок установки в адресной строке (или меню → «Установить …»).\n\n"
            "Ссылка сразу на категорию: добавьте в конец адреса `?cat=N3` (M1, M2, M3, N1, N2, N3, O4).")


main()
