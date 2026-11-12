# -*- coding: utf-8 -*-
"""
Общее приложение: все журналы испытаний ТС на одном сайте (M1, M2, M3, N1, N2, N3, O4).
Сверху — выбор категории, ниже — приложение этой категории (код и шаблоны берутся из папок
категорий: app.py + template.xlsx, m2/, m2_m3/, n1/, n2/, n3/, o4/ — ничего не дублируется).
Стиль: чёрный фон с «зерном», белый крупный заголовок, лаймовый акцент, тонкие линии, без скруглений
(тема — в .streamlit/config.toml, остальное — CSS ниже).

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
ORDER = [c[0] for c in CATS]
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


def choose(code):
    """Выбор категории: сбросить результаты и заменённый шаблон прошлой категории."""
    if code != st.session_state.get("cat"):
        for k in RESET_KEYS:
            st.session_state.pop(k, None)
        st.session_state["cat"] = code


# «зерно» фона — SVG-шум
NOISE = ("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' width='160' height='160'>"
         "<filter id='n'><feTurbulence type='fractalNoise' baseFrequency='.9' numOctaves='3' stitchTiles='stitch'/>"
         "</filter><rect width='100%' height='100%' filter='url(%23n)' opacity='.55'/></svg>")

CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Manrope:wght@400;500;600;700;800&display=swap');
:root{--lime:#C6F432;--bg:#0B0B0B;--line:rgba(255,255,255,.14);--mute:rgba(237,237,237,.62)}
html,body,.stApp,.stApp p,.stApp li,.stApp label,.stApp h1,.stApp h2,.stApp h3,.stApp button,.stApp input,.stApp textarea,
.stApp small,.stApp div[data-testid="stMarkdownContainer"],.stApp summary{font-family:'Manrope',system-ui,sans-serif!important}
[data-testid="stIconMaterial"],.material-symbols-rounded{font-family:'Material Symbols Rounded'!important}
.stApp{background:var(--bg)}
.stApp:before{content:"";position:fixed;inset:0;pointer-events:none;z-index:0;opacity:.07;background-image:url("%(noise)s")}
.bigword{position:fixed;left:-2vw;bottom:-7vw;font-weight:800;font-size:30vw;line-height:.8;letter-spacing:-.04em;
         color:rgba(255,255,255,.035);pointer-events:none;z-index:0;white-space:nowrap;user-select:none}
[data-testid="stHeader"]{background:transparent}
[data-testid="stToolbar"],#MainMenu,footer{visibility:hidden}
.block-container{position:relative;z-index:1;padding-top:1.4rem;padding-bottom:4rem;max-width:1120px}
*{border-radius:0!important}

/* верхняя строка */
.top{display:flex;align-items:center;justify-content:space-between;gap:1rem;padding:.2rem 0 1.4rem;
     border-bottom:1px solid var(--line);margin-bottom:2.2rem}
.logo{display:flex;align-items:stretch;font-weight:800;letter-spacing:.32em;font-size:.82rem;text-transform:uppercase}
.logo b{background:var(--lime);color:#0B0B0B;padding:.28rem .45rem .28rem .6rem}
.logo span{padding:.28rem .6rem;color:#fff}
.top .nav{color:var(--mute);font-size:.88rem;letter-spacing:.04em}
.top .nav i{font-style:normal;color:var(--lime)}

/* заголовок */
.hero{display:flex;gap:1.4rem}
.rail{display:flex;flex-direction:column;align-items:center;gap:.5rem;font-size:.7rem;color:var(--mute);padding-top:.5rem}
.rail .bar{width:1px;height:7.5rem;background:var(--line);position:relative}
.rail .bar:before{content:"";position:absolute;left:-1px;top:0;width:3px;height:%(fill)s%%;background:var(--lime)}
.hero h1{font-weight:800!important;font-size:clamp(2.1rem,6.2vw,4rem)!important;line-height:1.06!important;
         letter-spacing:-.02em;color:#fff;margin:0;padding:0}
.hero p{color:var(--mute);font-size:1rem;line-height:1.6;max-width:31rem;margin:1.2rem 0 0}
.cta{display:flex;align-items:center;gap:1rem;margin-top:1.8rem;font-weight:700;color:#fff;font-size:.95rem}
.cta:before{content:"";width:3rem;height:1px;background:var(--lime)}

/* список категорий справа */
.st-key-cats,.st-key-cats [data-testid="stVerticalBlock"]{gap:.15rem!important}
.st-key-cats [data-testid="stElementContainer"]{margin:0!important}
.st-key-cats button{justify-content:flex-start;background:transparent!important;border:0!important;box-shadow:none!important;
     color:#EDEDED!important;min-height:2.25rem!important;padding:.2rem .6rem;width:fit-content!important;font-weight:500;font-size:.98rem}
.st-key-cats button:hover{color:var(--lime)!important}
.st-key-cats button p{font-size:.98rem}
.st-key-cats [data-testid="stBaseButton-primary"]{background:var(--lime)!important;color:#0B0B0B!important;font-weight:700}
.st-key-cats [data-testid="stBaseButton-primary"] *,.st-key-cats button[kind="primary"] p{color:#0B0B0B!important}
.st-key-cats [data-testid="stButton"]{width:fit-content}

/* раздел */
.sec{display:flex;align-items:baseline;gap:1rem;margin:3rem 0 1rem;padding-top:1.4rem;border-top:1px solid var(--line)}
.sec .no{color:var(--lime);font-size:.8rem;font-weight:700;letter-spacing:.12em}
.sec .t{font-size:1.5rem;font-weight:800;color:#fff;letter-spacing:-.01em}
.sec .c{margin-left:auto;color:var(--mute);font-size:.9rem}
.sec .c b{color:#0B0B0B;background:var(--lime);padding:.08rem .4rem;margin-right:.45rem}
.st-key-body{max-width:780px}

/* элементы приложения категории */
.stButton>button,.stDownloadButton>button{min-height:3.1rem;font-size:1rem;font-weight:700}
[data-testid="stBaseButton-primary"]{background:var(--lime)!important;border:1px solid var(--lime)!important;color:#0B0B0B!important}
[data-testid="stBaseButton-primary"] *,button[kind="primary"] p,.stDownloadButton button[kind="primary"] p{color:#0B0B0B!important}
[data-testid="stBaseButton-primary"]:hover{filter:brightness(1.08)}
[data-testid="stBaseButton-primary"]:disabled{background:transparent!important;border:1px solid var(--line)!important}
[data-testid="stBaseButton-primary"]:disabled *{color:var(--mute)!important}
[data-testid="stBaseButton-secondary"]{background:transparent!important;border:1px solid var(--line)!important;color:#EDEDED!important}
[data-testid="stBaseButton-secondary"]:hover{border-color:var(--lime)!important;color:var(--lime)!important}
[data-testid="stFileUploaderDropzone"]{background:rgba(255,255,255,.025);border:1px dashed rgba(198,244,50,.45)}
[data-testid="stExpander"] details{border:1px solid var(--line);background:transparent}
[data-testid="stAlert"]{background:rgba(255,255,255,.03);border:1px solid var(--line)}
[data-testid="stMetricValue"]{color:var(--lime)}
hr{border-color:var(--line)!important}
@media (max-width:640px){.rail{display:none}.top .nav{display:none}.sec .c{display:none}}
</style>
"""


def main():
    st.set_page_config(page_title="Журналы испытаний ТС", page_icon="📋", layout="wide")

    # выбранная категория: из ссылки (?cat=N3) или прошлый выбор
    if st.session_state.get("cat") not in BY_CODE:
        q = str(st.query_params.get("cat", "")).upper()
        st.session_state["cat"] = q if q in BY_CODE else "M1"
    cur = st.session_state["cat"]
    idx = ORDER.index(cur) + 1

    st.markdown(CSS % {"noise": NOISE, "fill": round(idx / len(ORDER) * 100)}, unsafe_allow_html=True)
    st.markdown('<div class="bigword">ЖУРНАЛ</div>', unsafe_allow_html=True)
    st.markdown('<div class="top"><div class="logo"><b>ЖУРНАЛЫ</b><span>ИСПЫТАНИЙ</span></div>'
                '<div class="nav">Категория <i>%s</i> · M1 — O4</div></div>' % cur, unsafe_allow_html=True)

    left, right = st.columns([1.55, 1], gap="large")
    with left:
        st.markdown(
            '<div class="hero"><div class="rail">%02d<div class="bar"></div>%02d</div><div>'
            "<h1>Журналы<br>испытаний ТС</h1>"
            "<p>Автоматическое заполнение рабочих журналов по PDF-протоколам испытаний: "
            "средние значения, неопределённость и нормы — как в протоколе.</p>"
            '<div class="cta">Выберите категорию и загрузите протоколы</div></div></div>'
            % (idx, len(ORDER)), unsafe_allow_html=True)
    with right:
        with st.container(key="cats"):
            for code, title, _ in CATS:
                st.button("%s  ·  %s" % (code, title), key="cat_" + code, on_click=choose, args=(code,),
                          type="primary" if code == cur else "secondary")
    if st.query_params.get("cat") != cur:
        st.query_params["cat"] = cur

    code, title, _ = BY_CODE[cur]
    app = load_app(code)
    st.markdown('<div class="sec"><span class="no">%02d</span><span class="t">Протоколы и журналы</span>'
                '<span class="c"><b>%s</b>%s</span></div>' % (idx, code, title), unsafe_allow_html=True)
    with st.container(key="body"):
        app.main(embedded=True)
        st.write("")
        with st.expander("Установить на телефон или компьютер"):
            st.markdown(
                "- **iPhone:** открыть сайт в Safari → «Поделиться» → «На экран Домой».\n"
                "- **Android:** Chrome → меню ⋮ → «Добавить на главный экран».\n"
                "- **Компьютер:** Chrome/Edge → значок установки в адресной строке.\n\n"
                "Ссылка сразу на категорию: в конце адреса `?cat=N3` (M1, M2, M3, N1, N2, N3, O4).")


main()
