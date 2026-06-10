import os
import re
import time as time_module
from datetime import datetime
from io import BytesIO

import pandas as pd
import plotly.express as px
import streamlit as st

# === КОНСТАНТЫ ===
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE = os.path.join(BASE_DIR, "Кафе_Учет.xlsx")

STRUCTURE_COLUMNS = [
    "Основная категория", "Подкатегория", "Блюдо", "Цена (₸)", "Себестоимость (₸)"
]

DEFAULT_CATEGORIES = [
    "Напитки", "Первые блюда", "Вторые блюда", "Фастфуд", "Пицца",
    "Соусы", "Шашлыки", "Донер", "Выпечка", "Десерты",
]

# === НАСТРОЙКА СТРАНИЦЫ ===
st.set_page_config(
    page_title="Кафе Учет",
    page_icon="☕",
    layout="wide",
    initial_sidebar_state="expanded",
)


# === ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ===
def refresh_data():
    st.cache_data.clear()


def calculate_margin(price, cost):
    return (price - cost) / price * 100 if price > 0 else 0


def get_dish_price(dish_name, menu_df):
    if menu_df is None or menu_df.empty:
        return 0
    price_row = menu_df.loc[menu_df["Блюдо"] == dish_name, "Цена продажи (₸)"]
    return float(price_row.iloc[0]) if not price_row.empty else 0


def parse_dish_entries(dishes_text):
    """Разбирает строку заказа: '2x Капучино, Латте' → [(Капучино, 2), (Латте, 1)]."""
    entries = []
    for part in str(dishes_text).split(","):
        part = part.strip()
        if not part:
            continue
        match = re.match(r"^(\d+)x\s+(.+)$", part)
        if match:
            entries.append((match.group(2).strip(), int(match.group(1))))
        else:
            entries.append((part, 1))
    return entries


def calculate_food_cost(filtered_orders, menu_df):
    if menu_df is None or menu_df.empty or filtered_orders.empty:
        return 0
    cost_map = menu_df.set_index("Блюдо")["Себестоимость (₸)"].to_dict()
    total_cost = 0
    for dishes_text in filtered_orders["Блюда"].dropna():
        for dish, qty in parse_dish_entries(dishes_text):
            total_cost += cost_map.get(dish, 0) * qty
    revenue = filtered_orders["Сумма (₸)"].sum()
    return total_cost / revenue * 100 if revenue > 0 else 0


def safe_save_to_excel(df, sheet_name, mode="replace"):
    max_attempts = 3
    for attempt in range(max_attempts):
        try:
            with pd.ExcelWriter(
                EXCEL_FILE, engine="openpyxl", mode="a", if_sheet_exists=mode
            ) as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
            return True
        except PermissionError:
            if attempt < max_attempts - 1:
                time_module.sleep(1)
                continue
            st.error(
                "❌ Файл Excel заблокирован. Закройте файл 'Кафе_Учет.xlsx' и попробуйте снова."
            )
            return False
        except FileNotFoundError:
            try:
                with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl", mode="w") as writer:
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                return True
            except Exception as e:
                st.error(f"❌ Ошибка сохранения: {e}")
                return False
        except Exception as e:
            st.error(f"❌ Ошибка сохранения: {e}")
            return False
    return False


@st.cache_data
def load_categories():
    if not os.path.exists(EXCEL_FILE):
        return DEFAULT_CATEGORIES.copy()
    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name="КАТЕГОРИИ")
        categories = df["Категория"].dropna().astype(str).tolist()
        return categories if categories else DEFAULT_CATEGORIES.copy()
    except Exception:
        return DEFAULT_CATEGORIES.copy()


def save_categories(categories):
    df = pd.DataFrame({"Категория": categories})
    if safe_save_to_excel(df, "КАТЕГОРИИ", mode="replace"):
        refresh_data()


@st.cache_data
def load_subcategories():
    if not os.path.exists(EXCEL_FILE):
        return None
    try:
        return pd.read_excel(EXCEL_FILE, sheet_name="ПОДКАТЕГОРИИ")
    except Exception:
        return None


@st.cache_data
def load_full_structure():
    if not os.path.exists(EXCEL_FILE):
        return pd.DataFrame(columns=STRUCTURE_COLUMNS)
    try:
        return pd.read_excel(EXCEL_FILE, sheet_name="КАТЕГОРИИ_ПОЛНАЯ")
    except Exception:
        return pd.DataFrame(columns=STRUCTURE_COLUMNS)


def sync_subcategories_sheet(structure_df):
    """Синхронизирует лист ПОДКАТЕГОРИИ для раздела заказов."""
    dishes = structure_df[
        structure_df["Блюдо"].notna() & (structure_df["Блюдо"].astype(str).str.strip() != "")
    ][["Основная категория", "Подкатегория", "Блюдо"]]
    safe_save_to_excel(dishes, "ПОДКАТЕГОРИИ", mode="replace")


def sync_categories_sheet(structure_df):
    """Синхронизирует лист КАТЕГОРИИ из основных категорий структуры."""
    categories = (
        structure_df["Основная категория"]
        .dropna()
        .astype(str)
        .str.strip()
    )
    categories = sorted(categories[categories != ""].unique().tolist())
    if not categories:
        categories = DEFAULT_CATEGORIES.copy()
    safe_save_to_excel(pd.DataFrame({"Категория": categories}), "КАТЕГОРИИ", mode="replace")


def get_dishes_in_scope(structure_df, main_cat, subcat=None):
    """Возвращает названия блюд в категории или подкатегории."""
    mask = structure_df["Основная категория"] == main_cat
    if subcat is not None:
        mask &= structure_df["Подкатегория"] == subcat
    dishes = structure_df.loc[mask, "Блюдо"].dropna().astype(str).str.strip()
    return dishes[dishes != ""].tolist()


def remove_menu_dishes(dish_names):
    """Удаляет несколько блюд из листа МЕНЮ."""
    if not dish_names:
        return
    menu_df = pd.read_excel(EXCEL_FILE, sheet_name="МЕНЮ")
    updated = menu_df[~menu_df["Блюдо"].isin(dish_names)]
    safe_save_to_excel(updated, "МЕНЮ", mode="replace")


def sync_menu_with_structure(structure_df):
    """Приводит лист МЕНЮ в соответствие со структурой категорий."""
    valid = structure_df[
        structure_df["Блюдо"].notna() & (structure_df["Блюдо"].astype(str).str.strip() != "")
    ].copy()
    if valid.empty:
        safe_save_to_excel(
            pd.DataFrame(columns=[
                "Блюдо", "Категория", "Ингредиенты",
                "Себестоимость (₸)", "Цена продажи (₸)", "Маржинальность (%)",
            ]),
            "МЕНЮ",
            mode="replace",
        )
        return

    dish_info = valid.set_index("Блюдо")
    menu_df = pd.read_excel(EXCEL_FILE, sheet_name="МЕНЮ")
    menu_df = menu_df[menu_df["Блюдо"].isin(dish_info.index)]

    for dish in dish_info.index:
        price = dish_info.at[dish, "Цена (₸)"] or 0
        cost = dish_info.at[dish, "Себестоимость (₸)"] or 0
        category = dish_info.at[dish, "Основная категория"]
        if dish in menu_df["Блюдо"].values:
            idx = menu_df[menu_df["Блюдо"] == dish].index[0]
            menu_df.at[idx, "Категория"] = category
            menu_df.at[idx, "Цена продажи (₸)"] = price
            menu_df.at[idx, "Себестоимость (₸)"] = cost
            menu_df.at[idx, "Маржинальность (%)"] = calculate_margin(price, cost)
        else:
            menu_df = pd.concat([menu_df, pd.DataFrame([{
                "Блюдо": dish,
                "Категория": category,
                "Ингредиенты": "",
                "Себестоимость (₸)": cost,
                "Цена продажи (₸)": price,
                "Маржинальность (%)": calculate_margin(price, cost),
            }])], ignore_index=True)

    safe_save_to_excel(menu_df, "МЕНЮ", mode="replace")


def sync_all_sheets(structure_df):
    """Синхронизирует ПОДКАТЕГОРИИ, КАТЕГОРИИ и МЕНЮ со структурой."""
    sync_subcategories_sheet(structure_df)
    sync_categories_sheet(structure_df)
    sync_menu_with_structure(structure_df)


def save_full_structure(structure_df):
    if safe_save_to_excel(structure_df, "КАТЕГОРИИ_ПОЛНАЯ", mode="replace"):
        sync_all_sheets(structure_df)
        refresh_data()


def filter_orders_by_period(orders, period):
    today = pd.Timestamp.now().date()
    if period == "Сегодня":
        return orders[orders["Дата"].dt.date == today]
    if period == "Неделя":
        return orders[orders["Дата"].dt.date >= today - pd.Timedelta(days=7)]
    if period == "Месяц":
        return orders[orders["Дата"].dt.date >= today - pd.Timedelta(days=30)]
    return orders


def build_popular_dishes_df(filtered_orders):
    dish_counts = {}
    for dishes_text in filtered_orders["Блюда"].dropna():
        for dish, qty in parse_dish_entries(dishes_text):
            dish_counts[dish] = dish_counts.get(dish, 0) + qty
    if not dish_counts:
        return pd.DataFrame(columns=["Блюдо", "Количество"])
    counts = pd.Series(dish_counts).sort_values(ascending=False).reset_index()
    counts.columns = ["Блюдо", "Количество"]
    return counts


def build_report_excel(filtered_orders, menu_df, period_label):
    """Формирует Excel-файл отчёта в памяти."""
    revenue = filtered_orders["Сумма (₸)"].sum()
    num_orders = len(filtered_orders)
    avg_check = revenue / num_orders if num_orders > 0 else 0
    food_cost_pct = calculate_food_cost(filtered_orders, menu_df)
    total_cost = revenue * food_cost_pct / 100 if revenue > 0 else 0

    summary = pd.DataFrame({
        "Показатель": [
            "Период", "Дата формирования", "Заказов", "Выручка (₸)",
            "Средний чек (₸)", "Себестоимость (₸)", "Фудкост (%)", "Прибыль (₸)",
        ],
        "Значение": [
            period_label,
            datetime.now().strftime("%d.%m.%Y %H:%M"),
            num_orders,
            round(revenue, 0),
            round(avg_check, 0),
            round(total_cost, 0),
            round(food_cost_pct, 1),
            round(revenue - total_cost, 0),
        ],
    })

    orders_export = filtered_orders.copy()
    if "Дата" in orders_export.columns:
        orders_export["Дата"] = orders_export["Дата"].dt.strftime("%d.%m.%Y")

    daily = pd.DataFrame(columns=["Дата", "Выручка (₸)", "Заказов"])
    if not filtered_orders.empty:
        daily = (
            filtered_orders.groupby(filtered_orders["Дата"].dt.date)["Сумма (₸)"]
            .agg(["sum", "count"])
            .reset_index()
        )
        daily.columns = ["Дата", "Выручка (₸)", "Заказов"]

    popular = build_popular_dishes_df(filtered_orders)

    payment_stats = pd.DataFrame(columns=["Оплата", "Заказов", "Сумма (₸)"])
    if "Оплата" in filtered_orders.columns and not filtered_orders.empty:
        payment_stats = (
            filtered_orders.groupby("Оплата")
            .agg(**{"Заказов": ("Сумма (₸)", "count"), "Сумма (₸)": ("Сумма (₸)", "sum")})
            .reset_index()
        )

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="Сводка", index=False)
        orders_export.to_excel(writer, sheet_name="Заказы", index=False)
        daily.to_excel(writer, sheet_name="По дням", index=False)
        popular.to_excel(writer, sheet_name="Популярные блюда", index=False)
        payment_stats.to_excel(writer, sheet_name="По оплате", index=False)
    buffer.seek(0)
    return buffer


def upsert_menu_dish(dish_name, category, ingredients, cost, price):
    margin = calculate_margin(price, cost)
    menu_row = pd.DataFrame([{
        "Блюдо": dish_name,
        "Категория": category,
        "Ингредиенты": ingredients,
        "Себестоимость (₸)": cost,
        "Цена продажи (₸)": price,
        "Маржинальность (%)": margin,
    }])
    existing_menu = pd.read_excel(EXCEL_FILE, sheet_name="МЕНЮ")
    mask = existing_menu["Блюдо"] == dish_name
    if mask.any():
        idx = existing_menu[mask].index[0]
        for col, val in menu_row.iloc[0].items():
            existing_menu.at[idx, col] = val
        updated_menu = existing_menu
    else:
        updated_menu = pd.concat([existing_menu, menu_row], ignore_index=True)
    safe_save_to_excel(updated_menu, "МЕНЮ", mode="replace")


def remove_menu_dish(dish_name):
    menu_df = pd.read_excel(EXCEL_FILE, sheet_name="МЕНЮ")
    updated = menu_df[menu_df["Блюдо"] != dish_name]
    safe_save_to_excel(updated, "МЕНЮ", mode="replace")


# === КАСТОМНЫЙ CSS ===
st.markdown("""
<style>
    .stApp { background: linear-gradient(135deg, #2d2b2a 0%, #1a1a1a 100%); }
    div[data-testid="stMetric"] {
        background: linear-gradient(135deg, #3d3a38, #2c2a28);
        border-radius: 20px; padding: 20px; box-shadow: 0 4px 20px rgba(0,0,0,0.3);
        border: 1px solid #d4a373; transition: transform 0.3s ease;
    }
    div[data-testid="stMetric"]:hover { transform: translateY(-5px); }
    div[data-testid="stMetric"] label { color: #d4a373 !important; font-size: 16px !important; }
    div[data-testid="stMetric"] div { color: white !important; font-size: 32px !important; font-weight: bold !important; }
    h1, h2, h3, h4 { color: #d4a373 !important; font-family: 'Segoe UI', 'Arial', sans-serif; font-weight: 600 !important; }
    h1 { border-bottom: 3px solid #d4a373; display: inline-block; padding-bottom: 10px; }
    [data-testid="stSidebar"] { background: #1a1a1a; border-right: 2px solid #d4a373; }
    [data-testid="stSidebar"] * { color: #e0e0e0 !important; }
    .stButton > button {
        background: linear-gradient(135deg, #d4a373, #b5835a);
        color: white; border: none; border-radius: 30px; padding: 12px 28px;
        font-weight: bold; transition: all 0.3s ease;
    }
    .stButton > button:hover { transform: scale(1.02); background: linear-gradient(135deg, #e0b184, #c4936a); }
    .stTextInput > div > div > input, .stNumberInput > div > div > input,
    .stTextArea > div > div > textarea, .stSelectbox > div > div > select {
        background-color: #2a2a2a; border-radius: 15px; border: 1px solid #d4a373;
        font-size: 16px; color: white !important;
    }
    .stDataFrame { background: #2a2a2a; border-radius: 15px; border: 1px solid #d4a373; }
    .stDataFrame th { background: #d4a373 !important; color: #1a1a1a !important; }
    .stDataFrame td { color: white !important; }
    .dish-card {
        background: linear-gradient(135deg, #3d3a38, #2c2a28);
        border-radius: 20px; padding: 15px; margin: 10px 0;
        border-left: 5px solid #d4a373; color: white;
    }
    [data-testid="stMetricValue"] { color: white !important; font-size: 36px !important; font-weight: bold !important; }
</style>
""", unsafe_allow_html=True)

# === БАННЕР ===
st.markdown("""
<div style="text-align: center; margin-bottom: 30px;">
    <div style="font-size: 60px;">☕🍰🥗</div>
    <h1 style="font-size: 42px; margin: 0; color: #d4a373;">Система учета кафе</h1>
    <p style="color: #b0a090; font-size: 18px;">Ваш помощник в управлении ресторанным бизнесом</p>
</div>
""", unsafe_allow_html=True)

# === БОКОВАЯ ПАНЕЛЬ ===
with st.sidebar:
    st.markdown("### ☕ Меню управления")
    st.markdown("---")
    menu = st.selectbox(
        "Выберите раздел",
        ["📊 Дашборд", "📝 Заказы", "🍽️ Меню", "📦 Склад", "👥 Персонал", "📈 Отчеты"],
        format_func=lambda x: x.split(" ", 1)[1] if " " in x else x,
    )
    st.markdown("---")
    st.info("🗄️ Данные в файле Кафе_Учет.xlsx\n\n💾 Закройте Excel перед работой")
    st.markdown("---")
    st.markdown('<div style="font-size: 12px; color: #888;">© 2026 | Кафе Учет</div>', unsafe_allow_html=True)

# === ЗАГРУЗКА ДАННЫХ ===
@st.cache_data
def load_data():
    if not os.path.exists(EXCEL_FILE):
        return None, None, None, None
    try:
        orders = pd.read_excel(EXCEL_FILE, sheet_name="ЗАКАЗЫ")
        if "Дата" in orders.columns:
            orders["Дата"] = pd.to_datetime(orders["Дата"], errors="coerce", dayfirst=True)
            orders = orders[orders["Дата"].notna()]
        menu_data = pd.read_excel(EXCEL_FILE, sheet_name="МЕНЮ")
        ingredients = pd.read_excel(EXCEL_FILE, sheet_name="ИНГРЕДИЕНТЫ")
        recipes = pd.read_excel(EXCEL_FILE, sheet_name="РЕЦЕПТЫ")
        return orders, menu_data, ingredients, recipes
    except Exception as e:
        st.error(f"Ошибка загрузки данных: {e}")
        return None, None, None, None


orders, menu_data, ingredients, recipes = load_data()

# === ДАШБОРД ===
if menu == "📊 Дашборд":
    if orders is None:
        st.warning("📭 Файл данных не найден. Создайте 'Кафе_Учет.xlsx'.")
    elif len(orders) == 0:
        st.info("📭 Нет заказов. Добавьте первый заказ в разделе «Заказы».")
    else:
        st.markdown("### 📈 Общая статистика")
        period = st.radio("📅 Период", ["Сегодня", "Неделя", "Месяц", "Всё время"], horizontal=True)
        filtered = filter_orders_by_period(orders, period)

        revenue = filtered["Сумма (₸)"].sum()
        num = len(filtered)
        avg = revenue / num if num > 0 else 0
        food_cost = calculate_food_cost(filtered, menu_data)

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("💰 Выручка", f"{revenue:,.0f} ₸")
        c2.metric("📋 Заказов", num)
        c3.metric("🧾 Средний чек", f"{avg:,.0f} ₸")
        c4.metric("🎯 Фудкост", f"{food_cost:.0f}%")

        if len(filtered) == 0:
            st.info("За выбранный период заказов нет.")
        else:
            col_left, col_right = st.columns(2)
            with col_left:
                st.markdown("### 📅 Продажи по дням")
                daily = filtered.groupby(filtered["Дата"].dt.date)["Сумма (₸)"].sum().reset_index()
                fig = px.line(
                    daily, x="Дата", y="Сумма (₸)",
                    template="plotly_dark", color_discrete_sequence=["#d4a373"],
                )
                fig.update_layout(plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(fig, use_container_width=True)

            with col_right:
                st.markdown("### 🍽️ Популярные блюда")
                counts = build_popular_dishes_df(filtered).head(10)
                if not counts.empty:
                    fig = px.bar(
                        counts, x="Количество", y="Блюдо", orientation="h",
                        template="plotly_dark", color_discrete_sequence=["#d4a373"],
                    )
                    fig.update_layout(
                        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)", height=400
                    )
                    st.plotly_chart(fig, use_container_width=True)

# === ЗАКАЗЫ ===
elif menu == "📝 Заказы":
    if orders is None or menu_data is None:
        st.warning("📭 Нет данных. Проверьте файл 'Кафе_Учет.xlsx'.")
    else:
        st.markdown("### 🆕 Новый заказ")

        # Загрузка списка сотрудников из Excel
        @st.cache_data
        def load_staff_list():
            try:
                staff_df = pd.read_excel(EXCEL_FILE, sheet_name="ПЕРСОНАЛ")
                active_staff = staff_df[(staff_df["Активен"] == "Да") & 
                                         ((staff_df["Должность"] == "Официант") | 
                                          (staff_df["Должность"] == "Администратор"))]
                staff_names = active_staff["Имя"].tolist()
                return staff_names if staff_names else ["Анна", "Иван", "Елена", "Сергей"]
            except:
                return ["Анна", "Иван", "Елена", "Сергей"]

        waiters_list = load_staff_list()

        if "cart" not in st.session_state:
            st.session_state.cart = {}

        subcategories_df = load_subcategories()
        col1, col2 = st.columns(2)

        with col1:
            with st.container():
                st.markdown('<div class="dish-card">', unsafe_allow_html=True)
                st.markdown("#### 📋 Информация о заказе")

                c1, c2 = st.columns(2)
                with c1:
                    order_date = st.date_input("📅 Дата")
                    table = st.number_input("🍽️ Столик", min_value=1, value=1)
                    waiter = st.selectbox("👨‍🍳 Официант", waiters_list)
                with c2:
                    order_time = st.time_input("⏰ Время")
                    guests = st.number_input("👥 Гостей", min_value=1, value=2)
                    payment = st.selectbox("💳 Оплата", ["Наличные", "Карта"])

                st.markdown("---")
                st.markdown("#### 🍕 Выберите блюда")

                if subcategories_df is not None and len(subcategories_df) > 0:
                    main_categories = subcategories_df["Основная категория"].unique().tolist()
                    tabs = st.tabs(main_categories)

                    for tab_idx, main_cat in enumerate(main_categories):
                        with tabs[tab_idx]:
                            subcategories = (
                                subcategories_df[subcategories_df["Основная категория"] == main_cat]
                                ["Подкатегория"].unique().tolist()
                            )
                            for subcat in subcategories:
                                st.markdown(f"**📌 {subcat}**")
                                dishes_df = subcategories_df[
                                    (subcategories_df["Основная категория"] == main_cat)
                                    & (subcategories_df["Подкатегория"] == subcat)
                                ]
                                for idx, row in dishes_df.iterrows():
                                    dish = row["Блюдо"]
                                    price = get_dish_price(dish, menu_data)
                                    current_qty = st.session_state.cart.get(dish, 0)

                                    col_a, col_b, col_c, col_d = st.columns([4, 1, 1, 1])
                                    with col_a:
                                        st.markdown(f"**{dish}** — {price:,.0f} ₸")
                                    with col_b:
                                        if st.button("➖", key=f"minus_{main_cat}_{subcat}_{idx}"):
                                            if current_qty > 0:
                                                st.session_state.cart[dish] = current_qty - 1
                                                if st.session_state.cart[dish] == 0:
                                                    del st.session_state.cart[dish]
                                            st.rerun()
                                    with col_c:
                                        st.markdown(
                                            f"<div style='text-align: center; font-size: 18px; font-weight: bold;'>{current_qty}</div>",
                                            unsafe_allow_html=True,
                                        )
                                    with col_d:
                                        if st.button("➕", key=f"plus_{main_cat}_{subcat}_{idx}"):
                                            st.session_state.cart[dish] = current_qty + 1
                                            st.rerun()
                                    st.markdown("---")
                else:
                    st.warning("Структура меню не загружена. Добавьте блюда в разделе «Меню».")
                    for idx, row in menu_data.iterrows():
                        dish = row["Блюдо"]
                        price = row["Цена продажи (₸)"]
                        current_qty = st.session_state.cart.get(dish, 0)

                        col_a, col_b, col_c, col_d = st.columns([4, 1, 1, 1])
                        with col_a:
                            st.markdown(f"**{dish}** — {price:,.0f} ₸")
                        with col_b:
                            if st.button("➖", key=f"minus_simple_{idx}"):
                                if current_qty > 0:
                                    st.session_state.cart[dish] = current_qty - 1
                                    if st.session_state.cart[dish] == 0:
                                        del st.session_state.cart[dish]
                                st.rerun()
                        with col_c:
                            st.markdown(
                                f"<div style='text-align: center; font-size: 18px;'>{current_qty}</div>",
                                unsafe_allow_html=True,
                            )
                        with col_d:
                            if st.button("➕", key=f"plus_simple_{idx}"):
                                st.session_state.cart[dish] = current_qty + 1
                                st.rerun()
                        st.markdown("---")

                note = st.text_area("📝 Примечание")
                col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
                with col_btn2:
                    submitted = st.button("✅ Оформить заказ", use_container_width=True)

                if submitted:
                    if not st.session_state.cart:
                        st.warning("Добавьте хотя бы одно блюдо в заказ.")
                    else:
                        total = 0
                        dish_details = []
                        for dish, qty in st.session_state.cart.items():
                            price = get_dish_price(dish, menu_data)
                            if price <= 0:
                                st.error(f"Блюдо «{dish}» не найдено в меню или цена не указана.")
                                break
                            total += price * qty
                            dish_details.append(f"{qty}x {dish}" if qty > 1 else dish)
                        else:
                            new_row = pd.DataFrame([{
                                "Дата": order_date.strftime("%d.%m.%Y"),
                                "Время": order_time.strftime("%H:%M"),
                                "Столик": table,
                                "Гости": guests,
                                "Официант": waiter,
                                "Блюда": ", ".join(dish_details),
                                "Сумма (₸)": total,
                                "Оплата": payment,
                                "Примечание": note,
                            }])
                            try:
                                existing = pd.read_excel(EXCEL_FILE, sheet_name="ЗАКАЗЫ")
                                existing = existing[existing["Дата"].notna()]
                                updated = pd.concat([existing, new_row], ignore_index=True)
                                if safe_save_to_excel(updated, "ЗАКАЗЫ", mode="replace"):
                                    st.success(f"✅ Заказ оформлен! Сумма: {total:,.0f} ₸")
                                    st.balloons()
                                    st.session_state.cart = {}
                                    refresh_data()
                                    st.rerun()
                            except Exception as e:
                                st.error(f"Ошибка: {e}")

                if st.session_state.cart and st.button("🗑️ Очистить корзину"):
                    st.session_state.cart = {}
                    st.rerun()

                st.markdown("</div>", unsafe_allow_html=True)

        with col2:
            st.markdown("### 📜 Последние заказы")
            if len(orders) > 0:
                disp = orders.tail(10).copy()
                if "Дата" in disp.columns:
                    disp["Дата"] = disp["Дата"].dt.strftime("%d.%m.%Y")
                st.dataframe(
                    disp[["Дата", "Время", "Столик", "Блюда", "Сумма (₸)", "Официант"]],
                    use_container_width=True,
                    height=400,
                )

            if st.session_state.cart:
                st.markdown("---")
                st.markdown("### 🛒 Текущий заказ")
                total_preview = 0
                for dish, qty in st.session_state.cart.items():
                    price = get_dish_price(dish, menu_data)
                    subtotal = price * qty
                    total_preview += subtotal
                    st.markdown(f"- **{dish}**: {qty} шт × {price:,.0f} ₸ = {subtotal:,.0f} ₸")
                st.markdown(f"**Итого: {total_preview:,.0f} ₸**")

# === МЕНЮ ===
elif menu == "🍽️ Меню":
    if menu_data is None:
        st.warning("📭 Файл данных не найден.")
    else:
        st.markdown("### 🍽️ Управление меню")
        structure_df = load_full_structure()

        with st.expander("🏷️ Управление категориями и подкатегориями", expanded=False):
            main_cats = structure_df["Основная категория"].dropna().unique().tolist()
            if st.button("🔄 Синхронизировать все листы", key="sync_all_sheets"):
                sync_all_sheets(structure_df)
                refresh_data()
                st.success("✅ Листы ПОДКАТЕГОРИИ, КАТЕГОРИИ и МЕНЮ синхронизированы!")
                st.rerun()

            col_cat1, col_cat2 = st.columns(2)

            with col_cat1:
                st.markdown("**➕ Новая основная категория**")
                new_main_cat = st.text_input("Название категории", key="new_main_cat")
                if st.button("✅ Добавить категорию", key="add_main_cat"):
                    if new_main_cat and new_main_cat not in main_cats:
                        new_row = pd.DataFrame([{
                            "Основная категория": new_main_cat,
                            "Подкатегория": "",
                            "Блюдо": "",
                            "Цена (₸)": 0,
                            "Себестоимость (₸)": 0,
                        }])
                        save_full_structure(pd.concat([structure_df, new_row], ignore_index=True))
                        st.success(f"✅ Категория «{new_main_cat}» добавлена!")
                        st.rerun()

            with col_cat2:
                st.markdown("**📁 Новая подкатегория**")
                if main_cats:
                    selected_main = st.selectbox(
                        "Выберите основную категорию", main_cats, key="select_main_for_sub"
                    )
                    new_subcat = st.text_input("Название подкатегории", key="new_subcat")
                    if st.button("✅ Добавить подкатегорию", key="add_subcat"):
                        if new_subcat and selected_main:
                            exists = structure_df[
                                (structure_df["Основная категория"] == selected_main)
                                & (structure_df["Подкатегория"] == new_subcat)
                            ]
                            if exists.empty:
                                new_row = pd.DataFrame([{
                                    "Основная категория": selected_main,
                                    "Подкатегория": new_subcat,
                                    "Блюдо": "",
                                    "Цена (₸)": 0,
                                    "Себестоимость (₸)": 0,
                                }])
                                save_full_structure(
                                    pd.concat([structure_df, new_row], ignore_index=True)
                                )
                                st.success(
                                    f"✅ Подкатегория «{new_subcat}» добавлена в «{selected_main}»!"
                                )
                                st.rerun()
                            else:
                                st.warning("Такая подкатегория уже существует")
                else:
                    st.info("Сначала добавьте основную категорию.")

            st.markdown("---")
            st.markdown("**📋 Текущая структура**")
            for main_cat in sorted(main_cats):
                st.markdown(f"### 📌 {main_cat}")
                subcats = structure_df[structure_df["Основная категория"] == main_cat][
                    "Подкатегория"
                ].dropna().unique()
                subcats = [s for s in subcats if s != ""]

                for subcat in subcats:
                    col_del1, col_del2 = st.columns([4, 1])
                    with col_del1:
                        st.markdown(f"   📁 **{subcat}**")
                    with col_del2:
                        if st.button("🗑️", key=f"del_subcat_{main_cat}_{subcat}"):
                            dishes_to_remove = get_dishes_in_scope(
                                structure_df, main_cat, subcat
                            )
                            remove_menu_dishes(dishes_to_remove)
                            updated = structure_df[
                                ~(
                                    (structure_df["Основная категория"] == main_cat)
                                    & (structure_df["Подкатегория"] == subcat)
                                )
                            ]
                            save_full_structure(updated)
                            st.success(
                                f"Подкатегория «{subcat}» удалена"
                                + (f" ({len(dishes_to_remove)} блюд)" if dishes_to_remove else "")
                                + "!"
                            )
                            st.rerun()

                col_del_cat1, col_del_cat2 = st.columns([4, 1])
                with col_del_cat2:
                    if st.button("🗑️ Удалить категорию", key=f"del_main_cat_{main_cat}"):
                        dishes_to_remove = get_dishes_in_scope(structure_df, main_cat)
                        remove_menu_dishes(dishes_to_remove)
                        updated = structure_df[structure_df["Основная категория"] != main_cat]
                        save_full_structure(updated)
                        st.success(
                            f"Категория «{main_cat}» удалена"
                            + (f" ({len(dishes_to_remove)} блюд)" if dishes_to_remove else "")
                            + "!"
                        )
                        st.rerun()
                st.markdown("---")

        with st.expander("➕ Добавить новое блюдо", expanded=True):
            main_cats_list = structure_df["Основная категория"].dropna().unique().tolist()
            with st.form("new_dish_form", clear_on_submit=True):
                col1, col2 = st.columns(2)
                with col1:
                    if main_cats_list:
                        selected_main_dish = st.selectbox(
                            "Основная категория", main_cats_list, key="dish_main_cat"
                        )
                        subcats_list = structure_df[
                            structure_df["Основная категория"] == selected_main_dish
                        ]["Подкатегория"].dropna().unique()
                        subcats_list = [s for s in subcats_list if s != ""]
                        selected_subcat = st.selectbox(
                            "Подкатегория", subcats_list or [""], key="dish_subcat"
                        )
                    else:
                        selected_main_dish = ""
                        selected_subcat = ""
                        st.info("Сначала создайте категории выше.")
                    new_dish = st.text_input("🍽️ Название блюда")
                    new_ingredients = st.text_input("🥕 Ингредиенты (через запятую)")
                with col2:
                    new_cost = st.number_input("📦 Себестоимость (₸)", min_value=0, value=100)
                    new_price = st.number_input("💰 Цена продажи (₸)", min_value=0, value=300)

                submitted = st.form_submit_button("✅ Сохранить блюдо", use_container_width=True)

                if submitted:
                    if not (new_dish and selected_main_dish and selected_subcat):
                        st.warning("Заполните категорию, подкатегорию и название блюда.")
                    elif new_dish in structure_df["Блюдо"].values:
                        st.warning(f"Блюдо «{new_dish}» уже существует.")
                    else:
                        new_row = pd.DataFrame([{
                            "Основная категория": selected_main_dish,
                            "Подкатегория": selected_subcat,
                            "Блюдо": new_dish,
                            "Цена (₸)": new_price,
                            "Себестоимость (₸)": new_cost,
                        }])
                        save_full_structure(pd.concat([structure_df, new_row], ignore_index=True))
                        upsert_menu_dish(
                            new_dish, selected_main_dish, new_ingredients, new_cost, new_price
                        )
                        refresh_data()
                        st.success(
                            f"✅ Блюдо «{new_dish}» добавлено в «{selected_main_dish} → {selected_subcat}»!"
                        )
                        st.rerun()

        st.markdown("---")
        st.markdown("#### 📋 Полное меню")

        structure_df = load_full_structure()
        for main_cat in sorted(structure_df["Основная категория"].dropna().unique()):
            st.markdown(f"### 📌 {main_cat}")
            subcats = structure_df[structure_df["Основная категория"] == main_cat][
                "Подкатегория"
            ].dropna().unique()
            subcats = [s for s in subcats if s != ""]

            for subcat in subcats:
                st.markdown(f"####    📁 {subcat}")
                dishes_df = structure_df[
                    (structure_df["Основная категория"] == main_cat)
                    & (structure_df["Подкатегория"] == subcat)
                    & (structure_df["Блюдо"].astype(str).str.strip() != "")
                ]

                for idx, row in dishes_df.iterrows():
                    dish = row["Блюдо"]
                    price = row["Цена (₸)"] or 0
                    cost = row["Себестоимость (₸)"] or 0
                    margin = calculate_margin(price, cost)

                    col_d1, col_d2, col_d3, col_d4 = st.columns([5, 2, 1, 1])
                    with col_d1:
                        st.markdown(f"🍽️ **{dish}**")
                    with col_d2:
                        st.markdown(f"{price:,.0f} ₸ | {margin:.0f}%")
                    with col_d3:
                        if st.button("✏️", key=f"edit_full_{idx}"):
                            st.session_state.edit_dish_full = dish
                            st.session_state.edit_price_full = price
                            st.session_state.edit_cost_full = cost
                    with col_d4:
                        if st.button("🗑️", key=f"del_full_{idx}"):
                            updated = structure_df[
                                ~(
                                    (structure_df["Основная категория"] == main_cat)
                                    & (structure_df["Подкатегория"] == subcat)
                                    & (structure_df["Блюдо"] == dish)
                                )
                            ]
                            save_full_structure(updated)
                            remove_menu_dish(dish)
                            refresh_data()
                            st.rerun()
                    st.markdown("---")

        if "edit_dish_full" in st.session_state:
            st.markdown("---")
            st.subheader(f"✏️ Редактирование: {st.session_state.edit_dish_full}")
            with st.form("edit_full_form"):
                new_price = st.number_input(
                    "Цена (₸)", min_value=0, value=int(st.session_state.edit_price_full)
                )
                new_cost = st.number_input(
                    "Себестоимость (₸)", min_value=0, value=int(st.session_state.edit_cost_full)
                )
                if st.form_submit_button("✅ Сохранить"):
                    structure_df = load_full_structure()
                    idx = structure_df[
                        structure_df["Блюдо"] == st.session_state.edit_dish_full
                    ].index[0]
                    structure_df.at[idx, "Цена (₸)"] = new_price
                    structure_df.at[idx, "Себестоимость (₸)"] = new_cost
                    save_full_structure(structure_df)

                    menu_df = pd.read_excel(EXCEL_FILE, sheet_name="МЕНЮ")
                    if st.session_state.edit_dish_full in menu_df["Блюдо"].values:
                        menu_idx = menu_df[
                            menu_df["Блюдо"] == st.session_state.edit_dish_full
                        ].index[0]
                        category = menu_df.at[menu_idx, "Категория"]
                        ingredients_text = menu_df.at[menu_idx, "Ингредиенты"]
                        upsert_menu_dish(
                            st.session_state.edit_dish_full,
                            category,
                            ingredients_text,
                            new_cost,
                            new_price,
                        )

                    del st.session_state.edit_dish_full
                    refresh_data()
                    st.success("✅ Изменения сохранены!")
                    st.rerun()

        st.markdown("---")
        st.markdown("#### 📊 Маржинальность блюд")

        chart_data = menu_data.copy()
        chart_data["Маржинальность (%)"] = chart_data.apply(
            lambda x: calculate_margin(x["Цена продажи (₸)"], x["Себестоимость (₸)"]),
            axis=1,
        )
        color_col = "Категория" if "Категория" in chart_data.columns else None
        fig = px.bar(
            chart_data,
            x="Блюдо",
            y="Маржинальность (%)",
            color=color_col,
            title="Маржинальность блюд",
            text_auto=".1f",
            color_discrete_sequence=["#d4a373"] if color_col is None else None,
        )
        fig.update_layout(plot_bgcolor="rgba(0,0,0,0)", height=500, paper_bgcolor="rgba(0,0,0,0)")
        fig.update_traces(textfont_color="white")
        st.plotly_chart(fig, use_container_width=True)

# === СКЛАД ===
elif menu == "📦 Склад":
    if ingredients is None:
        st.warning("📭 Файл данных не найден.")
    else:
        st.markdown("### 📦 Управление складом")

        with st.expander("➕ Добавить ингредиент"):
            with st.form("new_ing"):
                col1, col2 = st.columns(2)
                with col1:
                    name = st.text_input("Название")
                    unit = st.selectbox("Ед.", ["кг", "г", "л", "шт"])
                    stock = st.number_input("Остаток", min_value=0, value=10)
                with col2:
                    min_stock = st.number_input("Мин. запас", min_value=0, value=2)
                    price = st.number_input("Цена за ед.", min_value=0, value=100)
                if st.form_submit_button("✅ Добавить"):
                    if not name.strip():
                        st.warning("Введите название ингредиента.")
                    else:
                        new = pd.DataFrame([{
                            "Ингредиент": name.strip(),
                            "Единица": unit,
                            "Остаток": stock,
                            "Мин. запас": min_stock,
                            "Цена за ед. (₸)": price,
                        }])
                        existing = pd.read_excel(EXCEL_FILE, sheet_name="ИНГРЕДИЕНТЫ")
                        updated = pd.concat([existing, new], ignore_index=True)
                        if safe_save_to_excel(updated, "ИНГРЕДИЕНТЫ", mode="replace"):
                            refresh_data()
                            st.rerun()

        ingredients = ingredients.copy()
        ingredients["Статус"] = ingredients.apply(
            lambda x: "⚠️" if x["Остаток"] < x["Мин. запас"] else "✅", axis=1
        )
        for i, row in ingredients.iterrows():
            st.markdown(f"""
            <div class="dish-card">
                <div style="display: flex; justify-content: space-between;">
                    <div><strong>{row['Ингредиент']}</strong> ({row['Единица']}) {row['Статус']}</div>
                    <span style="background: #d4a373; padding: 2px 10px; border-radius: 15px;">{row['Остаток']}</span>
                </div>
                <div>Мин. запас: {row['Мин. запас']} | Цена: {row['Цена за ед. (₸)']} ₸/{row['Единица']}</div>
            </div>
            """, unsafe_allow_html=True)
            if st.button("🗑️", key=f"del_ing_{i}"):
                updated = ingredients[ingredients["Ингредиент"] != row["Ингредиент"]].drop(
                    columns=["Статус"]
                )
                if safe_save_to_excel(updated, "ИНГРЕДИЕНТЫ", mode="replace"):
                    refresh_data()
                    st.rerun()

# === ОТЧЕТЫ ===
elif menu == "📈 Отчеты":
    if orders is None:
        st.warning("📭 Файл данных не найден.")
    elif len(orders) == 0:
        st.info("📭 Нет заказов для отчёта.")
    else:
        st.markdown("### 📈 Отчеты")

        col_period, col_export = st.columns([3, 1])
        with col_period:
            report_period = st.radio(
                "📅 Период отчёта",
                ["Сегодня", "Неделя", "Месяц", "Всё время"],
                horizontal=True,
            )
        filtered_report = filter_orders_by_period(orders, report_period)

        if len(filtered_report) == 0:
            st.info("За выбранный период заказов нет.")
        else:
            rev = filtered_report["Сумма (₸)"].sum()
            food_cost = calculate_food_cost(filtered_report, menu_data)
            num_orders = len(filtered_report)
            avg_check = rev / num_orders if num_orders > 0 else 0
            profit = rev * (1 - food_cost / 100)

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("💰 Выручка", f"{rev:,.0f} ₸")
            c2.metric("📋 Заказов", num_orders)
            c3.metric("🧾 Средний чек", f"{avg_check:,.0f} ₸")
            c4.metric("🎯 Фудкост", f"{food_cost:.0f}%")

            with col_export:
                report_file = build_report_excel(filtered_report, menu_data, report_period)
                period_slug = {
                    "Сегодня": "сегодня",
                    "Неделя": "неделя",
                    "Месяц": "месяц",
                    "Всё время": "все_время",
                }[report_period]
                filename = f"Отчет_кафе_{period_slug}_{datetime.now().strftime('%Y-%m-%d')}.xlsx"
                st.download_button(
                    label="📥 Скачать Excel",
                    data=report_file,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

            tab_orders, tab_daily, tab_popular, tab_payment = st.tabs(
                ["📋 Заказы", "📅 По дням", "🍽️ Блюда", "💳 Оплата"]
            )

            with tab_orders:
                disp = filtered_report.copy()
                if "Дата" in disp.columns:
                    disp["Дата"] = disp["Дата"].dt.strftime("%d.%m.%Y")
                st.dataframe(disp, use_container_width=True, height=400)

            with tab_daily:
                daily = (
                    filtered_report.groupby(filtered_report["Дата"].dt.date)["Сумма (₸)"]
                    .agg(["sum", "count"])
                    .reset_index()
                )
                daily.columns = ["Дата", "Выручка (₸)", "Заказов"]
                st.dataframe(daily, use_container_width=True)

            with tab_popular:
                popular = build_popular_dishes_df(filtered_report)
                if popular.empty:
                    st.info("Нет данных о блюдах.")
                else:
                    st.dataframe(popular, use_container_width=True)

            with tab_payment:
                if "Оплата" in filtered_report.columns:
                    payment_stats = (
                        filtered_report.groupby("Оплата")
                        .agg(**{"Заказов": ("Сумма (₸)", "count"), "Сумма (₸)": ("Сумма (₸)", "sum")})
                        .reset_index()
                    )
                    st.dataframe(payment_stats, use_container_width=True)
                else:
                    st.info("Данные об оплате отсутствуют.")

            st.caption(
                f"Прибыль за период (оценка): **{profit:,.0f} ₸** | "
                f"Файл Excel содержит 5 листов: Сводка, Заказы, По дням, Популярные блюда, По оплате."
            )
# === ПЕРСОНАЛ ===
elif menu == "👥 Персонал":
    st.markdown("### 👥 Управление персоналом")
    
    # Функция для загрузки персонала
    @st.cache_data
    def load_staff():
        try:
            df = pd.read_excel(EXCEL_FILE, sheet_name="ПЕРСОНАЛ")
            return df
        except:
            # Создаём пустой DataFrame с правильными колонками
            return pd.DataFrame(columns=["ID", "Имя", "Должность", "Телефон", "PIN-код", "Активен", "Дата добавления"])
    
    # Функция для сохранения персонала
    def save_staff(df):
        safe_save_to_excel(df, "ПЕРСОНАЛ", mode="replace")
        refresh_data()
    
    # Загружаем данные
    staff_df = load_staff()
    
    # Форма для добавления нового сотрудника
    with st.expander("➕ Добавить сотрудника", expanded=True):
        with st.form("add_staff_form"):
            col1, col2 = st.columns(2)
            with col1:
                new_name = st.text_input("👤 Имя сотрудника")
                new_role = st.selectbox("📌 Должность", ["Администратор", "Официант", "Повар", "Бухгалтер"])
                new_phone = st.text_input("📞 Телефон")
            with col2:
                new_pin = st.text_input("🔐 PIN-код (4 цифры)", max_chars=4, type="password")
                new_active = st.checkbox("Активен", value=True)
            
            if st.form_submit_button("✅ Добавить сотрудника"):
                if new_name and new_pin:
                    # Генерируем новый ID
                    new_id = staff_df["ID"].max() + 1 if not staff_df.empty else 1
                    new_row = pd.DataFrame([{
                        "ID": new_id,
                        "Имя": new_name,
                        "Должность": new_role,
                        "Телефон": new_phone,
                        "PIN-код": new_pin,
                        "Активен": "Да" if new_active else "Нет",
                        "Дата добавления": datetime.now().strftime("%d.%m.%Y")
                    }])
                    updated = pd.concat([staff_df, new_row], ignore_index=True)
                    save_staff(updated)
                    st.success(f"✅ Сотрудник '{new_name}' добавлен!")
                    st.rerun()
                else:
                    st.warning("Заполните имя и PIN-код")
    
    st.markdown("---")
    st.markdown("#### 📋 Список сотрудников")
    
    if staff_df.empty:
        st.info("Нет сотрудников. Добавьте первого.")
    else:
        # Отображаем каждого сотрудника
        for i, row in staff_df.iterrows():
            col1, col2, col3, col4, col5 = st.columns([3, 2, 2, 1, 1])
            
            with col1:
                st.markdown(f"**{row['Имя']}**")
                st.caption(row['Должность'])
            with col2:
                st.markdown(f"📞 {row['Телефон'] if pd.notna(row['Телефон']) else '—'}")
            with col3:
                status = "🟢 Активен" if row['Активен'] == "Да" else "🔴 Неактивен"
                st.markdown(status)
            with col4:
                if st.button("✏️", key=f"edit_staff_{i}"):
                    st.session_state.edit_staff = row['ID']
                    st.session_state.edit_staff_name = row['Имя']
                    st.session_state.edit_staff_role = row['Должность']
                    st.session_state.edit_staff_phone = row['Телефон'] if pd.notna(row['Телефон']) else ""
                    st.session_state.edit_staff_active = row['Активен'] == "Да"
            with col5:
                if st.button("🗑️", key=f"del_staff_{i}"):
                    updated = staff_df[staff_df["ID"] != row['ID']]
                    save_staff(updated)
                    st.success(f"🗑️ Сотрудник '{row['Имя']}' удалён!")
                    st.rerun()
            st.markdown("---")
        
        # Форма редактирования
        if 'edit_staff' in st.session_state:
            st.markdown("---")
            st.subheader(f"✏️ Редактирование: {st.session_state.edit_staff_name}")
            with st.form("edit_staff_form"):
                col1, col2 = st.columns(2)
                with col1:
                    edit_name = st.text_input("Имя", value=st.session_state.edit_staff_name)
                    edit_role = st.selectbox("Должность", ["Администратор", "Официант", "Повар", "Бухгалтер"], 
                                             index=["Администратор", "Официант", "Повар", "Бухгалтер"].index(st.session_state.edit_staff_role))
                    edit_phone = st.text_input("Телефон", value=st.session_state.edit_staff_phone)
                with col2:
                    edit_active = st.checkbox("Активен", value=st.session_state.edit_staff_active)
                
                if st.form_submit_button("✅ Сохранить изменения"):
                    idx = staff_df[staff_df["ID"] == st.session_state.edit_staff].index[0]
                    staff_df.at[idx, "Имя"] = edit_name
                    staff_df.at[idx, "Должность"] = edit_role
                    staff_df.at[idx, "Телефон"] = edit_phone
                    staff_df.at[idx, "Активен"] = "Да" if edit_active else "Нет"
                    save_staff(staff_df)
                    del st.session_state.edit_staff
                    st.success("✅ Изменения сохранены!")
                    st.rerun()
    
    # Статистика по сотрудникам (если есть заказы)
    if orders is not None and len(orders) > 0 and not staff_df.empty:
        st.markdown("---")
        st.markdown("#### 📊 Статистика сотрудников")
        
        # Считаем заказы по официантам
        staff_stats = []
        for _, row in staff_df.iterrows():
            if row['Должность'] in ["Официант", "Администратор"]:
                staff_orders = orders[orders["Официант"] == row['Имя']]
                if len(staff_orders) > 0:
                    staff_stats.append({
                        "Сотрудник": row['Имя'],
                        "Заказов": len(staff_orders),
                        "Выручка": staff_orders["Сумма (₸)"].sum(),
                        "Ср. чек": staff_orders["Сумма (₸)"].sum() / len(staff_orders)
                    })
        
        if staff_stats:
            stats_df = pd.DataFrame(staff_stats)
            st.dataframe(stats_df, use_container_width=True)
            
            # График выручки по сотрудникам
            fig = px.bar(stats_df, x="Сотрудник", y="Выручка", 
                         title="Выручка по сотрудникам",
                         template="plotly_dark", color_discrete_sequence=["#d4a373"])
            fig.update_layout(plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig, use_container_width=True)
st.markdown(
    '<p style="text-align: center; color: #888; font-size: 12px;">☕ Работает на Streamlit</p>',
    unsafe_allow_html=True,
)
