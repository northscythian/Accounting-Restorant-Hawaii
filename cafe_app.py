import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from openpyxl import load_workbook
import os
import time

# === НАСТРОЙКА СТРАНИЦЫ ===
st.set_page_config(
    page_title="Кафе Учет",
    page_icon="☕",
    layout="wide",
    initial_sidebar_state="expanded"
)

# === ФУНКЦИЯ ДЛЯ БЕЗОПАСНОЙ ЗАПИСИ В EXCEL ===
def safe_save_to_excel(df, sheet_name, mode='replace'):
    """Безопасное сохранение DataFrame в Excel с повторными попытками"""
    max_attempts = 3
    for attempt in range(max_attempts):
        try:
            with pd.ExcelWriter("Кафе_Учет.xlsx", engine="openpyxl", mode='a', if_sheet_exists=mode) as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
            return True
        except PermissionError:
            if attempt < max_attempts - 1:
                time.sleep(1)  # Ждём 1 секунду
                continue
            else:
                st.error("❌ Файл Excel заблокирован. Закройте файл 'Кафе_Учет.xlsx' и попробуйте снова.")
                return False
        except Exception as e:
            # Если файла нет или другая ошибка — пробуем пересоздать
            try:
                with pd.ExcelWriter("Кафе_Учет.xlsx", engine="openpyxl", mode='w') as writer:
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                return True
            except:
                st.error(f"❌ Ошибка сохранения: {e}")
                return False
    return False

# === ФУНКЦИЯ ДЛЯ ОБНОВЛЕНИЯ ДАННЫХ (сброс кеша) ===
def refresh_data():
    st.cache_data.clear()
    # Не вызываем st.rerun() здесь, чтобы избежать рекурсии

# === КАСТОМНЫЙ CSS (тёмная тема) ===
st.markdown("""
<style>
    /* Основной фон */
    .stApp {
        background: linear-gradient(135deg, #2d2b2a 0%, #1a1a1a 100%);
    }
    
    /* Карточки метрик */
    div[data-testid="stMetric"] {
        background: linear-gradient(135deg, #3d3a38, #2c2a28);
        border-radius: 20px;
        padding: 20px;
        box-shadow: 0 4px 20px rgba(0,0,0,0.3);
        border: 1px solid #d4a373;
        transition: transform 0.3s ease;
    }
    
    div[data-testid="stMetric"]:hover {
        transform: translateY(-5px);
    }
    
    div[data-testid="stMetric"] label {
        color: #d4a373 !important;
        font-size: 16px !important;
    }
    
    div[data-testid="stMetric"] div {
        color: white !important;
        font-size: 32px !important;
        font-weight: bold !important;
    }
    
    /* Заголовки */
    h1, h2, h3, h4 {
        color: #d4a373 !important;
        font-family: 'Segoe UI', 'Arial', sans-serif;
        font-weight: 600 !important;
        text-shadow: 1px 1px 2px rgba(0,0,0,0.5);
    }
    
    h1 {
        border-bottom: 3px solid #d4a373;
        display: inline-block;
        padding-bottom: 10px;
    }
    
    /* Боковая панель */
    [data-testid="stSidebar"] {
        background: #1a1a1a;
        border-right: 2px solid #d4a373;
    }
    
    [data-testid="stSidebar"] * {
        color: #e0e0e0 !important;
    }
    
    [data-testid="stSidebar"] .stSelectbox label {
        color: #d4a373 !important;
        font-weight: bold;
    }
    
    [data-testid="stSidebar"] .stSelectbox div {
        color: white !important;
    }
    
    /* Кнопки */
    .stButton > button {
        background: linear-gradient(135deg, #d4a373, #b5835a);
        color: white;
        border: none;
        border-radius: 30px;
        padding: 12px 28px;
        font-weight: bold;
        transition: all 0.3s ease;
        box-shadow: 0 2px 10px rgba(0,0,0,0.3);
    }
    
    .stButton > button:hover {
        transform: scale(1.02);
        background: linear-gradient(135deg, #e0b184, #c4936a);
        color: white;
    }
    
    /* Формы */
    .stTextInput > div > div > input, 
    .stNumberInput > div > div > input,
    .stTextArea > div > div > textarea,
    .stSelectbox > div > div > select,
    .stDateInput > div > div > input,
    .stTimeInput > div > div > input {
        background-color: #2a2a2a;
        border-radius: 15px;
        border: 1px solid #d4a373;
        font-size: 16px;
        color: white !important;
    }
    
    .stTextInput label, .stNumberInput label, .stTextArea label {
        color: #d4a373 !important;
        font-weight: bold;
    }
    
    /* Таблицы */
    .stDataFrame {
        background: #2a2a2a;
        border-radius: 15px;
        overflow: hidden;
        box-shadow: 0 2px 10px rgba(0,0,0,0.3);
        border: 1px solid #d4a373;
    }
    
    .stDataFrame table {
        color: white !important;
    }
    
    .stDataFrame th {
        background: #d4a373 !important;
        color: #1a1a1a !important;
        font-weight: bold;
    }
    
    .stDataFrame td {
        color: white !important;
    }
    
    /* Вкладки */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: #2a2a2a;
        border-radius: 30px;
        padding: 5px;
    }
    
    .stTabs [data-baseweb="tab"] {
        border-radius: 25px;
        padding: 8px 20px;
        font-weight: bold;
        color: #e0e0e0;
    }
    
    .stTabs [aria-selected="true"] {
        background-color: #d4a373 !important;
        color: #1a1a1a !important;
    }
    
    /* Карточки блюд */
    .dish-card {
        background: linear-gradient(135deg, #3d3a38, #2c2a28);
        border-radius: 20px;
        padding: 15px;
        margin: 10px 0;
        box-shadow: 0 2px 10px rgba(0,0,0,0.3);
        transition: all 0.3s ease;
        border-left: 5px solid #d4a373;
        color: white;
    }
    
    .dish-card strong {
        color: #d4a373;
    }
    
    .dish-card:hover {
        transform: translateX(5px);
        box-shadow: 0 4px 20px rgba(0,0,0,0.4);
    }
    
    /* Алерты */
    .stAlert {
        border-radius: 15px;
        border-left: 5px solid;
        background-color: #2a2a2a !important;
        color: white !important;
    }
    
    .stAlert div {
        color: white !important;
    }
    
    /* Информационная панель */
    .stInfo {
        background-color: #2a2a2a !important;
        color: #d4a373 !important;
    }
    
    /* Сайдбар футер */
    .sidebar-footer {
        position: fixed;
        bottom: 20px;
        left: 20px;
        font-size: 12px;
        color: #888;
        text-align: center;
    }
    
    /* Текст в дашборде */
    .stMarkdown p, .stMarkdown li {
        color: #e0e0e0;
    }
    
    /* Мультиселект */
    .stMultiSelect div {
        background-color: #2a2a2a;
        color: white;
    }
    
    /* Количество гостей и другие числа */
    .stNumberInput input {
        color: white !important;
    }
    
    /* Цифры в метриках — гарантированно белые */
    [data-testid="stMetricValue"] {
        color: white !important;
        font-size: 36px !important;
        font-weight: bold !important;
    }
    
    [data-testid="stMetricDelta"] {
        color: #d4a373 !important;
    }
    
    /* Текст в баннере */
    .banner-text {
        color: #e0e0e0 !important;
    }
</style>
""", unsafe_allow_html=True)

# === ДЕКОРАТИВНЫЙ БАННЕР ===
st.markdown("""
<div style="text-align: center; margin-bottom: 30px;">
    <div style="font-size: 60px;">☕🍰🥗</div>
    <h1 style="font-size: 42px; margin: 0; color: #d4a373;">Система учета кафе</h1>
    <p style="color: #b0a090; font-size: 18px; margin-top: 5px;">Ваш помощник в управлении ресторанным бизнесом</p>
</div>
""", unsafe_allow_html=True)

# === БОКОВАЯ ПАНЕЛЬ ===
with st.sidebar:
    st.markdown("### ☕ Меню управления")
    st.markdown("---")
    
    menu = st.selectbox(
        "Выберите раздел",
        ["📊 Дашборд", "📝 Заказы", "🍽️ Меню", "📦 Склад", "📈 Отчеты"],
        format_func=lambda x: x.split(" ")[1] if " " in x else x
    )
    
    st.markdown("---")
    st.markdown("### 📌 Информация")
    st.info("🗄️ Данные сохраняются в файле Кафе_Учет.xlsx\n\n💾 Закройте Excel перед работой с приложением")
    
    st.markdown("---")
    st.markdown('<div class="sidebar-footer">© 2026 | Кафе Учет<br>Версия 2.0</div>', unsafe_allow_html=True)

# === ЗАГРУЗКА ДАННЫХ ===
@st.cache_data
def load_data():
    if not os.path.exists("Кафе_Учет.xlsx"):
        return None, None, None, None
    
    try:
        orders = pd.read_excel("Кафе_Учет.xlsx", sheet_name="ЗАКАЗЫ")
        if "Дата" in orders.columns:
            orders["Дата"] = pd.to_datetime(orders["Дата"], errors="coerce", dayfirst=True)
            orders = orders[orders["Дата"].notna()]
        
        menu_data = pd.read_excel("Кафе_Учет.xlsx", sheet_name="МЕНЮ")
        ingredients = pd.read_excel("Кафе_Учет.xlsx", sheet_name="ИНГРЕДИЕНТЫ")
        recipes = pd.read_excel("Кафе_Учет.xlsx", sheet_name="РЕЦЕПТЫ")
        return orders, menu_data, ingredients, recipes
    except Exception as e:
        st.error(f"Ошибка загрузки данных: {e}")
        return None, None, None, None

orders, menu_data, ingredients, recipes = load_data()

# === ДАШБОРД ===
if menu == "📊 Дашборд" and orders is not None and len(orders) > 0:
    st.markdown("### 📈 Общая статистика")
    
    # Кнопки для выбора периода
    col_period, col_empty = st.columns([2, 3])
    with col_period:
        period = st.radio(
            "📅 Период",
            ["Сегодня", "Неделя", "Месяц", "Всё время"],
            horizontal=True,
            label_visibility="collapsed"
        )
    
    # Фильтруем данные по выбранному периоду
    today = pd.Timestamp.now().date()
    
    if period == "Сегодня":
        filtered_orders = orders[orders["Дата"].dt.date == today]
    elif period == "Неделя":
        week_ago = today - pd.Timedelta(days=7)
        filtered_orders = orders[orders["Дата"].dt.date >= week_ago]
    elif period == "Месяц":
        month_ago = today - pd.Timedelta(days=30)
        filtered_orders = orders[orders["Дата"].dt.date >= month_ago]
    else:
        filtered_orders = orders
    
    # Расчёт показателей
    total_revenue = pd.to_numeric(filtered_orders["Сумма (₸)"], errors="coerce").sum()
    num_orders = len(filtered_orders)
    avg_check = total_revenue / num_orders if num_orders > 0 else 0
    
    # Отображение метрик
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("💰 Выручка", f"{total_revenue:,.0f} ₸", delta=f"за {period.lower()}")
    with col2:
        st.metric("📋 Заказов", num_orders, delta=None)
    with col3:
        st.metric("🧾 Средний чек", f"{avg_check:,.0f} ₸", delta=None)
    with col4:
        st.metric("🎯 Фудкост", "32%", delta="-3%", delta_color="normal")
    
    st.markdown("---")
    
    # Графики в две колонки
    col_left, col_right = st.columns(2)
    
    with col_left:
        st.markdown("### 📅 Продажи по дням")
        if "Дата" in filtered_orders.columns and len(filtered_orders) > 0:
            daily_sales = filtered_orders.groupby(filtered_orders["Дата"].dt.date)["Сумма (₸)"].sum().reset_index()
            if len(daily_sales) > 0:
                fig = px.line(daily_sales, x="Дата", y="Сумма (₸)", 
                             title=f"Динамика выручки ({period})",
                             template="plotly_dark",
                             color_discrete_sequence=["#d4a373"])
                fig.update_layout(plot_bgcolor="rgba(0,0,0,0)", hovermode="x unified", paper_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info(f"Нет продаж за {period.lower()}")
    
    with col_right:
        st.markdown("### 🍽️ Популярные блюда")
        if "Блюда" in filtered_orders.columns and len(filtered_orders) > 0:
            all_dishes = []
            for dishes_str in filtered_orders["Блюда"].dropna():
                for dish in str(dishes_str).split(","):
                    all_dishes.append(dish.strip())
            if all_dishes:
                dish_counts = pd.Series(all_dishes).value_counts().head(10).reset_index()
                dish_counts.columns = ["Блюдо", "Количество"]
                fig = px.bar(dish_counts, x="Количество", y="Блюдо", 
                            orientation='h',
                            title=f"Топ-10 блюд ({period})",
                            template="plotly_dark",
                            color_discrete_sequence=["#d4a373"])
                fig.update_layout(plot_bgcolor="rgba(0,0,0,0)", height=400, paper_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info(f"Нет данных о блюдах за {period.lower()}")
        else:
            st.info("Нет данных о блюдах")

# === Если нет данных ===
elif orders is None or len(orders) == 0:
    st.info("📭 Пока нет данных. Добавьте первый заказ во вкладке 'Заказы'!")

# === ЗАКАЗЫ ===
elif menu == "📝 Заказы" and orders is not None:
    st.markdown("### 🆕 Новый заказ")
    
    col1, col2 = st.columns(2)
    
    with col1:
        with st.container():
            st.markdown('<div class="dish-card">', unsafe_allow_html=True)
            with st.form("new_order", clear_on_submit=True):
                st.markdown("#### 📋 Информация о заказе")
                
                c1, c2 = st.columns(2)
                with c1:
                    date = st.date_input("📅 Дата")
                    table = st.number_input("🍽️ Столик", min_value=1, max_value=20, value=1)
                    waiter = st.selectbox("👨‍🍳 Официант", ["Анна", "Иван", "Елена", "Сергей"])
                with c2:
                    time = st.time_input("⏰ Время")
                    guests = st.number_input("👥 Гостей", min_value=1, value=2)
                    payment = st.selectbox("💳 Оплата", ["Наличные", "Карта", "Безналичный"])
                
                st.markdown("#### 🍕 Выберите блюда")
                if menu_data is not None and len(menu_data) > 0:
                    selected_dishes = st.multiselect(
                        "Блюда", 
                        menu_data["Блюдо"].tolist(),
                        placeholder="Нажмите, чтобы выбрать блюда..."
                    )
                
                note = st.text_area("📝 Примечание к заказу", placeholder="Особые пожелания гостя...")
                
                submitted = st.form_submit_button("✅ Оформить заказ", use_container_width=True)
                
                if submitted and selected_dishes:
                    total = 0
                    for dish in selected_dishes:
                        price = menu_data[menu_data["Блюдо"] == dish]["Цена продажи (₸)"].values
                        if len(price) > 0:
                            total += price[0]
                    
                    new_row = pd.DataFrame([{
                        "Дата": date.strftime("%d.%m.%Y"),
                        "Время": time.strftime("%H:%M"),
                        "Столик": table,
                        "Гости": guests,
                        "Официант": waiter,
                        "Блюда": ", ".join(selected_dishes),
                        "Сумма (₸)": total,
                        "Оплата": payment,
                        "Примечание": note
                    }])
                    
                    if os.path.exists("Кафе_Учет.xlsx"):
                        try:
                            existing = pd.read_excel("Кафе_Учет.xlsx", sheet_name="ЗАКАЗЫ")
                            existing = existing[existing["Дата"].notna()]
                            updated = pd.concat([existing, new_row], ignore_index=True)
                            
                            # Используем безопасную запись
                            if safe_save_to_excel(updated, "ЗАКАЗЫ", mode='replace'):
                                st.success(f"✅ Заказ оформлен! Сумма: {total} ₸")
                                st.balloons()
                                st.cache_data.clear()
                                st.rerun()
                        except Exception as e:
                            st.error(f"Ошибка при сохранении: {e}")
            st.markdown('</div>', unsafe_allow_html=True)
    
    with col2:
        st.markdown("### 📜 Последние заказы")
        if len(orders) > 0:
            display_orders = orders.tail(10).copy()
            if "Дата" in display_orders.columns:
                display_orders["Дата"] = display_orders["Дата"].dt.strftime("%d.%m.%Y")
            st.dataframe(display_orders[["Дата", "Время", "Столик", "Блюда", "Сумма (₸)", "Официант"]], 
                        use_container_width=True, height=400)

# === МЕНЮ ===
elif menu == "🍽️ Меню" and menu_data is not None:
    st.markdown("### 🍽️ Управление меню")
    
    # Форма для добавления нового блюда
    with st.expander("➕ Добавить новое блюдо", expanded=True):
        with st.form("new_dish"):
            col1, col2 = st.columns(2)
            with col1:
                new_dish = st.text_input("🍽️ Название блюда")
                new_ingredients = st.text_input("🥕 Ингредиенты (через запятую)")
            with col2:
                new_cost = st.number_input("📦 Себестоимость (₸)", min_value=0, value=100)
                new_price = st.number_input("💰 Цена продажи (₸)", min_value=0, value=300)
            
            submitted = st.form_submit_button("✅ Сохранить блюдо", use_container_width=True)
            
            if submitted and new_dish:
                margin = (new_price - new_cost) / new_price * 100 if new_price > 0 else 0
                new_row = pd.DataFrame([{
                    "Блюдо": new_dish,
                    "Ингредиенты": new_ingredients,
                    "Себестоимость (₸)": new_cost,
                    "Цена продажи (₸)": new_price,
                    "Маржинальность (%)": margin
                }])
                
                if os.path.exists("Кафе_Учет.xlsx"):
                    try:
                        existing = pd.read_excel("Кафе_Учет.xlsx", sheet_name="МЕНЮ")
                        updated = pd.concat([existing, new_row], ignore_index=True)
                        if safe_save_to_excel(updated, "МЕНЮ", mode='replace'):
                            st.success(f"✅ Блюдо '{new_dish}' добавлено!")
                            st.cache_data.clear()
                            st.rerun()
                    except Exception as e:
                        st.error(f"Ошибка: {e}")
    
    st.markdown("---")
    st.markdown("#### 📋 Текущее меню")
    
    for i, row in menu_data.iterrows():
        try:
            margin = float(row["Маржинальность (%)"])
        except (ValueError, TypeError):
            try:
                cost = float(row["Себестоимость (₸)"])
                price = float(row["Цена продажи (₸)"])
                margin = (price - cost) / price * 100 if price > 0 else 0
            except:
                margin = 0
        
        c1, c2, c3 = st.columns([10, 1, 1])
        
        with c1:
            st.markdown(f"""
            <div class="dish-card" style="margin: 5px 0;">
                <div style="display: flex; justify-content: space-between;">
                    <strong>{row['Блюдо']}</strong>
                    <span style="background: #d4a373; padding: 2px 10px; border-radius: 15px;">{row['Цена продажи (₸)']} ₸</span>
                </div>
                <div style="font-size: 12px;">📦 Себестоимость: {row['Себестоимость (₸)']} ₸ | Маржинальность: {margin:.0f}%</div>
                <div style="font-size: 11px; color: #aaa;">🥕 {row.get('Ингредиенты', '—')}</div>
            </div>
            """, unsafe_allow_html=True)
        
        with c2:
            if st.button("✏️", key=f"edit_{i}"):
                st.info(f"Редактирование {row['Блюдо']} — пока в разработке")
        
        with c3:
            if st.button("🗑️", key=f"del_{i}"):
                menu_data = menu_data[menu_data["Блюдо"] != row['Блюдо']]
                if safe_save_to_excel(menu_data, "МЕНЮ", mode='replace'):
                    st.success(f"🗑️ Блюдо '{row['Блюдо']}' удалено!")
                    st.cache_data.clear()
                    st.rerun()
    
    st.markdown("---")
    st.markdown("#### 📊 Маржинальность блюд")
    
    chart_data = menu_data.copy()
    chart_data["Маржинальность (%)"] = chart_data.apply(
        lambda x: (x["Цена продажи (₸)"] - x["Себестоимость (₸)"]) / x["Цена продажи (₸)"] * 100 
        if x["Цена продажи (₸)"] > 0 else 0,
        axis=1
    )
    
    fig = px.bar(chart_data, x="Блюдо", y="Маржинальность (%)", 
                 title="Маржинальность блюд",
                 color="Маржинальность (%)",
                 color_continuous_scale="RdYlGn",
                 text_auto='.1f')
    fig.update_layout(plot_bgcolor="rgba(0,0,0,0)", height=400, showlegend=False, paper_bgcolor="rgba(0,0,0,0)")
    fig.update_traces(textfont_color="white")
    st.plotly_chart(fig, use_container_width=True)

# === СКЛАД ===
elif menu == "📦 Склад" and ingredients is not None:
    st.markdown("### 📦 Управление складом")
    
    # Форма для добавления нового ингредиента
    with st.expander("➕ Добавить новый ингредиент", expanded=True):
        with st.form("new_ingredient"):
            col1, col2 = st.columns(2)
            with col1:
                new_name = st.text_input("📦 Название ингредиента")
                new_unit = st.selectbox("📏 Единица измерения", ["кг", "г", "л", "мл", "шт", "уп"])
                new_stock = st.number_input("📊 Остаток", min_value=0, value=10)
            with col2:
                new_min = st.number_input("⚠️ Мин. запас", min_value=0, value=2)
                new_price = st.number_input("💰 Цена за ед. (₸)", min_value=0, value=100)
            
            submitted = st.form_submit_button("✅ Добавить ингредиент", use_container_width=True)
            
            if submitted and new_name:
                new_row = pd.DataFrame([{
                    "Ингредиент": new_name,
                    "Единица": new_unit,
                    "Остаток": new_stock,
                    "Мин. запас": new_min,
                    "Цена за ед. (₸)": new_price
                }])
                
                if os.path.exists("Кафе_Учет.xlsx"):
                    try:
                        existing = pd.read_excel("Кафе_Учет.xlsx", sheet_name="ИНГРЕДИЕНТЫ")
                        updated = pd.concat([existing, new_row], ignore_index=True)
                        if safe_save_to_excel(updated, "ИНГРЕДИЕНТЫ", mode='replace'):
                            st.success(f"✅ Ингредиент '{new_name}' добавлен!")
                            st.cache_data.clear()
                            st.rerun()
                    except Exception as e:
                        st.error(f"Ошибка: {e}")
    
    st.markdown("---")
    st.markdown("#### 📋 Остатки ингредиентов")
    
    ingredients["Статус"] = ingredients.apply(
        lambda x: "⚠️ Критический" if x["Остаток"] < x["Мин. запас"] else "✅ Норма", 
        axis=1
    )
    
    for i, row in ingredients.iterrows():
        if row["Статус"] == "⚠️ Критический":
            status_color = "#5a2a2a"
            status_icon = "🔴"
        else:
            status_color = "#2a2a2a"
            status_icon = "🟢"
        
        c1, c2, c3 = st.columns([10, 1, 1])
        
        with c1:
            st.markdown(f"""
            <div class="dish-card" style="background: {status_color}; margin: 5px 0;">
                <div style="display: flex; justify-content: space-between; align-items: center;">
                    <div>
                        <strong style="font-size: 16px;">📦 {row['Ингредиент']}</strong>
                        <span style="margin-left: 10px; font-size: 12px; color: #aaa;">({row['Единица']})</span>
                    </div>
                    <div>
                        <span style="background: #d4a373; color: #1a1a1a; padding: 4px 12px; border-radius: 20px; font-weight: bold;">
                            {row['Остаток']} {row['Единица']}
                        </span>
                    </div>
                </div>
                <div style="margin-top: 8px; font-size: 13px;">
                    📊 Мин. запас: {row['Мин. запас']} {row['Единица']} | 
                    💰 Цена: {row['Цена за ед. (₸)']} ₸/{row['Единица']}
                </div>
                <div style="margin-top: 5px; font-size: 12px;">
                    {status_icon} {row['Статус']}
                </div>
            </div>
            """, unsafe_allow_html=True)
        
        with c2:
            if st.button("📝", key=f"edit_ing_{i}"):
                st.session_state.edit_ingredient = row['Ингредиент']
                st.session_state.new_stock_value = row['Остаток']
        
        with c3:
            if st.button("🗑️", key=f"del_ing_{i}"):
                ingredients = ingredients[ingredients["Ингредиент"] != row['Ингредиент']]
                save_ingredients = ingredients.drop(columns=["Статус"])
                if safe_save_to_excel(save_ingredients, "ИНГРЕДИЕНТЫ", mode='replace'):
                    st.success(f"🗑️ Ингредиент '{row['Ингредиент']}' удалён!")
                    st.cache_data.clear()
                    st.rerun()
    
    if 'edit_ingredient' in st.session_state:
        st.markdown("---")
        st.subheader(f"✏️ Редактирование: {st.session_state.edit_ingredient}")
        
        with st.form("update_stock"):
            new_amount = st.number_input(
                "Новый остаток", 
                min_value=0, 
                value=int(st.session_state.new_stock_value) if st.session_state.new_stock_value else 0
            )
            
            col1, col2 = st.columns(2)
            with col1:
                if st.form_submit_button("✅ Сохранить"):
                    idx = ingredients[ingredients["Ингредиент"] == st.session_state.edit_ingredient].index[0]
                    ingredients.at[idx, "Остаток"] = new_amount
                    save_ingredients = ingredients.drop(columns=["Статус"])
                    
                    if safe_save_to_excel(save_ingredients, "ИНГРЕДИЕНТЫ", mode='replace'):
                        del st.session_state.edit_ingredient
                        del st.session_state.new_stock_value
                        st.success(f"✅ Остаток обновлён: {new_amount}")
                        st.cache_data.clear()
                        st.rerun()
            
            with col2:
                if st.form_submit_button("❌ Отмена"):
                    del st.session_state.edit_ingredient
                    del st.session_state.new_stock_value
                    st.rerun()

# === ОТЧЕТЫ ===
elif menu == "📈 Отчеты" and orders is not None:
    st.markdown("### 📈 Финансовые отчеты")
    
    tab1, tab2, tab3 = st.tabs(["📊 P&L Отчет", "📉 ABC-анализ", "📈 Тренды"])
    
    with tab1:
        revenue = pd.to_numeric(orders["Сумма (₸)"], errors="coerce").sum()
        cost = revenue * 0.32
        gross = revenue - cost
        expenses = 530000
        net = gross - expenses
        
        pl_data = pd.DataFrame({
            "Показатель": ["Выручка", "Себестоимость (32%)", "Валовая прибыль", "Расходы", "Чистая прибыль"],
            "Сумма (₸)": [revenue, cost, gross, expenses, net],
            "% от выручки": ["100%", "32%", f"{gross/revenue*100:.1f}%" if revenue > 0 else "0%", f"{expenses/revenue*100:.1f}%" if revenue > 0 else "0%", f"{net/revenue*100:.1f}%" if revenue > 0 else "0%"]
        })
        
        st.dataframe(pl_data, use_container_width=True, hide_index=True)
        
        fig = go.Figure(data=[go.Pie(labels=pl_data["Показатель"][:4], 
                                     values=pl_data["Сумма (₸)"][:4], 
                                     hole=0.4,
                                     marker=dict(colors=["#d4a373", "#e0b184", "#f5e0c0", "#b5835a"]))])
        fig.update_layout(title="Структура доходов", height=450, plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
        fig.update_traces(textfont_color="white", textposition='inside')
        st.plotly_chart(fig, use_container_width=True)
    
    with tab2:
        if menu_data is not None and "Блюда" in orders.columns:
            menu_with_sales = menu_data.copy()
            sales_count = {}
            for dishes_str in orders["Блюда"].dropna():
                for dish in str(dishes_str).split(","):
                    dish = dish.strip()
                    sales_count[dish] = sales_count.get(dish, 0) + 1
            
            menu_with_sales["Продажи"] = menu_with_sales["Блюдо"].map(sales_count).fillna(0)
            menu_with_sales["Выручка"] = menu_with_sales["Цена продажи (₸)"] * menu_with_sales["Продажи"]
            total_rev = menu_with_sales["Выручка"].sum()
            
            if total_rev > 0:
                menu_with_sales["Доля"] = menu_with_sales["Выручка"] / total_rev * 100
                menu_with_sales.sort_values("Доля", ascending=False, inplace=True)
                menu_with_sales["Категория"] = menu_with_sales["Доля"].apply(lambda x: "A (Звезды)" if x >= 40 else ("B (Середняки)" if x >= 15 else "C (Аутсайдеры)"))
                
                st.dataframe(menu_with_sales[["Блюдо", "Продажи", "Выручка", "Доля", "Категория"]], 
                            use_container_width=True, hide_index=True)
                
                fig = px.bar(menu_with_sales, x="Блюдо", y="Выручка", 
                            color="Категория", title="ABC-анализ по выручке",
                            color_discrete_map={"A (Звезды)": "#2ecc71", "B (Середняки)": "#f39c12", "C (Аутсайдеры)": "#e74c3c"})
                fig.update_layout(plot_bgcolor="rgba(0,0,0,0)", height=450, paper_bgcolor="rgba(0,0,0,0)")
                fig.update_traces(textfont_color="white")
                st.plotly_chart(fig, use_container_width=True)
    
    with tab3:
        if "Дата" in orders.columns and len(orders) > 0:
            daily_sales = orders.groupby(orders["Дата"].dt.date)["Сумма (₸)"].sum().reset_index()
            daily_sales["Скользящее среднее"] = daily_sales["Сумма (₸)"].rolling(window=3, min_periods=1).mean()
            
            fig = go.Figure()
            fig.add_trace(go.Scatter(x=daily_sales["Дата"], y=daily_sales["Сумма (₸)"], 
                                     mode='lines+markers', name='Ежедневная выручка',
                                     line=dict(color='#d4a373', width=3),
                                     marker=dict(size=8, color='#d4a373')))
            fig.add_trace(go.Scatter(x=daily_sales["Дата"], y=daily_sales["Скользящее среднее"], 
                                     mode='lines', name='Тренд (3 дня)',
                                     line=dict(color='#e74c3c', width=2, dash='dash')))
            fig.update_layout(title="Динамика продаж", 
                             xaxis_title="Дата", yaxis_title="Выручка (₸)",
                             plot_bgcolor="rgba(0,0,0,0)", hovermode="x unified", height=450,
                             paper_bgcolor="rgba(0,0,0,0)")
            fig.update_xaxes(gridcolor="#444", color="white")
            fig.update_yaxes(gridcolor="#444", color="white")
            st.plotly_chart(fig, use_container_width=True)

# === Если нет данных ===
elif orders is None or len(orders) == 0:
    st.info("📭 Пока нет данных. Добавьте первый заказ во вкладке 'Заказы'!")

st.markdown('<p style="text-align: center; color: #888; font-size: 12px;">☕ Работает на Streamlit | Данные в файле Кафе_Учет.xlsx</p>', unsafe_allow_html=True)