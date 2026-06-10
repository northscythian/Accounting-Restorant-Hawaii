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
    max_attempts = 3
    for attempt in range(max_attempts):
        try:
            with pd.ExcelWriter("Кафе_Учет.xlsx", engine="openpyxl", mode='a', if_sheet_exists=mode) as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
            return True
        except PermissionError:
            if attempt < max_attempts - 1:
                time.sleep(1)
                continue
            else:
                st.error("❌ Файл Excel заблокирован. Закройте файл 'Кафе_Учет.xlsx' и попробуйте снова.")
                return False
        except:
            try:
                with pd.ExcelWriter("Кафе_Учет.xlsx", engine="openpyxl", mode='w') as writer:
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                return True
            except:
                return False
    return False

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
    .category-tab { font-size: 18px; font-weight: bold; }
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
        ["📊 Дашборд", "📝 Заказы", "🍽️ Меню", "📦 Склад", "📈 Отчеты"],
        format_func=lambda x: x.split(" ")[1] if " " in x else x
    )
    st.markdown("---")
    st.info("🗄️ Данные в файле Кафе_Учет.xlsx\n\n💾 Закройте Excel перед работой")
    st.markdown("---")
    st.markdown('<div style="font-size: 12px; color: #888;">© 2026 | Кафе Учет</div>', unsafe_allow_html=True)

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
        st.error(f"Ошибка: {e}")
        return None, None, None, None

orders, menu_data, ingredients, recipes = load_data()

# === ДАШБОРД ===
if menu == "📊 Дашборд" and orders is not None and len(orders) > 0:
    st.markdown("### 📈 Общая статистика")
    period = st.radio("📅 Период", ["Сегодня", "Неделя", "Месяц", "Всё время"], horizontal=True)
    today = pd.Timestamp.now().date()
    if period == "Сегодня":
        filtered = orders[orders["Дата"].dt.date == today]
    elif period == "Неделя":
        filtered = orders[orders["Дата"].dt.date >= today - pd.Timedelta(days=7)]
    elif period == "Месяц":
        filtered = orders[orders["Дата"].dt.date >= today - pd.Timedelta(days=30)]
    else:
        filtered = orders
    revenue = filtered["Сумма (₸)"].sum()
    num = len(filtered)
    avg = revenue / num if num > 0 else 0
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("💰 Выручка", f"{revenue:,.0f} ₸")
    c2.metric("📋 Заказов", num)
    c3.metric("🧾 Средний чек", f"{avg:,.0f} ₸")
    c4.metric("🎯 Фудкост", "32%")
    
    col_left, col_right = st.columns(2)
    with col_left:
        st.markdown("### 📅 Продажи по дням")
        if len(filtered) > 0:
            daily = filtered.groupby(filtered["Дата"].dt.date)["Сумма (₸)"].sum().reset_index()
            fig = px.line(daily, x="Дата", y="Сумма (₸)", template="plotly_dark", color_discrete_sequence=["#d4a373"])
            fig.update_layout(plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig, use_container_width=True)
    with col_right:
        st.markdown("### 🍽️ Популярные блюда")
        all_dishes = []
        for d in filtered["Блюда"].dropna():
            for dish in str(d).split(","):
                all_dishes.append(dish.strip())
        if all_dishes:
            counts = pd.Series(all_dishes).value_counts().head(10).reset_index()
            counts.columns = ["Блюдо", "Количество"]
            fig = px.bar(counts, x="Количество", y="Блюдо", orientation='h', template="plotly_dark", color_discrete_sequence=["#d4a373"])
            fig.update_layout(plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)", height=400)
            st.plotly_chart(fig, use_container_width=True)

# === ЗАКАЗЫ ===
elif menu == "📝 Заказы" and orders is not None:
    st.markdown("### 🆕 Новый заказ")
    
    if 'cart' not in st.session_state:
        st.session_state.cart = {}
    
    # Загружаем структуру подкатегорий
    @st.cache_data
    def load_subcategories():
        try:
            df = pd.read_excel("Кафе_Учет.xlsx", sheet_name="ПОДКАТЕГОРИИ")
            return df
        except:
            return None
    
    subcategories_df = load_subcategories()
    
    col1, col2 = st.columns(2)
    
    with col1:
        with st.container():
            st.markdown('<div class="dish-card">', unsafe_allow_html=True)
            st.markdown("#### 📋 Информация о заказе")
            
            c1, c2 = st.columns(2)
            with c1:
                date = st.date_input("📅 Дата")
                table = st.number_input("🍽️ Столик", min_value=1, value=1)
                waiter = st.selectbox("👨‍🍳 Официант", ["Анна", "Иван", "Елена", "Сергей"])
            with c2:
                time = st.time_input("⏰ Время")
                guests = st.number_input("👥 Гостей", min_value=1, value=2)
                payment = st.selectbox("💳 Оплата", ["Наличные", "Карта"])
            
            st.markdown("---")
            st.markdown("#### 🍕 Выберите блюда")
            
            if subcategories_df is not None and len(subcategories_df) > 0:
                # Получаем уникальные основные категории
                main_categories = subcategories_df["Основная категория"].unique().tolist()
                
                # Создаём вкладки для основных категорий
                tabs = st.tabs(main_categories)
                
                for tab_idx, main_cat in enumerate(main_categories):
                    with tabs[tab_idx]:
                        # Фильтруем подкатегории для этой основной категории
                        subcategories = subcategories_df[subcategories_df["Основная категория"] == main_cat]["Подкатегория"].unique().tolist()
                        
                        # Для каждой подкатегории создаём свой блок
                        for subcat in subcategories:
                            st.markdown(f"**📌 {subcat}**")
                            dishes_df = subcategories_df[(subcategories_df["Основная категория"] == main_cat) & 
                                                          (subcategories_df["Подкатегория"] == subcat)]
                            
                            for idx, row in dishes_df.iterrows():
                                dish = row["Блюдо"]
                                
                                # Получаем цену из меню
                                if menu_data is not None:
                                    price_row = menu_data[menu_data["Блюдо"] == dish]
                                    if len(price_row) > 0:
                                        price = price_row["Цена продажи (₸)"].values[0]
                                    else:
                                        price = 0
                                else:
                                    price = 0
                                
                                current_qty = st.session_state.cart.get(dish, 0)
                                
                                # Строка с блюдом и кнопками
                                col_a, col_b, col_c, col_d = st.columns([4, 1, 1, 1])
                                with col_a:
                                    st.markdown(f"**{dish}** — {price} ₸")
                                with col_b:
                                    if st.button("➖", key=f"minus_{main_cat}_{subcat}_{idx}"):
                                        if current_qty > 0:
                                            st.session_state.cart[dish] = current_qty - 1
                                            if st.session_state.cart[dish] == 0:
                                                del st.session_state.cart[dish]
                                        st.rerun()
                                with col_c:
                                    st.markdown(f"<div style='text-align: center; font-size: 18px; font-weight: bold;'>{current_qty}</div>", unsafe_allow_html=True)
                                with col_d:
                                    if st.button("➕", key=f"plus_{main_cat}_{subcat}_{idx}"):
                                        st.session_state.cart[dish] = current_qty + 1
                                        st.rerun()
                                st.markdown("---")
            else:
                # Если нет структуры подкатегорий — показываем простой список
                st.warning("Структура меню не загружена. Сначала добавьте подкатегории.")
                
                if menu_data is not None:
                    for idx, row in menu_data.iterrows():
                        dish = row["Блюдо"]
                        price = row["Цена продажи (₸)"]
                        current_qty = st.session_state.cart.get(dish, 0)
                        
                        col_a, col_b, col_c, col_d = st.columns([4, 1, 1, 1])
                        with col_a:
                            st.markdown(f"**{dish}** — {price} ₸")
                        with col_b:
                            if st.button("➖", key=f"minus_simple_{idx}"):
                                if current_qty > 0:
                                    st.session_state.cart[dish] = current_qty - 1
                                    if st.session_state.cart[dish] == 0:
                                        del st.session_state.cart[dish]
                                st.rerun()
                        with col_c:
                            st.markdown(f"<div style='text-align: center; font-size: 18px;'>{current_qty}</div>", unsafe_allow_html=True)
                        with col_d:
                            if st.button("➕", key=f"plus_simple_{idx}"):
                                st.session_state.cart[dish] = current_qty + 1
                                st.rerun()
                        st.markdown("---")
            
            note = st.text_area("📝 Примечание")
            
            col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
            with col_btn2:
                submitted = st.button("✅ Оформить заказ", use_container_width=True)
            
            if submitted and st.session_state.cart:
                total = 0
                dish_details = []
                for dish, qty in st.session_state.cart.items():
                    price = menu_data[menu_data["Блюдо"] == dish]["Цена продажи (₸)"].values[0]
                    total += price * qty
                    if qty > 1:
                        dish_details.append(f"{qty}x {dish}")
                    else:
                        dish_details.append(dish)
                
                new_row = pd.DataFrame([{
                    "Дата": date.strftime("%d.%m.%Y"),
                    "Время": time.strftime("%H:%M"),
                    "Столик": table,
                    "Гости": guests,
                    "Официант": waiter,
                    "Блюда": ", ".join(dish_details),
                    "Сумма (₸)": total,
                    "Оплата": payment,
                    "Примечание": note
                }])
                
                try:
                    existing = pd.read_excel("Кафе_Учет.xlsx", sheet_name="ЗАКАЗЫ")
                    existing = existing[existing["Дата"].notna()]
                    updated = pd.concat([existing, new_row], ignore_index=True)
                    if safe_save_to_excel(updated, "ЗАКАЗЫ", mode='replace'):
                        st.success(f"✅ Заказ оформлен! Сумма: {total} ₸")
                        st.balloons()
                        st.session_state.cart = {}
                        st.cache_data.clear()
                        st.rerun()
                except Exception as e:
                    st.error(f"Ошибка: {e}")
            
            if st.session_state.cart and st.button("🗑️ Очистить корзину"):
                st.session_state.cart = {}
                st.rerun()
            
            st.markdown('</div>', unsafe_allow_html=True)
    
    with col2:
        st.markdown("### 📜 Последние заказы")
        if len(orders) > 0:
            disp = orders.tail(10).copy()
            if "Дата" in disp.columns:
                disp["Дата"] = disp["Дата"].dt.strftime("%d.%m.%Y")
            st.dataframe(disp[["Дата", "Время", "Столик", "Блюда", "Сумма (₸)", "Официант"]], use_container_width=True, height=400)
        
        if st.session_state.cart:
            st.markdown("---")
            st.markdown("### 🛒 Текущий заказ")
            total_preview = 0
            for dish, qty in st.session_state.cart.items():
                price = menu_data[menu_data["Блюдо"] == dish]["Цена продажи (₸)"].values[0]
                subtotal = price * qty
                total_preview += subtotal
                st.markdown(f"- **{dish}**: {qty} шт × {price} ₸ = {subtotal} ₸")
            st.markdown(f"**Итого: {total_preview} ₸**")

# === МЕНЮ ===
elif menu == "🍽️ Меню" and menu_data is not None:
    st.markdown("### 🍽️ Управление меню")
    
    # Функция для загрузки полной структуры
    @st.cache_data
    def load_full_structure():
        try:
            df = pd.read_excel("Кафе_Учет.xlsx", sheet_name="КАТЕГОРИИ_ПОЛНАЯ")
            return df
        except:
            return pd.DataFrame(columns=["Основная категория", "Подкатегория", "Блюдо", "Цена (₸)", "Себестоимость (₸)"])
    
    # Функция для сохранения полной структуры
    def save_full_structure(df):
        with pd.ExcelWriter("Кафе_Учет.xlsx", engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            df.to_excel(writer, sheet_name="КАТЕГОРИИ_ПОЛНАЯ", index=False)
        st.cache_data.clear()
    
    # Загружаем структуру
    structure_df = load_full_structure()
    
    # ===== УПРАВЛЕНИЕ КАТЕГОРИЯМИ =====
    with st.expander("🏷️ Управление категориями и подкатегориями", expanded=False):
        # Получаем список основных категорий
        main_cats = structure_df["Основная категория"].unique().tolist()
        
        # Колонки: добавление категории и добавление подкатегории
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
                        "Себестоимость (₸)": 0
                    }])
                    structure_df = pd.concat([structure_df, new_row], ignore_index=True)
                    save_full_structure(structure_df)
                    st.success(f"✅ Категория '{new_main_cat}' добавлена!")
                    st.rerun()
        
        with col_cat2:
            st.markdown("**📁 Новая подкатегория**")
            selected_main = st.selectbox("Выберите основную категорию", main_cats, key="select_main_for_sub")
            new_subcat = st.text_input("Название подкатегории", key="new_subcat")
            if st.button("✅ Добавить подкатегорию", key="add_subcat"):
                if new_subcat and selected_main:
                    # Проверяем, есть ли уже такая подкатегория
                    existing = structure_df[(structure_df["Основная категория"] == selected_main) & 
                                             (structure_df["Подкатегория"] == new_subcat)]
                    if len(existing) == 0:
                        new_row = pd.DataFrame([{
                            "Основная категория": selected_main,
                            "Подкатегория": new_subcat,
                            "Блюдо": "",
                            "Цена (₸)": 0,
                            "Себестоимость (₸)": 0
                        }])
                        structure_df = pd.concat([structure_df, new_row], ignore_index=True)
                        save_full_structure(structure_df)
                        st.success(f"✅ Подкатегория '{new_subcat}' добавлена в '{selected_main}'!")
                        st.rerun()
                    else:
                        st.warning("Такая подкатегория уже существует")
        
        st.markdown("---")
        st.markdown("**📋 Текущая структура**")
        
        # Отображаем текущую структуру с возможностью удаления
        for main_cat in sorted(main_cats):
            st.markdown(f"### 📌 {main_cat}")
            subcats = structure_df[structure_df["Основная категория"] == main_cat]["Подкатегория"].dropna().unique()
            subcats = [s for s in subcats if s != ""]
            
            for subcat in subcats:
                col_del1, col_del2 = st.columns([4, 1])
                with col_del1:
                    st.markdown(f"   📁 **{subcat}**")
                with col_del2:
                    if st.button("🗑️", key=f"del_subcat_{main_cat}_{subcat}"):
                        # Удаляем подкатегорию (и все блюда в ней)
                        structure_df = structure_df[~((structure_df["Основная категория"] == main_cat) & 
                                                       (structure_df["Подкатегория"] == subcat))]
                        save_full_structure(structure_df)
                        st.success(f"Подкатегория '{subcat}' удалена!")
                        st.rerun()
            
            # Кнопка удаления основной категории
            col_del_cat1, col_del_cat2 = st.columns([4, 1])
            with col_del_cat2:
                if st.button("🗑️ Удалить категорию", key=f"del_main_cat_{main_cat}"):
                    structure_df = structure_df[structure_df["Основная категория"] != main_cat]
                    save_full_structure(structure_df)
                    st.success(f"Категория '{main_cat}' удалена!")
                    st.rerun()
            st.markdown("---")
    
    # ===== ДОБАВЛЕНИЕ БЛЮДА =====
    with st.expander("➕ Добавить новое блюдо", expanded=True):
        with st.form("new_order", clear_on_submit=True):
            col1, col2 = st.columns(2)
            with col1:
                # Выбираем основную категорию
                main_cats_list = structure_df["Основная категория"].unique().tolist()
                selected_main_dish = st.selectbox("Основная категория", main_cats_list, key="dish_main_cat")
                
                # Выбираем подкатегорию (только для выбранной основной)
                subcats_list = structure_df[structure_df["Основная категория"] == selected_main_dish]["Подкатегория"].dropna().unique()
                subcats_list = [s for s in subcats_list if s != ""]
                selected_subcat = st.selectbox("Подкатегория", subcats_list, key="dish_subcat")
                
                new_dish = st.text_input("🍽️ Название блюда")
                new_ingredients = st.text_input("🥕 Ингредиенты (через запятую)")
            with col2:
                new_cost = st.number_input("📦 Себестоимость (₸)", min_value=0, value=100)
                new_price = st.number_input("💰 Цена продажи (₸)", min_value=0, value=300)
            
            submitted = st.form_submit_button("✅ Сохранить блюдо", use_container_width=True)
            
            if submitted and new_dish and selected_main_dish and selected_subcat:
                # Добавляем в структуру
                new_row = pd.DataFrame([{
                    "Основная категория": selected_main_dish,
                    "Подкатегория": selected_subcat,
                    "Блюдо": new_dish,
                    "Цена (₸)": new_price,
                    "Себестоимость (₸)": new_cost
                }])
                structure_df = pd.concat([structure_df, new_row], ignore_index=True)
                save_full_structure(structure_df)
                
                # Добавляем в лист МЕНЮ (для совместимости с отчётами)
                margin = (new_price - new_cost) / new_price * 100 if new_price > 0 else 0
                menu_row = pd.DataFrame([{
                    "Блюдо": new_dish,
                    "Категория": selected_main_dish,
                    "Ингредиенты": new_ingredients,
                    "Себестоимость (₸)": new_cost,
                    "Цена продажи (₸)": new_price,
                    "Маржинальность (%)": margin
                }])
                
                existing_menu = pd.read_excel("Кафе_Учет.xlsx", sheet_name="МЕНЮ")
                updated_menu = pd.concat([existing_menu, menu_row], ignore_index=True)
                safe_save_to_excel(updated_menu, "МЕНЮ", mode='replace')
                
                st.success(f"✅ Блюдо '{new_dish}' добавлено в '{selected_main_dish} → {selected_subcat}'!")
                st.rerun()
    
    st.markdown("---")
    st.markdown("#### 📋 Полное меню")
    
    # Отображаем всё меню
    for main_cat in sorted(structure_df["Основная категория"].unique()):
        st.markdown(f"### 📌 {main_cat}")
        
        subcats = structure_df[structure_df["Основная категория"] == main_cat]["Подкатегория"].dropna().unique()
        subcats = [s for s in subcats if s != ""]
        
        for subcat in subcats:
            st.markdown(f"####    📁 {subcat}")
            dishes_df = structure_df[(structure_df["Основная категория"] == main_cat) & 
                                      (structure_df["Подкатегория"] == subcat) &
                                      (structure_df["Блюдо"] != "")]
            
            for idx, row in dishes_df.iterrows():
                dish = row["Блюдо"]
                price = row["Цена (₸)"] if row["Цена (₸)"] else 0
                cost = row["Себестоимость (₸)"] if row["Себестоимость (₸)"] else 0
                margin = (price - cost) / price * 100 if price > 0 else 0
                
                col_d1, col_d2, col_d3, col_d4 = st.columns([5, 2, 1, 1])
                with col_d1:
                    st.markdown(f"🍽️ **{dish}**")
                with col_d2:
                    st.markdown(f"{price} ₸ | {margin:.0f}%")
                with col_d3:
                    if st.button("✏️", key=f"edit_full_{idx}"):
                        st.session_state.edit_dish_full = dish
                        st.session_state.edit_price_full = price
                        st.session_state.edit_cost_full = cost
                with col_d4:
                    if st.button("🗑️", key=f"del_full_{idx}"):
                        structure_df = structure_df[~((structure_df["Основная категория"] == main_cat) & 
                                                       (structure_df["Подкатегория"] == subcat) &
                                                       (structure_df["Блюдо"] == dish))]
                        save_full_structure(structure_df)
                        st.rerun()
                st.markdown("---")
    
    # Редактирование цены
    if 'edit_dish_full' in st.session_state:
        st.markdown("---")
        st.subheader(f"✏️ Редактирование: {st.session_state.edit_dish_full}")
        with st.form("edit_full_form"):
            new_price = st.number_input("Цена (₸)", min_value=0, value=int(st.session_state.edit_price_full))
            new_cost = st.number_input("Себестоимость (₸)", min_value=0, value=int(st.session_state.edit_cost_full))
            if st.form_submit_button("✅ Сохранить"):
                idx = structure_df[structure_df["Блюдо"] == st.session_state.edit_dish_full].index[0]
                structure_df.at[idx, "Цена (₸)"] = new_price
                structure_df.at[idx, "Себестоимость (₸)"] = new_cost
                save_full_structure(structure_df)
                
                # Обновляем также в листе МЕНЮ
                menu_data = pd.read_excel("Кафе_Учет.xlsx", sheet_name="МЕНЮ")
                if st.session_state.edit_dish_full in menu_data["Блюдо"].values:
                    menu_idx = menu_data[menu_data["Блюдо"] == st.session_state.edit_dish_full].index[0]
                    menu_data.at[menu_idx, "Цена продажи (₸)"] = new_price
                    menu_data.at[menu_idx, "Себестоимость (₸)"] = new_cost
                    menu_data.at[menu_idx, "Маржинальность (%)"] = (new_price - new_cost) / new_price * 100 if new_price > 0 else 0
                    safe_save_to_excel(menu_data, "МЕНЮ", mode='replace')
                
                del st.session_state.edit_dish_full
                st.success("✅ Изменения сохранены!")
                st.rerun()
    
    # ===== ФОРМА ДЛЯ ДОБАВЛЕНИЯ НОВОГО БЛЮДА =====
    with st.expander("➕ Добавить новое блюдо", expanded=True):
        with st.form("new_dish"):
            col1, col2 = st.columns(2)
            with col1:
                new_dish = st.text_input("🍽️ Название блюда")
                new_category = st.selectbox("📂 Категория", existing_categories)
                new_ingredients = st.text_input("🥕 Ингредиенты (через запятую)")
            with col2:
                new_cost = st.number_input("📦 Себестоимость (₸)", min_value=0, value=100)
                new_price = st.number_input("💰 Цена продажи (₸)", min_value=0, value=300)
            
            submitted = st.form_submit_button("✅ Сохранить блюдо", use_container_width=True)
            
            if submitted and new_dish:
                margin = (new_price - new_cost) / new_price * 100 if new_price > 0 else 0
                new_row = pd.DataFrame([{
                    "Блюдо": new_dish,
                    "Категория": new_category,
                    "Ингредиенты": new_ingredients,
                    "Себестоимость (₸)": new_cost,
                    "Цена продажи (₸)": new_price,
                    "Маржинальность (%)": margin
                }])
                
                if os.path.exists("Кафе_Учет.xlsx"):
                    try:
                        existing = pd.read_excel("Кафе_Учет.xlsx", sheet_name="МЕНЮ")
                        if "Категория" not in existing.columns:
                            existing["Категория"] = "Напитки"
                        updated = pd.concat([existing, new_row], ignore_index=True)
                        if safe_save_to_excel(updated, "МЕНЮ", mode='replace'):
                            st.success(f"✅ Блюдо '{new_dish}' добавлено!")
                            st.cache_data.clear()
                            st.rerun()
                    except Exception as e:
                        st.error(f"Ошибка: {e}")
    
    st.markdown("---")
    st.markdown("#### 📋 Текущее меню")
    
    # Обновляем список категорий на случай, если добавили новые
    existing_categories = load_categories()
    
    for i, row in menu_data.iterrows():
        try:
            margin = float(row["Маржинальность (%)"])
        except:
            try:
                cost = float(row["Себестоимость (₸)"])
                price = float(row["Цена продажи (₸)"])
                margin = (price - cost) / price * 100 if price > 0 else 0
            except:
                margin = 0
        
        current_cat = row.get("Категория", "Напитки")
        if current_cat not in existing_categories:
            existing_categories.append(current_cat)
            save_categories(existing_categories)
        
        # Карточка блюда с выбором категории
        c1, c2, c3, c4 = st.columns([5, 2, 1, 1])
        
        with c1:
            st.markdown(f"""
            <div style="background: #2a2a2a; border-radius: 10px; padding: 10px; margin: 5px 0;">
                <strong style="color: #d4a373;">{row['Блюдо']}</strong><br>
                <span style="font-size: 12px;">📦 {row['Себестоимость (₸)']} ₸ | {margin:.0f}%</span>
            </div>
            """, unsafe_allow_html=True)
        
        with c2:
            # Выбор категории
            cat_index = existing_categories.index(current_cat) if current_cat in existing_categories else 0
            new_cat = st.selectbox(
                "Категория",
                existing_categories,
                index=cat_index,
                key=f"cat_{i}",
                label_visibility="collapsed"
            )
            if new_cat != current_cat:
                menu_data.at[i, "Категория"] = new_cat
                if safe_save_to_excel(menu_data, "МЕНЮ", mode='replace'):
                    st.cache_data.clear()
                    st.rerun()
        
        with c3:
            if st.button("✏️", key=f"edit_{i}"):
                st.session_state.edit_dish = row['Блюдо']
                st.session_state.edit_price = row['Цена продажи (₸)']
                st.session_state.edit_cost = row['Себестоимость (₸)']
        
        with c4:
            if st.button("🗑️", key=f"del_{i}"):
                menu_data = menu_data[menu_data["Блюдо"] != row['Блюдо']]
                if safe_save_to_excel(menu_data, "МЕНЮ", mode='replace'):
                    st.success(f"🗑️ Блюдо '{row['Блюдо']}' удалено!")
                    st.cache_data.clear()
                    st.rerun()
    
    # Форма для редактирования цены/себестоимости
    if 'edit_dish' in st.session_state:
        st.markdown("---")
        st.subheader(f"✏️ Редактирование: {st.session_state.edit_dish}")
        with st.form("edit_dish_form"):
            new_price = st.number_input("💰 Новая цена продажи (₸)", min_value=0, value=int(st.session_state.edit_price))
            new_cost = st.number_input("📦 Новая себестоимость (₸)", min_value=0, value=int(st.session_state.edit_cost))
            if st.form_submit_button("✅ Сохранить изменения"):
                idx = menu_data[menu_data["Блюдо"] == st.session_state.edit_dish].index[0]
                menu_data.at[idx, "Цена продажи (₸)"] = new_price
                menu_data.at[idx, "Себестоимость (₸)"] = new_cost
                menu_data.at[idx, "Маржинальность (%)"] = (new_price - new_cost) / new_price * 100 if new_price > 0 else 0
                if safe_save_to_excel(menu_data, "МЕНЮ", mode='replace'):
                    del st.session_state.edit_dish
                    del st.session_state.edit_price
                    del st.session_state.edit_cost
                    st.success("✅ Изменения сохранены!")
                    st.cache_data.clear()
                    st.rerun()
    
    st.markdown("---")
    st.markdown("#### 📊 Маржинальность блюд по категориям")
    
    chart_data = menu_data.copy()
    chart_data["Маржинальность (%)"] = chart_data.apply(
        lambda x: (x["Цена продажи (₸)"] - x["Себестоимость (₸)"]) / x["Цена продажи (₸)"] * 100 if x["Цена продажи (₸)"] > 0 else 0,
        axis=1
    )
    
    if "Категория" in chart_data.columns:
        fig = px.bar(chart_data, x="Блюдо", y="Маржинальность (%)", color="Категория",
                     title="Маржинальность блюд", text_auto='.1f')
    else:
        fig = px.bar(chart_data, x="Блюдо", y="Маржинальность (%)",
                     title="Маржинальность блюд", text_auto='.1f', color_discrete_sequence=["#d4a373"])
    
    fig.update_layout(plot_bgcolor="rgba(0,0,0,0)", height=500, paper_bgcolor="rgba(0,0,0,0)")
    fig.update_traces(textfont_color="white")
    st.plotly_chart(fig, use_container_width=True)

# === СКЛАД ===
elif menu == "📦 Склад" and ingredients is not None:
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
                new = pd.DataFrame([{"Ингредиент": name, "Единица": unit, "Остаток": stock, "Мин. запас": min_stock, "Цена за ед. (₸)": price}])
                existing = pd.read_excel("Кафе_Учет.xlsx", sheet_name="ИНГРЕДИЕНТЫ")
                updated = pd.concat([existing, new], ignore_index=True)
                if safe_save_to_excel(updated, "ИНГРЕДИЕНТЫ", mode='replace'):
                    st.rerun()
    
    ingredients["Статус"] = ingredients.apply(lambda x: "⚠️" if x["Остаток"] < x["Мин. запас"] else "✅", axis=1)
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
            ingredients = ingredients[ingredients["Ингредиент"] != row['Ингредиент']]
            save = ingredients.drop(columns=["Статус"])
            if safe_save_to_excel(save, "ИНГРЕДИЕНТЫ", mode='replace'):
                st.rerun()

# === ОТЧЕТЫ ===
elif menu == "📈 Отчеты" and orders is not None:
    st.markdown("### 📈 Отчеты")
    rev = orders["Сумма (₸)"].sum()
    st.metric("Выручка", f"{rev:,.0f} ₸")
    st.dataframe(orders.tail(20), use_container_width=True)

elif orders is None or len(orders) == 0:
    st.info("📭 Нет данных. Добавьте первый заказ.")

st.markdown('<p style="text-align: center; color: #888; font-size: 12px;">☕ Работает на Streamlit</p>', unsafe_allow_html=True)