import streamlit as st
import pandas as pd
import uuid
from datetime import datetime
import requests
import time
import json
import re
import os

# =====================================================================
# НАСТРОЙКА СТРАНИЦЫ
# =====================================================================
st.set_page_config(
    page_title="Система управления кафе",
    page_icon="☕",
    layout="wide",
    initial_sidebar_state="expanded"
)

# =====================================================================
# КОНСТАНТЫ
# =====================================================================
TELEGRAM_BOT_TOKEN = "8895293528:AAHX6k1TCDBSw_019BvrSU0jQYXyMTek7pY"
TELEGRAM_CHAT_ID = "6017544408"

ORDER_STATUSES = ["Новый", "Готовится", "Готов", "Оплачен", "Отменён"]

FALLBACK_MENU = [
    {"Блюдо": "Маргарита", "Категория": "Пиццы", "Цена продажи (₸)": 2500},
    {"Блюдо": "Диавола", "Категория": "Пиццы", "Цена продажи (₸)": 3200},
    {"Блюдо": "Четыре сыра", "Категория": "Пиццы", "Цена продажи (₸)": 3500},
    {"Блюдо": "Карбонара", "Категория": "Пиццы", "Цена продажи (₸)": 3600},
    {"Блюдо": "Каприччоза", "Категория": "Пиццы", "Цена продажи (₸)": 3800},
    {"Блюдо": "Брускетта", "Категория": "Закуски", "Цена продажи (₸)": 1200},
    {"Блюдо": "Фриттура Миста", "Категория": "Закуски", "Цена продажи (₸)": 2200},
    {"Блюдо": "Капрезе", "Категория": "Салаты", "Цена продажи (₸)": 1800},
    {"Блюдо": "Руккола с пармезаном", "Категория": "Салаты", "Цена продажи (₸)": 1600},
    {"Блюдо": "Лимонад", "Категория": "Напитки", "Цена продажи (₸)": 500},
    {"Блюдо": "Тирамису", "Категория": "Десерты", "Цена продажи (₸)": 750},
    {"Блюдо": "Панна Котта", "Категория": "Десерты", "Цена продажи (₸)": 650},
]

# =====================================================================
# CSS (полная версия с мобильной адаптацией)
# =====================================================================
st.markdown("""
<style>
    /* ===== БАЗОВЫЕ СТИЛИ ===== */
    .main .block-container {
        padding: 1rem !important;
        max-width: 100% !important;
    }
    
    /* ===== КНОПКИ ===== */
    .stButton button {
        min-height: 44px !important;
        font-size: 15px !important;
        border-radius: 24px !important;
        transition: transform 0.08s ease, background 0.15s ease !important;
    }
    .stButton button:active {
        transform: scale(0.94) !important;
    }
    
    /* ===== ПОЛЯ ВВОДА ===== */
    .stTextInput input, 
    .stNumberInput input, 
    .stTextArea textarea, 
    .stSelectbox select {
        font-size: 16px !important;
        padding: 10px !important;
        min-height: 44px !important;
        border-radius: 12px !important;
    }
    
    /* ===== КАРТОЧКИ ===== */
    .order-card { 
        background: #2a2a2a; 
        border-radius: 15px; 
        padding: 15px; 
        margin-bottom: 15px; 
        border-left: 5px solid #d4a373; 
        color: white;
    }
    .dish-card {
        background: linear-gradient(135deg, #3d3a38, #2c2a28);
        border-radius: 20px;
        padding: 15px;
        margin: 10px 0;
        border-left: 5px solid #d4a373;
        color: white;
    }
    
    /* ===== САЙДБАР ===== */
    [data-testid="stSidebar"] {
        background: #1a1a1a;
        border-right: 2px solid #d4a373;
    }
    [data-testid="stSidebar"] * {
        color: #e0e0e0 !important;
    }
    
    /* ===== ЗАГОЛОВКИ ===== */
    h1, h2, h3 {
        color: #d4a373 !important;
    }
    
    /* ===== ТАБЛИЦЫ ===== */
    .stDataFrame {
        border-radius: 12px !important;
        overflow: hidden !important;
    }
    
    /* ===== ТОСТ-УВЕДОМЛЕНИЯ ===== */
    .stAlert {
        border-radius: 12px !important;
        font-size: 14px !important;
    }
    
    /* ================================================================
       МОБИЛЬНАЯ АДАПТАЦИЯ
       ================================================================ */
    
    /* Очень маленькие экраны (iPhone SE, 320-480px) */
    @media (max-width: 480px) {
        .main .block-container {
            padding: 0.5rem !important;
        }
        
        .stButton button {
            font-size: 13px !important;
            min-height: 38px !important;
            padding: 6px 12px !important;
        }
        
        h1 {
            font-size: 20px !important;
        }
        h2 {
            font-size: 17px !important;
        }
        h3 {
            font-size: 15px !important;
        }
        
        .row-widget.stColumns {
            gap: 4px !important;
        }
        
        .stSelectbox select,
        .stTextInput input,
        .stNumberInput input {
            font-size: 14px !important;
            padding: 8px !important;
            min-height: 36px !important;
        }
        
        .stTextArea textarea {
            font-size: 14px !important;
            padding: 8px !important;
            min-height: 60px !important;
        }
        
        /* Уменьшаем отступы в карточках */
        .order-card {
            padding: 10px !important;
            margin-bottom: 10px !important;
        }
        .dish-card {
            padding: 10px !important;
            margin: 6px 0 !important;
        }
        
        /* Меньше отступов в сайдбаре */
        [data-testid="stSidebar"] {
            padding: 0.5rem !important;
        }
        
        /* Убираем лишние отступы у метрик */
        [data-testid="stMetric"] {
            padding: 8px !important;
        }
        [data-testid="stMetric"] label {
            font-size: 12px !important;
        }
        [data-testid="stMetric"] div {
            font-size: 20px !important;
        }
    }
    
    /* Средние экраны (планшеты, 481-768px) */
    @media (min-width: 481px) and (max-width: 768px) {
        .main .block-container {
            padding: 0.8rem !important;
        }
        
        h1 {
            font-size: 28px !important;
        }
        h2 {
            font-size: 22px !important;
        }
        h3 {
            font-size: 18px !important;
        }
        
        .stButton button {
            font-size: 14px !important;
            min-height: 40px !important;
        }
        
        .stSelectbox select,
        .stTextInput input,
        .stNumberInput input {
            font-size: 15px !important;
            padding: 9px !important;
        }
    }
    
    /* Большие экраны (ноутбуки и выше) */
    @media (min-width: 769px) {
        .main .block-container {
            padding: 1.5rem !important;
            max-width: 1200px !important;
            margin: 0 auto !important;
        }
    }
    
    /* ================================================================
       ТЕМНАЯ ТЕМА ДЛЯ САЙДБАРА (опционально)
       ================================================================ */
    .sidebar-dark {
        background: #1a1a1a !important;
    }
    
    /* ================================================================
       АНИМАЦИИ
       ================================================================ */
    @keyframes fadeIn {
        from { opacity: 0; transform: translateY(10px); }
        to { opacity: 1; transform: translateY(0); }
    }
    
    .fade-in {
        animation: fadeIn 0.3s ease;
    }
</style>
""", unsafe_allow_html=True)

# =====================================================================
# РАБОТА С GOOGLE SHEETS
# =====================================================================
@st.cache_data(ttl=600)
def load_sheet_data(worksheet_name: str) -> pd.DataFrame:
    """Загружает данные из Google Sheets с преобразованием типов"""
    try:
        import gspread
        from oauth2client.service_account import ServiceAccountCredentials
        
        if "google" not in st.secrets:
            return pd.DataFrame()
        
        creds_dict = dict(st.secrets["google"])
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        
        workbook_name = st.secrets.get("GOOGLE_SHEET_NAME", "Cafe_System")
        wb = client.open(workbook_name)
        
        try:
            ws = wb.worksheet(worksheet_name)
            records = ws.get_all_records()
            df = pd.DataFrame(records) if records else pd.DataFrame()
            
            # === ПРЕОБРАЗОВАНИЕ ТИПОВ ===
            # Все возможные числовые колонки
            numeric_columns = [
                "Сумма (₸)", "Цена продажи (₸)", "Себестоимость (₸)", 
                "Количество", "Цена за ед. (₸)", "Остаток",
                "Столик", "Гости", "ID", "Сумма"
            ]
            for col in numeric_columns:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
            
            # Преобразование строковых колонок (удаление лишних пробелов)
            string_columns = ["Блюдо", "Категория", "Ингредиент", "Единица", "Статус", "Оплата"]
            for col in string_columns:
                if col in df.columns:
                    df[col] = df[col].astype(str).str.strip()
            
            return df
        except Exception as e:
            st.warning(f"Ошибка загрузки листа {worksheet_name}: {e}")
            return pd.DataFrame()
    except Exception as e:
        st.warning(f"Ошибка подключения к Google Sheets: {e}")
        return pd.DataFrame()

def save_to_google_sheet(df: pd.DataFrame, worksheet_name: str, headers: list = None) -> bool:
    """Сохраняет данные в Google Sheets"""
    try:
        import gspread
        from oauth2client.service_account import ServiceAccountCredentials
        
        if "google" not in st.secrets:
            return False
        
        creds_dict = dict(st.secrets["google"])
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        
        workbook_name = st.secrets.get("GOOGLE_SHEET_NAME", "Cafe_System")
        wb = client.open(workbook_name)
        
        try:
            ws = wb.worksheet(worksheet_name)
            ws.clear()
        except:
            ws = wb.add_worksheet(title=worksheet_name, rows="1000", cols="30")
        
        if headers:
            ws.update("A1", [headers])
            values = df.fillna("").astype(str).values.tolist()
            if values:
                ws.update("A2", values)
        else:
            clean_df = df.fillna("")
            values = [clean_df.columns.tolist()] + clean_df.astype(str).values.tolist()
            if values:
                ws.update("A1", values)
        return True
    except Exception:
        return False

def append_to_google_sheet(row_data: list, worksheet_name: str) -> bool:
    """Добавляет строку в Google Sheets с подробной диагностикой"""
    try:
        import gspread
        from oauth2client.service_account import ServiceAccountCredentials
        
        if "google" not in st.secrets:
            st.error("❌ Нет секретов Google в .streamlit/secrets.toml")
            return False
        
        creds_dict = dict(st.secrets["google"])
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        
        workbook_name = st.secrets.get("GOOGLE_SHEET_NAME", "Cafe_System")
        wb = client.open(workbook_name)
        
        # Получаем или создаём лист
        try:
            ws = wb.worksheet(worksheet_name)
        except Exception:
            ws = wb.add_worksheet(title=worksheet_name, rows="1000", cols="30")
            if worksheet_name == "ЗАКАЗЫ":
                headers = ["ID", "Дата", "Время", "Столик", "Гости", "Официант", 
                          "Блюда", "Сумма (₸)", "Оплата", "Примечание", "Статус", "Создано"]
                ws.update("A1", [headers])
                st.info("📄 Лист 'ЗАКАЗЫ' создан")
        
        # Добавляем строку
        ws.append_row(row_data)
        st.success(f"✅ Заказ добавлен в Google Sheets! Строк в таблице: {ws.row_count}")
        return True
    except gspread.exceptions.APIError as e:
        st.error(f"❌ Ошибка API Google: {e}")
        st.info("💡 Проверьте, что сервисный аккаунт добавлен как редактор таблицы")
        return False
    except Exception as e:
        st.error(f"❌ Ошибка: {e}")
        return False

def send_to_telegram(text: str) -> bool:
    try:
        url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
        resp = requests.post(
            url,
            json={"chat_id": TELEGRAM_CHAT_ID, "text": text, "parse_mode": "HTML"},
            timeout=10,
        )
        return resp.ok
    except Exception:
        return False
        

# =====================================================================
# ИНИЦИАЛИЗАЦИЯ SESSION STATE
# =====================================================================
defaults = {
    "cart": {},
    "waiter_name": "",
    "session_orders": [],
    "current_user": None,
    "confirm_open": False,
}
for key, val in defaults.items():
    if key not in st.session_state:
        st.session_state[key] = val
        
# =====================================================================
# ПРОВЕРКА PIN-КОДА И РОЛЕЙ
# =====================================================================

def authenticate_user(pin_input: str) -> dict:
    """Проверяет PIN-код и возвращает данные сотрудника"""
    staff_df = load_sheet_data("ПЕРСОНАЛ")
    if staff_df.empty:
        return None
    
    staff_df["PIN-код"] = staff_df["PIN-код"].astype(str).str.strip()
    user = staff_df[staff_df["PIN-код"] == pin_input.strip()]
    
    if not user.empty:
        return {
            "id": int(user.iloc[0]["ID"]) if "ID" in user.columns else 1,
            "name": user.iloc[0]["Имя"],
            "role": user.iloc[0]["Должность"],
            "active": user.iloc[0]["Активен"] if "Активен" in user.columns else "Да"
        }
    return None

def get_user_permissions(role: str) -> dict:
    """Возвращает доступные разделы для роли"""
    # Доступ для всех
    base_sections = {
        "📊 Дашборд": True,
        "📝 Заказы (официант)": True,
    }
    
    # Дополнительные разделы для Администратора
    admin_sections = {
        "🍽️ Меню": True,
        "📦 Закупки и склад": True,
        "👥 Персонал": True,
        "📈 Отчеты": True,
    }
    
    if role == "Администратор":
        return {**base_sections, **admin_sections}
    elif role == "Официант":
        return base_sections
    elif role == "Повар":
        return {
            "📊 Дашборд": True,
            "📝 Заказы (официант)": True,
        }
    else:
        return base_sections

# Инициализация состояния входа
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
    st.session_state.user = None

# --- ЭКРАН ВХОДА ---
if not st.session_state.authenticated:
    st.markdown("## 🔐 Вход в систему")
    st.markdown("Введите ваш PIN-код для доступа")
    
    pin_input = st.text_input("PIN-код", type="password", max_chars=4, placeholder="0000")
    
    col1, col2 = st.columns(2)
    with col1:
        if st.button("🔓 Войти", use_container_width=True):
            if pin_input:
                user = authenticate_user(pin_input)
                if user:
                    if user["active"] == "Да":
                        st.session_state.authenticated = True
                        st.session_state.user = user
                        st.success(f"✅ Добро пожаловать, {user['name']}!")
                        time.sleep(0.5)
                        st.rerun()
                    else:
                        st.error("❌ Сотрудник неактивен. Обратитесь к администратору.")
                else:
                    st.error("❌ Неверный PIN-код. Попробуйте ещё раз.")
            else:
                st.warning("⚠️ Введите PIN-код")
    
    with col2:
        st.caption("👤 Сотрудник с правами администратора может управлять персоналом")
    
    st.stop()  # Останавливаем выполнение, пока пользователь не войдёт        

# =====================================================================
# БОКОВОЕ МЕНЮ (с фильтрацией по ролям)
# =====================================================================
with st.sidebar:
    st.markdown("## ☕ Меню управления")
    
    # Информация о пользователе
    if st.session_state.user:
        st.markdown(f"""
        <div style="background: #2a2a2a; padding: 10px; border-radius: 10px; margin-bottom: 15px; border-left: 3px solid #d4a373;">
            <span style="color: #d4a373;">👤 {st.session_state.user['name']}</span><br>
            <span style="color: #888; font-size: 12px;">📌 {st.session_state.user['role']}</span>
        </div>
        """, unsafe_allow_html=True)
        
        if st.button("🚪 Выйти", use_container_width=True):
            st.session_state.authenticated = False
            st.session_state.user = None
            st.rerun()
    
    # Получаем доступные разделы
    permissions = get_user_permissions(st.session_state.user['role'] if st.session_state.user else "Гость")
    menu_sections = [section for section, allowed in permissions.items() if allowed]
    
    if not menu_sections:
        st.warning("У вас нет доступа к разделам")
        st.stop()
    
    selected_section = st.radio("Выберите раздел", menu_sections)
    
    st.markdown("---")
    st.caption("© 2026 | Система учета кафе")

## =====================================================================
# 1. ДАШБОРД
# =====================================================================
if selected_section == "📊 Дашборд":
    st.markdown("## 📊 Дашборд")
    
    orders_df = load_sheet_data("ЗАКАЗЫ")
    user_role = st.session_state.user['role'] if st.session_state.user else ""
    user_name = st.session_state.user['name'] if st.session_state.user else ""
    
    if orders_df.empty:
        st.info("📭 Нет заказов.")
    else:
        # Фильтруем заказы для официанта
        if user_role == "Официант":
            orders_df = orders_df[orders_df["Официант"] == user_name]
            if orders_df.empty:
                st.info(f"📭 У вас ещё нет заказов, {user_name}. Принимайте заказы в разделе «Заказы»!")
                st.stop()
        
        # Обработка суммы
        if "Сумма (₸)" not in orders_df.columns:
            orders_df["Сумма (₸)"] = 0
        
        total_revenue = orders_df["Сумма (₸)"].sum()
        total_orders = len(orders_df)
        
        col1, col2, col3 = st.columns(3)
        col1.metric("💰 Выручка", f"{total_revenue:,.0f} ₸")
        col2.metric("📋 Заказов", total_orders)
        col3.metric("📈 Средний чек", f"{total_revenue/total_orders:,.0f} ₸" if total_orders > 0 else "0 ₸")
        
        # Для администратора — дополнительная статистика
        if user_role == "Администратор":
            col4 = st.columns(1)[0]
            with col4:
                st.metric("👥 Активных сотрудников", len(load_sheet_data("ПЕРСОНАЛ")))
                st.metric("🍽️ Блюд в меню", len(load_sheet_data("МЕНЮ")))

# =====================================================================
# 2. ЗАКАЗЫ (ОФИЦИАНТ)
# =====================================================================
elif selected_section == "📝 Заказы (официант)":
    st.markdown("## 👨‍🍳 Приём заказов")
    
    waiter_name = st.session_state.user['name'] if st.session_state.user else ""
    user_role = st.session_state.user['role'] if st.session_state.user else ""
    
    if not waiter_name:
        st.error("❌ Ошибка: имя не найдено. Перезайдите в систему.")
        st.stop()
    
    st.caption(f"👤 Официант: **{waiter_name}**")
    
    # Загрузка меню
    menu_df = load_sheet_data("МЕНЮ")
    if menu_df.empty:
        menu_df = pd.DataFrame(FALLBACK_MENU)
    
    col_menu, col_cart = st.columns([2, 1])
    
    # =============================================================
    # ЛЕВАЯ КОЛОНКА — МЕНЮ
    # =============================================================
    with col_menu:
        # Вкладки для меню и списка заказов
        tab_menu, tab_orders = st.tabs(["🍽️ Меню", "📋 Мои заказы"])
        
        # --- ВКЛАДКА "МЕНЮ" ---
        with tab_menu:
            categories = menu_df["Категория"].unique().tolist() if "Категория" in menu_df.columns else ["Все"]
            tabs = st.tabs(categories)
            
            for tab_idx, (tab, category) in enumerate(zip(tabs, categories)):
                with tab:
                    items = menu_df[menu_df["Категория"] == category] if "Категория" in menu_df.columns else menu_df
                    if items.empty:
                        st.info("Нет блюд")
                        continue
                    cols = st.columns(3)
                    for idx, (_, row) in enumerate(items.iterrows()):
                        dish = row["Блюдо"]
                        price = int(row.get("Цена продажи (₸)", 0))
                        qty = st.session_state.cart.get(dish, {}).get("quantity", 0)
                        
                        with cols[idx % 3]:
                            st.markdown(f"**{dish}**")
                            st.caption(f"{price:,} ₸")
                            c1, c2 = st.columns(2)
                            with c1:
                                if st.button("➕", key=f"add_{tab_idx}_{idx}"):
                                    if dish in st.session_state.cart:
                                        st.session_state.cart[dish]["quantity"] += 1
                                    else:
                                        st.session_state.cart[dish] = {"name": dish, "price": price, "quantity": 1}
                                    st.rerun()
                            with c2:
                                if st.button("➖", key=f"sub_{tab_idx}_{idx}"):
                                    if dish in st.session_state.cart:
                                        if st.session_state.cart[dish]["quantity"] > 1:
                                            st.session_state.cart[dish]["quantity"] -= 1
                                        else:
                                            del st.session_state.cart[dish]
                                    st.rerun()
        
               # --- ВКЛАДКА "МОИ ЗАКАЗЫ" ---
        with tab_orders:
            st.markdown("### 📋 Мои заказы")
            
            all_orders = load_sheet_data("ЗАКАЗЫ")
            
            if all_orders.empty:
                st.info("📭 Нет заказов")
            else:
                if "Официант" in all_orders.columns:
                    my_orders = all_orders[all_orders["Официант"] == waiter_name].copy()
                else:
                    my_orders = all_orders.copy()
                
                if my_orders.empty:
                    st.info(f"📭 У вас пока нет заказов, {waiter_name}.")
                else:
                    if "Создано" in my_orders.columns:
                        my_orders = my_orders.sort_values("Создано", ascending=False)
                    
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("📋 Всего заказов", len(my_orders))
                    with col2:
                        total_sum = my_orders["Сумма (₸)"].sum() if "Сумма (₸)" in my_orders.columns else 0
                        st.metric("💰 Общая сумма", f"{total_sum:,.0f} ₸")
                    with col3:
                        avg_check = total_sum / len(my_orders) if len(my_orders) > 0 else 0
                        st.metric("📈 Средний чек", f"{avg_check:,.0f} ₸")
                    
                    st.markdown("---")
                    
                    for _, order in my_orders.iterrows():
                        order_id = order.get("ID", "—")
                        table = order.get("Столик", "—")
                        dishes = order.get("Блюда", "—")
                        total = order.get("Сумма (₸)", 0)
                        status = order.get("Статус", "Новый")
                        created = order.get("Создано", "—")
                        comment = order.get("Примечание", "")
                        
                        status_colors = {
                            "Новый": "🔵", "Готовится": "🟡", "Готов": "🟢",
                            "Оплачен": "✅", "Отменён": "🔴"
                        }
                        status_icon = status_colors.get(status, "⚪")
                        
                        st.markdown(f"""
                        <div style="background: #2a2a2a; border-radius: 15px; padding: 12px 16px; margin-bottom: 12px; border-left: 4px solid #d4a373;">
                            <div style="display: flex; justify-content: space-between; align-items: center;">
                                <div>
                                    <span style="font-weight: bold; color: #d4a373;">#{order_id}</span>
                                    <span style="color: #888; font-size: 13px; margin-left: 10px;">🕐 {created}</span>
                                </div>
                                <div>
                                    <span style="background: #3d3a38; padding: 4px 12px; border-radius: 20px; font-size: 13px;">
                                        {status_icon} {status}
                                    </span>
                                </div>
                            </div>
                            <div style="margin-top: 6px; color: #e0e0e0;">
                                🍽️ Стол {table} · 🍕 {dishes}
                            </div>
                            <div style="margin-top: 6px; display: flex; justify-content: space-between; align-items: center;">
                                <span style="color: #d4a373; font-weight: bold;">💰 {total:,.0f} ₸</span>
                                {f'<span style="color: #888; font-size: 12px;">📝 {comment}</span>' if comment else ''}
                            </div>
                        </div>
                        """, unsafe_allow_html=True)
                        
                        # ===== СМЕНА СТАТУСА (для администратора и повара) =====
                        if user_role in ["Администратор", "Повар"]:
                            col_status, col_btn = st.columns([3, 1])
                            with col_status:
                                new_status = st.selectbox(
                                    "Статус",
                                    ORDER_STATUSES,
                                    index=ORDER_STATUSES.index(status) if status in ORDER_STATUSES else 0,
                                    key=f"status_{order_id}",
                                    label_visibility="collapsed"
                                )
                            with col_btn:
                                if new_status != status:
                                    if st.button("✅ Обновить", key=f"update_{order_id}"):
                                        # Обновляем статус в Google Sheets
                                        if update_order_status(order_id, new_status):
                                            st.success(f"✅ Статус #{order_id} → {new_status}")
                                            st.rerun()
                    
                    if st.button("🔄 Обновить список", use_container_width=True):
                        st.rerun()
    # =============================================================
    # ПРАВАЯ КОЛОНКА — КОРЗИНА
    # =============================================================
    with col_cart:
        cart_count = sum(v["quantity"] for v in st.session_state.cart.values())
        st.markdown(f"### 🛒 Корзина ({cart_count})" if cart_count else "### 🛒 Корзина")
        
        table_number = st.number_input("🍽️ Столик", min_value=1, max_value=50, value=1)
        guests = st.number_input("👥 Гостей", min_value=1, max_value=20, value=2)
        payment_method = st.selectbox("💳 Оплата", ["Наличные", "Карта", "Kaspi", "Перевод"])
        comment = st.text_area("📝 Примечание", placeholder="Без лука, острое...")
        
        if not st.session_state.cart:
            st.info("🛒 Корзина пуста")
        else:
            total = 0
            for dish_name, item in list(st.session_state.cart.items()):
                subtotal = item["price"] * item["quantity"]
                total += subtotal
                col1, col2 = st.columns([3, 1])
                with col1:
                    st.markdown(f"**{item['name']}** x{item['quantity']} = {subtotal:,} ₸")
                with col2:
                    if st.button("🗑️", key=f"del_{dish_name}"):
                        del st.session_state.cart[dish_name]
                        st.rerun()
            
            st.markdown("---")
            st.markdown(f"### 💰 ИТОГО: {total:,} ₸")
            
            # ========== КНОПКА ОТПРАВКИ ЗАКАЗА ==========
            if st.button("✅ Отправить заказ", use_container_width=True):
                if not st.session_state.cart:
                    st.error("❌ Корзина пуста!")
                else:
                    # Формируем заказ
                    dishes_list = [f"{v['quantity']}x {v['name']}" for v in st.session_state.cart.values()]
                    dishes_text = ", ".join(dishes_list)
                    now = datetime.now()
                    
                    order_id = uuid.uuid4().hex[:10].upper()
                    
                    row_data = [
                        order_id,
                        now.strftime("%d.%m.%Y"),
                        now.strftime("%H:%M"),
                        table_number,
                        guests,
                        waiter_name,
                        dishes_text,
                        total,
                        payment_method,
                        comment,
                        "Новый",
                        now.strftime("%d.%m.%Y %H:%M")
                    ]
                    
                    telegram_text = f"""
🍕 <b>НОВЫЙ ЗАКАЗ!</b>
👤 <b>Официант:</b> {waiter_name}
🍽️ <b>Столик:</b> {table_number}
👥 <b>Гостей:</b> {guests}

📋 <b>Состав:</b>
{dishes_text}

💰 <b>Сумма:</b> {total:,} ₸
💳 <b>Оплата:</b> {payment_method}
📝 <b>Примечание:</b> {comment or '—'}

🕐 {now.strftime("%d.%m.%Y %H:%M")}
"""
                    
                    with st.spinner("Отправка заказа..."):
                        tg_ok = send_to_telegram(telegram_text)
                        gs_ok = append_to_google_sheet(row_data, "ЗАКАЗЫ")
                    
                    if tg_ok or gs_ok:
                        st.session_state.cart = {}
                        st.success(f"✅ Заказ #{order_id} отправлен!")
                        st.balloons()
                        time.sleep(0.5)
                        st.rerun()
                    else:
                        st.error("❌ Ошибка отправки заказа!")

# =====================================================================
# 3. МЕНЮ (полное управление)
# =====================================================================
elif selected_section == "🍽️ Меню":
    st.markdown("## 🍽️ Управление меню")
    
    menu_df = load_sheet_data("МЕНЮ")
    
    # --- Вкладки: Категории и Блюда ---
    tab_categories, tab_dishes = st.tabs(["📁 Категории", "🍕 Блюда"])
    
    # =============================================================
    # ВКЛАДКА "КАТЕГОРИИ"
    # =============================================================
    with tab_categories:
        st.markdown("### 📁 Управление категориями")
        
        # Получаем текущие категории
        categories = menu_df["Категория"].unique().tolist() if not menu_df.empty else []
        
        col1, col2 = st.columns(2)
        
        # Добавление категории
        with col1:
            st.markdown("#### ➕ Добавить категорию")
            new_category = st.text_input("Название категории", key="new_category_input")
            if st.button("✅ Добавить категорию", key="add_category"):
                if new_category and new_category not in categories:
                    # Добавляем новую категорию с пустым блюдом-заглушкой
                    new_row = pd.DataFrame([{
                        "Блюдо": f"--- {new_category} ---",
                        "Категория": new_category,
                        "Цена продажи (₸)": 0,
                        "Себестоимость (₸)": 0
                    }])
                    updated_df = pd.concat([menu_df, new_row], ignore_index=True) if not menu_df.empty else new_row
                    if save_to_google_sheet(updated_df, "МЕНЮ"):
                        st.success(f"✅ Категория '{new_category}' добавлена!")
                        st.rerun()
                elif new_category in categories:
                    st.warning("⚠️ Такая категория уже существует")
                else:
                    st.warning("⚠️ Введите название категории")
        
        # Удаление категории
        with col2:
            st.markdown("#### 🗑️ Удалить категорию")
            if categories:
                category_to_delete = st.selectbox("Выберите категорию", categories, key="delete_category_select")
                if st.button("🗑️ Удалить категорию", key="delete_category"):
                    if category_to_delete:
                        # Удаляем все блюда в этой категории
                        updated_df = menu_df[menu_df["Категория"] != category_to_delete]
                        if save_to_google_sheet(updated_df, "МЕНЮ"):
                            st.success(f"✅ Категория '{category_to_delete}' удалена вместе со всеми блюдами!")
                            st.rerun()
            else:
                st.info("Нет категорий для удаления")
        
        # Список категорий
        st.markdown("---")
        st.markdown("#### 📋 Текущие категории")
        if categories:
            for cat in sorted(categories):
                st.markdown(f"- 📁 **{cat}**")
        else:
            st.info("Нет категорий")
    
    # =============================================================
    # ВКЛАДКА "БЛЮДА"
    # =============================================================
    with tab_dishes:
        st.markdown("### 🍕 Управление блюдами")
        
        # --- Форма добавления блюда ---
        with st.expander("➕ Добавить новое блюдо", expanded=True):
            col1, col2 = st.columns(2)
            with col1:
                dish_name = st.text_input("🍽️ Название блюда", key="new_dish_name")
                dish_category = st.selectbox(
                    "📁 Категория", 
                    categories if categories else ["Пиццы", "Закуски", "Салаты", "Напитки", "Десерты"],
                    key="new_dish_category"
                )
            with col2:
                dish_price = st.number_input("💰 Цена (₸)", min_value=0, value=500, step=50, key="new_dish_price")
                dish_cost = st.number_input("📦 Себестоимость (₸)", min_value=0, value=200, step=50, key="new_dish_cost")
            
            if st.button("✅ Сохранить блюдо", key="save_dish"):
                if dish_name:
                    if dish_name in menu_df["Блюдо"].values:
                        st.warning("⚠️ Блюдо с таким названием уже существует")
                    else:
                        new_row = pd.DataFrame([{
                            "Блюдо": dish_name,
                            "Категория": dish_category,
                            "Цена продажи (₸)": dish_price,
                            "Себестоимость (₸)": dish_cost
                        }])
                        updated_df = pd.concat([menu_df, new_row], ignore_index=True) if not menu_df.empty else new_row
                        # Убираем строки-заглушки, если они есть
                        updated_df = updated_df[~updated_df["Блюдо"].str.startswith("---", na=False)]
                        if save_to_google_sheet(updated_df, "МЕНЮ"):
                            st.success(f"✅ Блюдо '{dish_name}' добавлено!")
                            st.rerun()
                else:
                    st.warning("⚠️ Введите название блюда")
        
        # --- Таблица всех блюд ---
        if not menu_df.empty:
            # Убираем строки-заглушки для отображения
            display_df = menu_df[~menu_df["Блюдо"].str.startswith("---", na=False)].copy()
            
            if not display_df.empty:
                st.markdown("#### 📋 Все блюда")
                
                # Редактирование существующих блюд
                st.markdown("##### ✏️ Редактировать блюдо")
                
                # Выбор блюда для редактирования
                dish_list = display_df["Блюдо"].tolist()
                selected_dish = st.selectbox("Выберите блюдо для редактирования", dish_list, key="edit_dish_select")
                
                if selected_dish:
                    dish_row = display_df[display_df["Блюдо"] == selected_dish].iloc[0]
                    
                    col1, col2, col3, col4 = st.columns([2, 2, 2, 1])
                    with col1:
                        new_name = st.text_input("Название", value=dish_row["Блюдо"], key="edit_dish_name")
                    with col2:
                        new_category = st.selectbox(
                            "Категория", 
                            categories if categories else ["Пиццы", "Закуски", "Салаты", "Напитки", "Десерты"],
                            index=categories.index(dish_row["Категория"]) if dish_row["Категория"] in categories else 0,
                            key="edit_dish_category"
                        )
                    with col3:
                        new_price = st.number_input("Цена (₸)", value=int(dish_row["Цена продажи (₸)"]), min_value=0, step=50, key="edit_dish_price")
                    with col4:
                        new_cost = st.number_input("Себестоимость (₸)", value=int(dish_row["Себестоимость (₸)"]), min_value=0, step=50, key="edit_dish_cost")
                    
                    col_btn1, col_btn2, col_btn3 = st.columns([1, 1, 1])
                    with col_btn1:
                        if st.button("💾 Сохранить изменения", key="save_dish_edit"):
                            # Обновляем данные
                            idx = menu_df[menu_df["Блюдо"] == selected_dish].index[0]
                            menu_df.at[idx, "Блюдо"] = new_name
                            menu_df.at[idx, "Категория"] = new_category
                            menu_df.at[idx, "Цена продажи (₸)"] = new_price
                            menu_df.at[idx, "Себестоимость (₸)"] = new_cost
                            
                            if save_to_google_sheet(menu_df, "МЕНЮ"):
                                st.success(f"✅ Блюдо '{new_name}' обновлено!")
                                st.rerun()
                    
                    with col_btn2:
                        if st.button("🗑️ Удалить блюдо", key="delete_dish"):
                            updated_df = menu_df[menu_df["Блюдо"] != selected_dish]
                            if save_to_google_sheet(updated_df, "МЕНЮ"):
                                st.success(f"✅ Блюдо '{selected_dish}' удалено!")
                                st.rerun()
                
                # Таблица всех блюд
                st.markdown("---")
                st.markdown("#### 📊 Полное меню")
                
                # Добавляем колонку с маржинальностью
                display_df["Маржинальность (%)"] = display_df.apply(
                    lambda x: ((x["Цена продажи (₸)"] - x["Себестоимость (₸)"]) / x["Цена продажи (₸)"] * 100) 
                    if x["Цена продажи (₸)"] > 0 else 0, 
                    axis=1
                )
                display_df["Маржинальность (%)"] = display_df["Маржинальность (%)"].round(1)
                
                st.dataframe(
                    display_df.style.format({
                        "Цена продажи (₸)": "{:,.0f}",
                        "Себестоимость (₸)": "{:,.0f}",
                        "Маржинальность (%)": "{:.1f}%"
                    }),
                    use_container_width=True
                )
                
                # Показать статистику по категориям
                st.markdown("---")
                st.markdown("#### 📊 Статистика по категориям")
                category_stats = display_df.groupby("Категория").agg({
                    "Блюдо": "count",
                    "Цена продажи (₸)": "mean",
                    "Маржинальность (%)": "mean"
                }).round(1)
                category_stats.columns = ["Кол-во блюд", "Средняя цена (₸)", "Средняя маржинальность (%)"]
                st.dataframe(category_stats, use_container_width=True)
                
            else:
                st.info("Нет блюд для отображения")
        else:
            st.info("Меню пусто. Добавьте первую категорию и блюдо.")
            
# =====================================================================
# 4. ЗАКУПКИ И СКЛАД
# =====================================================================
elif selected_section == "📦 Закупки и склад":
    st.markdown("## 📦 Закупки и склад")
    
    tab1, tab2 = st.tabs(["📦 Склад", "🛒 Закупки"])
    
    # =============================================================
    # ВКЛАДКА "СКЛАД"
    # =============================================================
    with tab1:
        st.markdown("### 📦 Управление складом")
        
        stock_df = load_sheet_data("СКЛАД")
        
        with st.expander("➕ Добавить ингредиент на склад", expanded=True):
            col1, col2 = st.columns(2)
            with col1:
                ing_name = st.text_input("📦 Название ингредиента", key="stock_name")
                ing_unit = st.selectbox("📏 Единица измерения", ["кг", "г", "л", "шт", "уп."], key="stock_unit")
            with col2:
                ing_quantity = st.number_input("📊 Количество", min_value=0.0, value=10.0, step=0.5, key="stock_qty")
                ing_min_stock = st.number_input("⚠️ Мин. запас", min_value=0.0, value=2.0, step=0.5, key="stock_min")
            
            if st.button("✅ Добавить на склад", key="add_stock"):
                if ing_name:
                    new_row = pd.DataFrame([{
                        "Ингредиент": ing_name,
                        "Единица": ing_unit,
                        "Количество": ing_quantity,
                        "Мин. запас": ing_min_stock
                    }])
                    updated_df = pd.concat([stock_df, new_row], ignore_index=True) if not stock_df.empty else new_row
                    if save_to_google_sheet(updated_df, "СКЛАД"):
                        st.success(f"✅ Ингредиент '{ing_name}' добавлен на склад!")
                        st.rerun()
                else:
                    st.warning("⚠️ Введите название ингредиента")
        
        if not stock_df.empty:
            st.markdown("### 📋 Текущие остатки")
            stock_df["Статус"] = stock_df.apply(
                lambda x: "⚠️ Критический запас" if x["Количество"] < x["Мин. запас"] else "✅ В наличии",
                axis=1
            )
            
            st.dataframe(
                stock_df.style.applymap(
                    lambda val: "color: #e74c3c; font-weight: bold" if val == "⚠️ Критический запас" else "color: #27ae60",
                    subset=["Статус"]
                ),
                use_container_width=True
            )
            
            col1, col2, col3 = st.columns(3)
            col1.metric("📦 Всего ингредиентов", len(stock_df))
            col2.metric("⚠️ Критический запас", len(stock_df[stock_df["Количество"] < stock_df["Мин. запас"]]))
            col3.metric("📊 Всего единиц", f"{stock_df['Количество'].sum():.0f}")
        else:
            st.info("📭 Склад пуст. Добавьте ингредиенты через форму выше.")
    
    # =============================================================
    # ВКЛАДКА "ЗАКУПКИ"
    # =============================================================
    with tab2:
        st.markdown("### 🛒 Управление закупками")
        
        purchases_df = load_sheet_data("ЗАКУПКИ")
        
        with st.expander("➕ Добавить закупку", expanded=True):
            col1, col2 = st.columns(2)
            with col1:
                purchase_supplier = st.text_input("🏢 Поставщик", key="purchase_supplier")
                purchase_product = st.text_input("📦 Товар", key="purchase_product")
            with col2:
                purchase_qty = st.number_input("📊 Количество", min_value=0.0, value=1.0, step=0.5, key="purchase_qty")
                purchase_price = st.number_input("💰 Цена за ед. (₸)", min_value=0, value=100, step=50, key="purchase_price")
            
            purchase_note = st.text_area(
                "📝 Примечание",
                placeholder="Номер накладной, срок годности и т.д.",
                key="purchase_note"
            )
            
            if st.button("✅ Добавить закупку", key="add_purchase"):
                if purchase_product and purchase_supplier:
                    new_row = pd.DataFrame([{
                        "Дата": datetime.now().strftime("%d.%m.%Y"),
                        "Поставщик": purchase_supplier,
                        "Товар": purchase_product,
                        "Количество": purchase_qty,
                        "Цена за ед. (₸)": purchase_price,
                        "Сумма (₸)": purchase_qty * purchase_price,
                        "Примечание": purchase_note
                    }])
                    updated_df = pd.concat([purchases_df, new_row], ignore_index=True) if not purchases_df.empty else new_row
                    if save_to_google_sheet(updated_df, "ЗАКУПКИ"):
                        st.success(f"✅ Закупка '{purchase_product}' добавлена!")
                        st.rerun()
                else:
                    st.warning("⚠️ Заполните поставщика и товар")
        
        if not purchases_df.empty:
            st.markdown("### 📋 История закупок")
            st.dataframe(purchases_df, use_container_width=True)
            
            col1, col2, col3 = st.columns(3)
            col1.metric("🛒 Всего закупок", len(purchases_df))
            col2.metric("💰 Общая сумма", f"{purchases_df['Сумма (₸)'].sum():,.0f} ₸")
            col3.metric("📈 Средняя цена", f"{purchases_df['Цена за ед. (₸)'].mean():,.0f} ₸")
        else:
            st.info("📭 Нет закупок. Добавьте первую закупку через форму выше.")           
            
# =====================================================================
# 5. ПЕРСОНАЛ (с PIN-кодами)
# =====================================================================
elif selected_section == "👥 Персонал":
    st.markdown("## 👥 Управление персоналом")
    
    staff_df = load_sheet_data("ПЕРСОНАЛ")
    
    with st.expander("➕ Добавить сотрудника", expanded=True):
        col1, col2 = st.columns(2)
        with col1:
            name = st.text_input("Имя")
            role = st.selectbox("Должность", ["Администратор", "Официант", "Повар", "Бухгалтер"])
        with col2:
            phone = st.text_input("Телефон")
            pin = st.text_input("PIN-код (4 цифры)", type="password", max_chars=4)
        
        if st.button("✅ Добавить сотрудника"):
            if name and pin and len(pin) == 4:
                new_row = pd.DataFrame([{
                    "ID": len(staff_df) + 1 if not staff_df.empty else 1,
                    "Имя": name,
                    "Должность": role,
                    "Телефон": phone,
                    "PIN-код": pin,
                    "Активен": "Да"
                }])
                updated_df = pd.concat([staff_df, new_row], ignore_index=True) if not staff_df.empty else new_row
                if save_to_google_sheet(updated_df, "ПЕРСОНАЛ"):
                    st.success(f"✅ Сотрудник '{name}' добавлен!")
                    st.rerun()
            else:
                st.warning("Заполните имя и PIN-код из 4 цифр")
    
    if not staff_df.empty:
        st.markdown("### 📋 Список сотрудников")
        st.dataframe(staff_df, use_container_width=True)

# =====================================================================
# 6. ОТЧЕТЫ
# =====================================================================
elif selected_section == "📈 Отчеты":
    st.markdown("## 📈 Отчеты")
    
    orders_df = load_sheet_data("ЗАКАЗЫ")
    
    if orders_df.empty:
        st.info("Нет заказов для отчетов")
    else:
        total_orders = len(orders_df)
        total_revenue = orders_df["Сумма (₸)"].sum() if "Сумма (₸)" in orders_df.columns else 0
        avg_check = total_revenue / total_orders if total_orders > 0 else 0
        
        col1, col2, col3 = st.columns(3)
        col1.metric("📋 Всего заказов", total_orders)
        col2.metric("💰 Общая выручка", f"{total_revenue:,.0f} ₸")
        col3.metric("📈 Средний чек", f"{avg_check:,.0f} ₸")
        
        st.markdown("### 📋 Детали заказов")
        st.dataframe(orders_df, use_container_width=True)

st.markdown("---")
st.caption("☕ Система управления кафе | Единое приложение | © 2026")