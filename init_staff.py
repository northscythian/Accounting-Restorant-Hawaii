import pandas as pd
import os

file_path = "Кафе_Учет.xlsx"

staff_data = pd.DataFrame({
    "ID": [1, 2, 3],
    "Имя": ["Анна", "Иван", "Сергей"],
    "Должность": ["Официант", "Официант", "Администратор"],
    "Телефон": ["+7XXX", "+7XXX", "+7XXX"],
    "PIN-код": ["1234", "5678", "0000"],
    "Активен": ["Да", "Да", "Да"],
    "Дата добавления": ["10.06.2026", "10.06.2026", "10.06.2026"]
})

if os.path.exists(file_path):
    with pd.ExcelWriter(file_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        staff_data.to_excel(writer, sheet_name="ПЕРСОНАЛ", index=False)
    print("✅ Лист 'ПЕРСОНАЛ' создан!")
else:
    print("Файл не найден")