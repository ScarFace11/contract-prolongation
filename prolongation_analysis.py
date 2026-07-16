# =============================================================================
# Анализ пролонгаций аккаунт-менеджеров за 2023 год
# =============================================================================

# Работа с данными
import pandas as pd
import numpy as np

# Визуализация
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import seaborn as sns

# Excel
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# =============================================================================
# 1. ЗАГРУЗКА ДАННЫХ
# =============================================================================

prolongations = pd.read_csv(
    "attached_assets/prolongations_1784238197480.csv",
    sep=None,
    engine="python",
    encoding="utf-8-sig"
)

financial = pd.read_csv(
    "attached_assets/financial_data_1784238197480.csv",
    sep=None,
    engine="python",
    encoding="utf-8-sig"
)

# =============================================================================
# 2. ВСПОМОГАТЕЛЬНЫЕ СЛОВАРИ И ФУНКЦИИ
# =============================================================================

months_column = {
    1: "Январь", 2: "Февраль", 3: "Март",    4: "Апрель",
    5: "Май",    6: "Июнь",   7: "Июль",    8: "Август",
    9: "Сентябрь", 10: "Октябрь", 11: "Ноябрь", 12: "Декабрь"
}

months_parse = {v.lower(): k for k, v in months_column.items()}

def parse_month(value):
    """Парсит строку вида 'январь 2023' в pd.Timestamp."""
    month_str, year_str = value.lower().split()
    return pd.Timestamp(year=int(year_str), month=months_parse[month_str], day=1)

def month_label(date):
    """Возвращает название колонки финансовых данных, например 'Январь 2023'."""
    return f"{months_column[date.month]} {date.year}"

# =============================================================================
# 3. ПОДГОТОВКА ДАННЫХ
# =============================================================================

# --- prolongations: парсим month ---
prolongations["month"] = prolongations["month"].apply(parse_month)

# --- Приводим id к строке в обоих датафреймах ---
financial["id"] = financial["id"].astype(str)
prolongations["id"] = prolongations["id"].astype(str)

# --- Определяем колонки с суммами отгрузок ---
service_columns = ["id", "Причина дубля", "Account"]
month_columns = [col for col in financial.columns if col not in service_columns]

def money_to_float(value):
    """
    Преобразует сумму из строкового формата ('36 220,00', 'в ноль', 'стоп')
    в float. Нечисловые значения → 0.
    """
    if pd.isna(value):
        return 0.0
    if isinstance(value, str):
        cleaned = (
            value
            .replace("\xa0", "")
            .replace(" ", "")
            .replace(",", ".")
        )
        try:
            return float(cleaned)
        except ValueError:
            return 0.0
    return float(value)

# Применяем конвертацию ко всем месячным колонкам
for col in month_columns:
    financial[col] = financial[col].apply(money_to_float)

# --- Агрегируем дубли по id: суммируем отгрузки ---
# (одному проекту может соответствовать несколько строк — «первая часть», «вторая часть» и т.д.)
financial_grouped = (
    financial
    .groupby("id")[month_columns]
    .sum()
    .reset_index()
)

# =============================================================================
# 4. ОБЪЕДИНЕНИЕ: ПРОЕКТЫ + ФИНАНСЫ
# =============================================================================

# Источник AM — prolongations (первичен согласно ТЗ)
projects = (
    prolongations[["id", "month", "AM"]]
    .merge(financial_grouped, on="id", how="left")
)
projects["AM"] = projects["AM"].fillna("Не указан")

# Заполняем пропуски в суммах нулями (проект есть, но отгрузки не было)
projects[month_columns] = projects[month_columns].fillna(0)

# =============================================================================
# 5. РАСЧЁТ KPI ПРОЛОНГАЦИЙ
# =============================================================================

def calculate_manager_kpi(data, report_month):
    """
    Рассчитывает коэффициенты пролонгации для каждого менеджера
    за указанный отчётный месяц.

    Логика (пример: report_month = май 2023):

    Коэффициент 1 месяц:
      Числитель   — сумма отгрузки в МАЕ  проектов, завершившихся в АПРЕЛЕ
      Знаменатель — сумма отгрузки в АПРЕЛЕ тех же проектов
      (пролонгация в первый месяц после завершения)

    Коэффициент 2 месяц:
      Рассматриваем проекты, завершившиеся в МАРТЕ и НЕ продлившиеся в АПРЕЛЕ
      (т.е. сумма апреля == 0)
      Числитель   — сумма отгрузки в МАЕ  таких проектов
      Знаменатель — сумма отгрузки в МАРТЕ таких проектов
    """
    result = []

    current   = pd.Timestamp(report_month)           # отчётный месяц
    previous  = current - pd.DateOffset(months=1)    # месяц назад
    two_month = current - pd.DateOffset(months=2)    # два месяца назад

    current_col   = month_label(current)
    previous_col  = month_label(previous)
    two_month_col = month_label(two_month)

    # Пропускаем, если нужные колонки отсутствуют в данных
    for col in [current_col, previous_col, two_month_col]:
        if col not in month_columns:
            return pd.DataFrame()

    for manager, df in data.groupby("AM"):

        # ── Коэффициент 1 месяц ───────────────────────────────────────────
        # Проекты, чей последний месяц реализации = previous (апрель)
        first_month_projects = df[df["month"] == previous]

        # Знаменатель: сумма отгрузки за апрель по этим проектам
        first_denominator = first_month_projects[previous_col].sum()

        # Числитель: сумма отгрузки за май по тем же проектам
        first_numerator = first_month_projects[current_col].sum()

        coef_first = (
            first_numerator / first_denominator
            if first_denominator > 0 else 0.0
        )

        # ── Коэффициент 2 месяц ───────────────────────────────────────────
        # Проекты, чей последний месяц реализации = two_month (март)
        second_month_projects = df[df["month"] == two_month]

        # Отбираем только те, у которых НЕТ отгрузки в previous (апрель).
        # Используем <= 0 вместо == 0, чтобы корректно обработать любые
        # остаточные нули/NaN после fillna(0).
        not_extended_first = second_month_projects[previous_col] <= 0

        denominator_projects = second_month_projects[not_extended_first]

        # Знаменатель: сумма отгрузки за март по непродлённым в апреле
        denominator_second = denominator_projects[two_month_col].sum()

        # Числитель: сумма отгрузки за май по тем же проектам
        numerator_second = denominator_projects[current_col].sum()

        coef_second = (
            numerator_second / denominator_second
            if denominator_second > 0 else 0.0
        )

        result.append({
            "Менеджер":               manager,
            "Месяц":                  current.strftime("%Y-%m"),
            "1 месяц база":           first_denominator,
            "1 месяц продлено":       first_numerator,
            "Коэффициент 1 месяц":    coef_first,
            "2 месяц база":           denominator_second,
            "2 месяц продлено":       numerator_second,
            "Коэффициент 2 месяц":    coef_second,
        })

    return pd.DataFrame(result)

# --- Считаем KPI за каждый месяц 2023 года ---
all_results = []
for month in pd.date_range("2023-01-01", "2023-12-01", freq="MS"):
    temp = calculate_manager_kpi(projects, month)
    if not temp.empty:
        all_results.append(temp)

kpi_monthly = pd.concat(all_results, ignore_index=True)

# =============================================================================
# 6. АГРЕГАЦИЯ: ОТДЕЛ (по месяцам) и ГОД (по менеджерам)
# =============================================================================

# --- Показатели отдела по месяцам ---
department = (
    kpi_monthly
    .groupby("Месяц")
    .agg({
        "1 месяц база":     "sum",
        "1 месяц продлено": "sum",
        "2 месяц база":     "sum",
        "2 месяц продлено": "sum",
    })
    .reset_index()
)

# Пересчитываем коэффициенты через суммы (а не среднее коэффициентов)
department["Коэффициент 1 месяц"] = np.where(
    department["1 месяц база"] > 0,
    department["1 месяц продлено"] / department["1 месяц база"],
    0.0
)
department["Коэффициент 2 месяц"] = np.where(
    department["2 месяц база"] > 0,
    department["2 месяц продлено"] / department["2 месяц база"],
    0.0
)

# --- Годовые показатели по менеджерам ---
year_manager = (
    kpi_monthly
    .groupby("Менеджер")
    .agg({
        "1 месяц база":     "sum",
        "1 месяц продлено": "sum",
        "2 месяц база":     "sum",
        "2 месяц продлено": "sum",
    })
    .reset_index()
)

year_manager["Коэффициент 1 месяц"] = np.where(
    year_manager["1 месяц база"] > 0,
    year_manager["1 месяц продлено"] / year_manager["1 месяц база"],
    np.nan  # менеджер не имел завершившихся проектов → нет данных
)
year_manager["Коэффициент 2 месяц"] = np.where(
    year_manager["2 месяц база"] > 0,
    year_manager["2 месяц продлено"] / year_manager["2 месяц база"],
    np.nan
)

# =============================================================================
# 7. ВИЗУАЛИЗАЦИЯ (ДО конвертации коэффициентов в строки!)
# =============================================================================

fig, axes = plt.subplots(1, 2, figsize=(16, 5))
fig.suptitle("Динамика коэффициентов пролонгации отдела — 2023", fontsize=14, fontweight="bold")

# Подготовим данные с числовыми значениями для графика
dept_plot = department.copy()
dept_plot["Месяц_label"] = dept_plot["Месяц"]

ax1 = axes[0]
ax1.plot(dept_plot["Месяц_label"], dept_plot["Коэффициент 1 месяц"],
         marker="o", color="#2196F3", label="Коэф. 1 месяц")
ax1.set_title("Коэффициент пролонгации: 1-й месяц")
ax1.set_xlabel("Месяц")
ax1.set_ylabel("Коэффициент")
ax1.yaxis.set_major_formatter(mticker.PercentFormatter(xmax=1))
ax1.tick_params(axis="x", rotation=45)
ax1.grid(True, alpha=0.3)
ax1.legend()

ax2 = axes[1]
ax2.plot(dept_plot["Месяц_label"], dept_plot["Коэффициент 2 месяц"],
         marker="o", color="#FF5722", label="Коэф. 2 месяц")
ax2.set_title("Коэффициент пролонгации: 2-й месяц")
ax2.set_xlabel("Месяц")
ax2.set_ylabel("Коэффициент")
ax2.yaxis.set_major_formatter(mticker.PercentFormatter(xmax=1))
ax2.tick_params(axis="x", rotation=45)
ax2.grid(True, alpha=0.3)
ax2.legend()

plt.tight_layout()
plt.savefig("prolongation_chart.png", dpi=150, bbox_inches="tight")
plt.close()
print("График сохранён: prolongation_chart.png")

# =============================================================================
# 8. ФОРМАТИРОВАНИЕ И СОХРАНЕНИЕ В EXCEL
# =============================================================================

def fmt_pct(x):
    """Форматирует float как процент; NaN → 'Нет данных'."""
    if pd.isna(x):
        return "Нет данных"
    return f"{x:.1%}"

def fmt_money(x):
    """Форматирует число как сумму с пробелами."""
    if pd.isna(x):
        return ""
    return f"{x:,.0f}".replace(",", " ")

file_name = "prolongation_report.xlsx"

# Копии для Excel с форматированием (не портим числовые данные для графика)
dept_excel   = department.copy()
monthly_excel = kpi_monthly.copy()
year_excel   = year_manager.copy()

# Форматируем коэффициенты в строки
for df in [dept_excel, monthly_excel, year_excel]:
    df["Коэффициент 1 месяц"] = df["Коэффициент 1 месяц"].apply(fmt_pct)
    df["Коэффициент 2 месяц"] = df["Коэффициент 2 месяц"].apply(fmt_pct)

# Форматируем денежные суммы
money_cols = ["1 месяц база", "1 месяц продлено", "2 месяц база", "2 месяц продлено"]
for df in [dept_excel, monthly_excel, year_excel]:
    for col in money_cols:
        df[col] = df[col].apply(fmt_money)

# Записываем в Excel
with pd.ExcelWriter(file_name, engine="openpyxl") as writer:
    dept_excel.to_excel(writer, sheet_name="Department", index=False)
    monthly_excel.to_excel(writer, sheet_name="Managers", index=False)
    year_excel.to_excel(writer, sheet_name="Year KPI", index=False)

# =============================================================================
# 9. СТИЛИЗАЦИЯ EXCEL (openpyxl)
# =============================================================================

HEADER_COLOR  = "2D6A9F"   # тёмно-синий
HEADER_FONT_COLOR = "FFFFFF"
ALT_ROW_COLOR = "EBF3FB"   # светло-голубой
BORDER_COLOR  = "B0C4DE"

thin_border = Border(
    left=Side(style="thin", color=BORDER_COLOR),
    right=Side(style="thin", color=BORDER_COLOR),
    top=Side(style="thin", color=BORDER_COLOR),
    bottom=Side(style="thin", color=BORDER_COLOR),
)

def style_sheet(ws):
    """Применяет единый стиль к листу."""
    # Заголовок
    for cell in ws[1]:
        cell.font = Font(bold=True, color=HEADER_FONT_COLOR, size=11)
        cell.fill = PatternFill("solid", fgColor=HEADER_COLOR)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    # Данные
    for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
        bg = ALT_ROW_COLOR if row_idx % 2 == 0 else "FFFFFF"
        for cell in row:
            cell.fill = PatternFill("solid", fgColor=bg)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border

    # Автоширина колонок
    for col in ws.columns:
        max_len = max(
            (len(str(cell.value)) if cell.value is not None else 0)
            for cell in col
        )
        ws.column_dimensions[col[0].column_letter].width = max(max_len + 4, 12)

    # Высота строки заголовка
    ws.row_dimensions[1].height = 30

wb = load_workbook(file_name)
for ws in wb.worksheets:
    style_sheet(ws)
wb.save(file_name)

print(f"Отчёт сохранён: {file_name}")
print("\n=== Краткие итоги ===")
print("\nКоэффициенты пролонгации отдела по месяцам:")
print(dept_excel[["Месяц", "1 месяц база", "1 месяц продлено", "Коэффициент 1 месяц",
                   "2 месяц база", "2 месяц продлено", "Коэффициент 2 месяц"]].to_string(index=False))
print("\nГодовые коэффициенты по менеджерам:")
print(year_excel[["Менеджер", "Коэффициент 1 месяц", "Коэффициент 2 месяц"]].to_string(index=False))
