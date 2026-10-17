"""Оконное приложение для создания демонстрационного отчёта в Google Sheets."""

import os
import random
import threading
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, ttk

from dotenv import load_dotenv

from google_sheets import GoogleSheets


BASE_DIR = Path(__file__).resolve().parent
DATE_FORMAT = "%d.%m.%Y"
MANAGERS = ("Анна", "Борис", "Виктория", "Глеб", "Дарья")


@dataclass(frozen=True)
class ReportRow:
    """Одна строка демонстрационного отчёта."""

    report_date: date
    manager: str
    leads: int
    deals: int
    revenue: int

    @property
    def average_check(self) -> float:
        return self.revenue / self.deals if self.deals else 0

    @property
    def conversion(self) -> float:
        return self.deals / self.leads if self.leads else 0


def load_settings() -> tuple[Path, str]:
    """Загружает локальные настройки подключения из файла .env."""
    load_dotenv(BASE_DIR / ".env")
    credentials_value = os.getenv("GOOGLE_CREDENTIALS_FILE", "").strip()
    spreadsheet_id = os.getenv("SPREADSHEET_ID", "").strip()

    if not credentials_value:
        raise ValueError("В .env не указан GOOGLE_CREDENTIALS_FILE.")
    if not spreadsheet_id:
        raise ValueError(
            "В .env не указан SPREADSHEET_ID. Добавьте ID созданной Google Таблицы."
        )

    credentials_path = Path(credentials_value)
    if not credentials_path.is_absolute():
        credentials_path = BASE_DIR / credentials_path
    return credentials_path, spreadsheet_id


def generate_report_rows(
    start_date: date,
    end_date: date,
    row_count: int = 10,
) -> list[ReportRow]:
    """Создаёт демонстрационные показатели внутри выбранного периода."""
    if end_date < start_date:
        raise ValueError("Дата окончания не может быть раньше даты начала.")

    day_span = (end_date - start_date).days
    rows: list[ReportRow] = []
    for index in range(row_count):
        leads = random.randint(15, 80)
        deals = random.randint(1, max(1, min(leads, 25)))
        revenue = deals * random.randint(8_000, 35_000)
        rows.append(
            ReportRow(
                report_date=start_date + timedelta(days=random.randint(0, day_span)),
                manager=MANAGERS[index % len(MANAGERS)],
                leads=leads,
                deals=deals,
                revenue=revenue,
            )
        )
    return sorted(rows, key=lambda row: (row.report_date, row.manager))


def grid_range(
    sheet_id: int,
    start_row: int,
    end_row: int,
    start_column: int,
    end_column: int,
) -> dict[str, int]:
    """Формирует диапазон GridRange с индексами, начинающимися с нуля."""
    return {
        "sheetId": sheet_id,
        "startRowIndex": start_row,
        "endRowIndex": end_row,
        "startColumnIndex": start_column,
        "endColumnIndex": end_column,
    }


def build_format_requests(
    sheet_id: int,
    first_data_row: int,
    last_data_row: int,
    total_row: int,
) -> list[dict]:
    """Создаёт операции оформления готового отчёта."""
    dark_blue = {"red": 0.12, "green": 0.28, "blue": 0.49}
    medium_blue = {"red": 0.22, "green": 0.46, "blue": 0.70}
    pale_blue = {"red": 0.91, "green": 0.95, "blue": 0.99}
    pale_gray = {"red": 0.95, "green": 0.95, "blue": 0.95}
    border_color = {"red": 0.82, "green": 0.82, "blue": 0.82}
    white = {"red": 1, "green": 1, "blue": 1}
    black = {"red": 0, "green": 0, "blue": 0}

    requests: list[dict] = [
        {
            "mergeCells": {
                "range": grid_range(sheet_id, 0, 1, 0, 7),
                "mergeType": "MERGE_ALL",
            }
        },
        *[
            {
                "mergeCells": {
                    "range": grid_range(sheet_id, row_start, row_end, col_start, col_end),
                    "mergeType": "MERGE_ALL",
                }
            }
            for row_start, row_end, col_start, col_end in (
                (2, 3, 0, 2),
                (2, 3, 2, 4),
                (2, 3, 4, 6),
                (4, 5, 0, 2),
                (4, 5, 2, 4),
                (4, 5, 4, 6),
                (5, 6, 0, 2),
                (5, 6, 2, 4),
                (5, 6, 4, 6),
                (total_row - 1, total_row, 0, 2),
            )
        ],
        {
            "repeatCell": {
                "range": grid_range(sheet_id, 0, 1, 0, 7),
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": dark_blue,
                        "horizontalAlignment": "CENTER",
                        "verticalAlignment": "MIDDLE",
                        "textFormat": {
                            "foregroundColor": white,
                            "fontSize": 18,
                            "bold": True,
                        },
                    }
                },
                "fields": "userEnteredFormat",
            }
        },
        {
            "repeatCell": {
                "range": grid_range(sheet_id, 2, 3, 0, 7),
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": pale_gray,
                        "verticalAlignment": "MIDDLE",
                        "textFormat": {"foregroundColor": black},
                    }
                },
                "fields": "userEnteredFormat",
            }
        },
        {
            "repeatCell": {
                "range": grid_range(sheet_id, 2, 3, 0, 2),
                "cell": {"userEnteredFormat": {"textFormat": {"bold": True}}},
                "fields": "userEnteredFormat.textFormat.bold",
            }
        },
        {
            "repeatCell": {
                "range": grid_range(sheet_id, 2, 3, 4, 6),
                "cell": {"userEnteredFormat": {"textFormat": {"bold": True}}},
                "fields": "userEnteredFormat.textFormat.bold",
            }
        },
        {
            "repeatCell": {
                "range": grid_range(sheet_id, 4, 5, 0, 7),
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": medium_blue,
                        "horizontalAlignment": "CENTER",
                        "textFormat": {"foregroundColor": white, "bold": True},
                    }
                },
                "fields": "userEnteredFormat",
            }
        },
        {
            "repeatCell": {
                "range": grid_range(sheet_id, 5, 6, 0, 7),
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": pale_blue,
                        "horizontalAlignment": "CENTER",
                        "textFormat": {"fontSize": 12, "bold": True},
                    }
                },
                "fields": "userEnteredFormat",
            }
        },
        {
            "repeatCell": {
                "range": grid_range(sheet_id, 7, 8, 0, 7),
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": dark_blue,
                        "horizontalAlignment": "CENTER",
                        "verticalAlignment": "MIDDLE",
                        "wrapStrategy": "WRAP",
                        "textFormat": {"foregroundColor": white, "bold": True},
                    }
                },
                "fields": "userEnteredFormat",
            }
        },
        {
            "repeatCell": {
                "range": grid_range(sheet_id, total_row - 1, total_row, 0, 7),
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": pale_blue,
                        "textFormat": {"bold": True},
                    }
                },
                "fields": "userEnteredFormat",
            }
        },
        {
            "repeatCell": {
                "range": grid_range(sheet_id, 0, total_row, 0, 7),
                "cell": {
                    "userEnteredFormat": {
                        "verticalAlignment": "MIDDLE",
                        "borders": {
                            side: {"style": "SOLID", "color": border_color}
                            for side in ("top", "bottom", "left", "right")
                        },
                    }
                },
                "fields": "userEnteredFormat.verticalAlignment,userEnteredFormat.borders",
            }
        },
        {
            "repeatCell": {
                "range": grid_range(sheet_id, first_data_row - 1, total_row, 4, 6),
                "cell": {
                    "userEnteredFormat": {
                        "numberFormat": {
                            "type": "NUMBER",
                            "pattern": '#,##0.00 "₽"',
                        }
                    }
                },
                "fields": "userEnteredFormat.numberFormat",
            }
        },
        {
            "repeatCell": {
                "range": grid_range(sheet_id, first_data_row - 1, total_row, 6, 7),
                "cell": {
                    "userEnteredFormat": {
                        "numberFormat": {"type": "PERCENT", "pattern": "0.0%"}
                    }
                },
                "fields": "userEnteredFormat.numberFormat",
            }
        },
        {
            "repeatCell": {
                "range": grid_range(sheet_id, 5, 6, 4, 6),
                "cell": {
                    "userEnteredFormat": {
                        "numberFormat": {
                            "type": "NUMBER",
                            "pattern": '#,##0.00 "₽"',
                        }
                    }
                },
                "fields": "userEnteredFormat.numberFormat",
            }
        },
        {
            "repeatCell": {
                "range": grid_range(sheet_id, 5, 6, 6, 7),
                "cell": {
                    "userEnteredFormat": {
                        "numberFormat": {"type": "PERCENT", "pattern": "0.0%"}
                    }
                },
                "fields": "userEnteredFormat.numberFormat",
            }
        },
        {
            "updateSheetProperties": {
                "properties": {
                    "sheetId": sheet_id,
                    "gridProperties": {"frozenRowCount": 8},
                },
                "fields": "gridProperties.frozenRowCount",
            }
        },
    ]

    for row_index in range(first_data_row - 1, last_data_row):
        if (row_index - first_data_row + 1) % 2:
            requests.append(
                {
                    "repeatCell": {
                        "range": grid_range(sheet_id, row_index, row_index + 1, 0, 7),
                        "cell": {"userEnteredFormat": {"backgroundColor": pale_blue}},
                        "fields": "userEnteredFormat.backgroundColor",
                    }
                }
            )

    widths = (105, 130, 85, 85, 130, 130, 105)
    for column_index, width in enumerate(widths):
        requests.append(
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": "COLUMNS",
                        "startIndex": column_index,
                        "endIndex": column_index + 1,
                    },
                    "properties": {"pixelSize": width},
                    "fields": "pixelSize",
                }
            }
        )

    for row_index, height in ((0, 44), (2, 32), (4, 28), (5, 36), (7, 38)):
        requests.append(
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": "ROWS",
                        "startIndex": row_index,
                        "endIndex": row_index + 1,
                    },
                    "properties": {"pixelSize": height},
                    "fields": "pixelSize",
                }
            }
        )
    return requests


def create_report(
    sheets: GoogleSheets,
    start_date: date,
    end_date: date,
    department: str,
) -> str:
    """Создаёт новый лист, записывает данные и применяет оформление."""
    rows = generate_report_rows(start_date, end_date)
    sheet_title = datetime.now().strftime("Отчёт_%Y%m%d_%H%M%S")
    sheet_id = sheets.create_sheet(sheet_title)
    quoted_title = sheets.quote_sheet_title(sheet_title)

    first_data_row = 9
    last_data_row = first_data_row + len(rows) - 1
    total_row = last_data_row + 1

    values: list[list[object]] = [
        ["АВТОМАТИЧЕСКИЙ ОТЧЁТ"],
        [],
        [
            "Период",
            "",
            f"{start_date.strftime(DATE_FORMAT)} — {end_date.strftime(DATE_FORMAT)}",
            "",
            "Отдел",
            "",
            department,
        ],
        [],
        ["Лиды", "", "Сделки", "", "Выручка", "", "Конверсия"],
        [],
        [],
        ["Дата", "Менеджер", "Лиды", "Сделки", "Выручка", "Средний чек", "Конверсия"],
    ]
    values.extend(
        [
            row.report_date.strftime(DATE_FORMAT),
            row.manager,
            row.leads,
            row.deals,
            row.revenue,
            round(row.average_check, 2),
            row.conversion,
        ]
        for row in rows
    )
    values.append(["Итого", "", "", "", "", "", ""])

    sheets.write_range(f"{quoted_title}!A1:G{total_row}", values, "RAW")
    sheets.write_ranges(
        [
            {"range": f"{quoted_title}!A6", "values": [[f"=C{total_row}"]]},
            {"range": f"{quoted_title}!C6", "values": [[f"=D{total_row}"]]},
            {"range": f"{quoted_title}!E6", "values": [[f"=E{total_row}"]]},
            {"range": f"{quoted_title}!G6", "values": [[f"=G{total_row}"]]},
            {
                "range": f"{quoted_title}!C{total_row}:G{total_row}",
                "values": [[
                    f"=SUM(C{first_data_row}:C{last_data_row})",
                    f"=SUM(D{first_data_row}:D{last_data_row})",
                    f"=SUM(E{first_data_row}:E{last_data_row})",
                    f"=IFERROR(E{total_row}/D{total_row};0)",
                    f"=IFERROR(D{total_row}/C{total_row};0)",
                ]],
            },
        ],
        "USER_ENTERED",
    )
    sheets.batch_update(
        build_format_requests(
            sheet_id,
            first_data_row,
            last_data_row,
            total_row,
        )
    )
    return sheet_title


class ReportGeneratorApp:
    """Интерфейс для ввода параметров и запуска генерации отчёта."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Генератор отчётов")
        self.root.geometry("560x360")
        self.root.minsize(520, 330)

        today = date.today()
        month_start = today.replace(day=1)

        self.start_date = tk.StringVar(value=month_start.strftime(DATE_FORMAT))
        self.end_date = tk.StringVar(value=today.strftime(DATE_FORMAT))
        self.department = tk.StringVar(value="Отдел продаж")
        self.status = tk.StringVar(value="Заполните параметры и создайте отчёт.")

        self._build_ui()

    def _build_ui(self) -> None:
        container = ttk.Frame(self.root, padding=28)
        container.pack(fill="both", expand=True)

        ttk.Label(
            container,
            text="Автоматический отчёт в Google Sheets",
            font=("Segoe UI", 16, "bold"),
        ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 22))

        fields = (
            ("Дата начала", self.start_date),
            ("Дата окончания", self.end_date),
            ("Отдел или проект", self.department),
        )
        for row_index, (label, variable) in enumerate(fields, start=1):
            ttk.Label(container, text=label).grid(
                row=row_index,
                column=0,
                sticky="w",
                padx=(0, 18),
                pady=8,
            )
            ttk.Entry(container, textvariable=variable, width=34).grid(
                row=row_index,
                column=1,
                sticky="ew",
                pady=8,
            )

        ttk.Label(
            container,
            text="Формат даты: ДД.ММ.ГГГГ",
            foreground="#555555",
        ).grid(row=4, column=1, sticky="w", pady=(0, 16))

        self.generate_button = ttk.Button(
            container,
            text="Сгенерировать отчёт",
            command=self._start_generation,
        )
        self.generate_button.grid(row=5, column=0, columnspan=2, pady=(4, 18))

        ttk.Separator(container).grid(row=6, column=0, columnspan=2, sticky="ew")
        ttk.Label(
            container,
            textvariable=self.status,
            wraplength=490,
            justify="left",
        ).grid(row=7, column=0, columnspan=2, sticky="w", pady=(16, 0))

        container.columnconfigure(1, weight=1)

    def _read_form(self) -> tuple[date, date, str]:
        try:
            start_date = datetime.strptime(
                self.start_date.get().strip(), DATE_FORMAT
            ).date()
            end_date = datetime.strptime(
                self.end_date.get().strip(), DATE_FORMAT
            ).date()
        except ValueError as error:
            raise ValueError("Введите даты в формате ДД.ММ.ГГГГ.") from error

        if end_date < start_date:
            raise ValueError("Дата окончания не может быть раньше даты начала.")
        department = self.department.get().strip()
        if not department:
            raise ValueError("Укажите название отдела или проекта.")
        return start_date, end_date, department

    def _start_generation(self) -> None:
        try:
            parameters = self._read_form()
        except ValueError as error:
            messagebox.showerror("Проверьте данные", str(error))
            return

        self.generate_button.state(["disabled"])
        self.status.set("Подключаюсь к Google Sheets и формирую отчёт...")
        threading.Thread(
            target=self._generate_in_background,
            args=parameters,
            daemon=True,
        ).start()

    def _generate_in_background(
        self,
        start_date: date,
        end_date: date,
        department: str,
    ) -> None:
        try:
            credentials_path, spreadsheet_id = load_settings()
            sheets = GoogleSheets(credentials_path, spreadsheet_id)
            sheet_title = create_report(
                sheets,
                start_date,
                end_date,
                department,
            )
        except Exception as error:
            self.root.after(0, self._show_error, str(error))
            return
        self.root.after(0, self._show_success, sheet_title)

    def _show_success(self, sheet_title: str) -> None:
        self.generate_button.state(["!disabled"])
        message = f"Готово. В таблице создан лист «{sheet_title}»."
        self.status.set(message)
        messagebox.showinfo("Отчёт создан", message)

    def _show_error(self, error_message: str) -> None:
        self.generate_button.state(["!disabled"])
        self.status.set("Не удалось создать отчёт. Проверьте настройки и доступ.")
        messagebox.showerror("Ошибка", error_message)


def main() -> None:
    """Запускает оконное приложение."""
    root = tk.Tk()
    ReportGeneratorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

