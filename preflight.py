"""Живая проверка авторизации и доступа к заданной Google Таблице."""

from report_generator import load_settings
from google_sheets import GoogleSheets


def main() -> None:
    """Проверяет подключение, метаданные и чтение первой ячейки."""
    credentials_path, spreadsheet_id = load_settings()
    sheets = GoogleSheets(credentials_path, spreadsheet_id)

    metadata = sheets.get_metadata()
    spreadsheet_title = metadata.get("properties", {}).get("title", "Без названия")
    sheet_names = [
        item.get("properties", {}).get("title", "Без названия")
        for item in metadata.get("sheets", [])
    ]

    if not sheet_names:
        raise RuntimeError("В Google Таблице не найдено ни одного листа.")

    first_sheet = sheet_names[0]
    first_cell = sheets.read_range(
        f"{sheets.quote_sheet_title(first_sheet)}!A1"
    )

    print("Проверка Google Sheets API выполнена успешно.")
    print(f"Таблица: {spreadsheet_title}")
    print(f"Листы: {', '.join(sheet_names)}")
    print(f"Чтение {first_sheet}!A1: {'есть значение' if first_cell else 'ячейка пуста'}")


if __name__ == "__main__":
    try:
        main()
    except Exception as error:
        print(f"Проверка не пройдена: {error}")
        raise SystemExit(1) from error

