"""Универсальный клиент для работы с Google Таблицами через сервисный аккаунт."""

from pathlib import Path
from typing import Any, Sequence

from google.oauth2 import service_account
from googleapiclient.discovery import Resource, build
from googleapiclient.errors import HttpError


SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


class GoogleSheetsError(RuntimeError):
    """Ошибка обращения к Google Sheets API."""


class GoogleSheets:
    """Выполняет чтение, запись и изменение указанной Google Таблицы."""

    def __init__(self, credentials_path: str | Path, spreadsheet_id: str) -> None:
        credentials_file = Path(credentials_path).expanduser().resolve()
        if not credentials_file.is_file():
            raise FileNotFoundError(
                f"JSON-ключ сервисного аккаунта не найден: {credentials_file}"
            )
        if not spreadsheet_id.strip():
            raise ValueError("Не указан SPREADSHEET_ID.")

        credentials = service_account.Credentials.from_service_account_file(
            str(credentials_file),
            scopes=SCOPES,
        )
        self.spreadsheet_id = spreadsheet_id.strip()
        self.service: Resource = build(
            "sheets",
            "v4",
            credentials=credentials,
            cache_discovery=False,
        )

    @staticmethod
    def quote_sheet_title(sheet_title: str) -> str:
        """Экранирует название листа для A1-диапазона."""
        return "'" + sheet_title.replace("'", "''") + "'"

    @staticmethod
    def _raise_api_error(action: str, error: HttpError) -> None:
        status = getattr(error.resp, "status", "неизвестен")
        raise GoogleSheetsError(
            f"Не удалось {action}. Google Sheets API вернул статус {status}: {error}"
        ) from error

    def get_metadata(self) -> dict[str, Any]:
        """Возвращает основные сведения о таблице и её листах."""
        try:
            return (
                self.service.spreadsheets()
                .get(
                    spreadsheetId=self.spreadsheet_id,
                    fields="properties.title,sheets.properties",
                )
                .execute()
            )
        except HttpError as error:
            self._raise_api_error("получить сведения о таблице", error)

    def get_spreadsheet_title(self) -> str:
        """Возвращает название Google Таблицы."""
        metadata = self.get_metadata()
        return str(metadata.get("properties", {}).get("title", ""))

    def get_sheet_names(self) -> list[str]:
        """Возвращает названия всех листов."""
        metadata = self.get_metadata()
        return [
            str(sheet["properties"]["title"])
            for sheet in metadata.get("sheets", [])
        ]

    def get_sheet_id(self, sheet_title: str) -> int:
        """Возвращает числовой идентификатор листа по его названию."""
        metadata = self.get_metadata()
        for sheet in metadata.get("sheets", []):
            properties = sheet.get("properties", {})
            if properties.get("title") == sheet_title:
                return int(properties["sheetId"])
        raise ValueError(f"Лист с названием «{sheet_title}» не найден.")

    def read_range(self, range_name: str) -> list[list[Any]]:
        """Читает значения из диапазона в A1-нотации."""
        try:
            result = (
                self.service.spreadsheets()
                .values()
                .get(
                    spreadsheetId=self.spreadsheet_id,
                    range=range_name,
                )
                .execute()
            )
            return result.get("values", [])
        except HttpError as error:
            self._raise_api_error(f"прочитать диапазон {range_name}", error)

    def read_all_cells(self, sheet_title: str | None = None) -> list[list[Any]]:
        """Читает все заполненные ячейки выбранного или первого листа."""
        selected_title = sheet_title
        if selected_title is None:
            sheet_names = self.get_sheet_names()
            if not sheet_names:
                return []
            selected_title = sheet_names[0]
        return self.read_range(self.quote_sheet_title(selected_title))

    def write_range(
        self,
        range_name: str,
        values: Sequence[Sequence[Any]],
        value_input_option: str = "RAW",
    ) -> dict[str, Any]:
        """Записывает двумерный список значений в заданный диапазон."""
        if value_input_option not in {"RAW", "USER_ENTERED"}:
            raise ValueError("value_input_option должен быть RAW или USER_ENTERED.")
        try:
            return (
                self.service.spreadsheets()
                .values()
                .update(
                    spreadsheetId=self.spreadsheet_id,
                    range=range_name,
                    valueInputOption=value_input_option,
                    body={"values": [list(row) for row in values]},
                )
                .execute()
            )
        except HttpError as error:
            self._raise_api_error(f"записать диапазон {range_name}", error)

    def write_ranges(
        self,
        ranges: Sequence[dict[str, Any]],
        value_input_option: str = "RAW",
    ) -> dict[str, Any]:
        """Записывает значения сразу в несколько диапазонов."""
        if value_input_option not in {"RAW", "USER_ENTERED"}:
            raise ValueError("value_input_option должен быть RAW или USER_ENTERED.")
        try:
            return (
                self.service.spreadsheets()
                .values()
                .batchUpdate(
                    spreadsheetId=self.spreadsheet_id,
                    body={
                        "valueInputOption": value_input_option,
                        "data": list(ranges),
                    },
                )
                .execute()
            )
        except HttpError as error:
            self._raise_api_error("записать несколько диапазонов", error)

    def update_cell(
        self,
        cell_address: str,
        value: Any,
        sheet_title: str | None = None,
        value_input_option: str = "RAW",
    ) -> dict[str, Any]:
        """Обновляет одну ячейку."""
        range_name = cell_address
        if sheet_title:
            range_name = f"{self.quote_sheet_title(sheet_title)}!{cell_address}"
        return self.write_range(range_name, [[value]], value_input_option)

    def append_row(
        self,
        values: Sequence[Any],
        sheet_title: str | None = None,
        value_input_option: str = "RAW",
    ) -> dict[str, Any]:
        """Добавляет строку после последней заполненной строки листа."""
        selected_title = sheet_title
        if selected_title is None:
            sheet_names = self.get_sheet_names()
            if not sheet_names:
                raise ValueError("Таблица не содержит листов.")
            selected_title = sheet_names[0]
        range_name = self.quote_sheet_title(selected_title)
        try:
            return (
                self.service.spreadsheets()
                .values()
                .append(
                    spreadsheetId=self.spreadsheet_id,
                    range=range_name,
                    valueInputOption=value_input_option,
                    insertDataOption="INSERT_ROWS",
                    body={"values": [list(values)]},
                )
                .execute()
            )
        except HttpError as error:
            self._raise_api_error(f"добавить строку на лист {selected_title}", error)

    def clear_range(self, range_name: str) -> dict[str, Any]:
        """Очищает значения в диапазоне без удаления форматирования."""
        try:
            return (
                self.service.spreadsheets()
                .values()
                .clear(
                    spreadsheetId=self.spreadsheet_id,
                    range=range_name,
                    body={},
                )
                .execute()
            )
        except HttpError as error:
            self._raise_api_error(f"очистить диапазон {range_name}", error)

    def create_sheet(self, sheet_title: str) -> int:
        """Создаёт лист и возвращает его числовой идентификатор."""
        if not sheet_title.strip():
            raise ValueError("Название листа не может быть пустым.")
        try:
            result = (
                self.service.spreadsheets()
                .batchUpdate(
                    spreadsheetId=self.spreadsheet_id,
                    body={
                        "requests": [
                            {"addSheet": {"properties": {"title": sheet_title}}}
                        ]
                    },
                )
                .execute()
            )
            return int(result["replies"][0]["addSheet"]["properties"]["sheetId"])
        except HttpError as error:
            self._raise_api_error(f"создать лист «{sheet_title}»", error)

    def delete_row(self, row_number: int, sheet_title: str) -> dict[str, Any]:
        """Удаляет строку листа по номеру, начиная с единицы."""
        if row_number < 1:
            raise ValueError("Номер строки должен быть не меньше 1.")
        sheet_id = self.get_sheet_id(sheet_title)
        return self.batch_update(
            [
                {
                    "deleteDimension": {
                        "range": {
                            "sheetId": sheet_id,
                            "dimension": "ROWS",
                            "startIndex": row_number - 1,
                            "endIndex": row_number,
                        }
                    }
                }
            ]
        )

    def delete_sheet(self, sheet_title: str) -> dict[str, Any]:
        """Удаляет лист по его названию."""
        return self.batch_update(
            [{"deleteSheet": {"sheetId": self.get_sheet_id(sheet_title)}}]
        )

    def batch_update(self, requests: Sequence[dict[str, Any]]) -> dict[str, Any]:
        """Выполняет пакет структурных и форматирующих операций."""
        if not requests:
            return {"replies": []}
        try:
            return (
                self.service.spreadsheets()
                .batchUpdate(
                    spreadsheetId=self.spreadsheet_id,
                    body={"requests": list(requests)},
                )
                .execute()
            )
        except HttpError as error:
            self._raise_api_error("выполнить пакетное обновление", error)

