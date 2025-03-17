from pathlib import Path
from typing import Any
import logging

from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils import get_column_letter

from .exceptions import DataError
from .cell_utils import parse_cell_range

logger = logging.getLogger(__name__)

def read_excel_range(
    filepath: Path | str,
    sheet_name: str,
    start_cell: str = "A1",
    end_cell: str | None = None,
    preview_only: bool = False
) -> list[dict[str, Any]]:
    """Read data from Excel range with optional preview mode"""
    try:
        wb = load_workbook(filepath, read_only=True)
        
        if sheet_name not in wb.sheetnames:
            raise DataError(f"Sheet '{sheet_name}' not found")
            
        ws = wb[sheet_name]

        # Parse start cell
        if ':' in start_cell:
            start_cell, end_cell = start_cell.split(':')
            
        # Get start coordinates
        try:
            start_coords = parse_cell_range(f"{start_cell}:{start_cell}")
            if not start_coords or not all(coord is not None for coord in start_coords[:2]):
                raise DataError(f"Invalid start cell reference: {start_cell}")
            start_row, start_col = start_coords[0], start_coords[1]
        except ValueError as e:
            raise DataError(f"Invalid start cell format: {str(e)}")

        # Determine end coordinates
        if end_cell:
            try:
                end_coords = parse_cell_range(f"{end_cell}:{end_cell}")
                if not end_coords or not all(coord is not None for coord in end_coords[:2]):
                    raise DataError(f"Invalid end cell reference: {end_cell}")
                end_row, end_col = end_coords[0], end_coords[1]
            except ValueError as e:
                raise DataError(f"Invalid end cell format: {str(e)}")
        else:
            # For single cell, use same coordinates
            end_row, end_col = start_row, start_col

        # Validate range bounds
        if start_row > ws.max_row or start_col > ws.max_column:
            raise DataError(
                f"Start cell out of bounds. Sheet dimensions are "
                f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
            )

        data = []
        # If it's a single cell or single row, just read the values directly
        if start_row == end_row:
            row_data = {}
            for col in range(start_col, end_col + 1):
                cell = ws.cell(row=start_row, column=col)
                col_name = f"Column_{col}"
                row_data[col_name] = cell.value
            if any(v is not None for v in row_data.values()):
                data.append(row_data)
        else:
            # Multiple rows - use header row
            headers = []
            for col in range(start_col, end_col + 1):
                cell_value = ws.cell(row=start_row, column=col).value
                headers.append(str(cell_value) if cell_value is not None else f"Column_{col}")

            # Get data rows
            max_rows = min(start_row + 5, end_row) if preview_only else end_row
            for row in range(start_row + 1, max_rows + 1):
                row_data = {}
                for col, header in enumerate(headers, start=start_col):
                    cell = ws.cell(row=row, column=col)
                    row_data[header] = cell.value
                if any(v is not None for v in row_data.values()):
                    data.append(row_data)

        wb.close()
        return data
    except DataError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to read Excel range: {e}")
        raise DataError(str(e))

def write_data(
    filepath: str,
    sheet_name: str | None,
    data: list[dict[str, Any]] | None,
    start_cell: str = "A1",
    write_headers: bool = True,
) -> dict[str, str]:
    """Write data to Excel sheet with workbook handling"""
    try:
        if not data:
            raise DataError("No data provided to write")
            
        wb = load_workbook(filepath)

        # If no sheet specified, use active sheet
        if not sheet_name:
            sheet_name = wb.active.title
        elif sheet_name not in wb.sheetnames:
            wb.create_sheet(sheet_name)

        ws = wb[sheet_name]

        # Validate start cell
        try:
            start_coords = parse_cell_range(start_cell)
            if not start_coords or not all(coord is not None for coord in start_coords[:2]):
                raise DataError(f"Invalid start cell reference: {start_cell}")
        except ValueError as e:
            raise DataError(f"Invalid start cell format: {str(e)}")

        if len(data) > 0:
            # Check if first row of data contains headers
            first_row = data[0]
            has_headers = all(
                isinstance(value, str) and value.strip() == key.strip()
                for key, value in first_row.items()
            )
            
            # If first row contains headers, skip it when write_headers is True
            if has_headers and write_headers:
                data = data[1:]

            _write_data_to_worksheet(ws, data, start_cell, write_headers)

        wb.save(filepath)
        wb.close()

        return {"message": f"Data written to {sheet_name}", "active_sheet": sheet_name}
    except DataError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to write data: {e}")
        raise DataError(str(e))

def _write_data_to_worksheet(
    worksheet: Worksheet, 
    data: list[dict[str, Any]], 
    start_cell: str = "A1",
    write_headers: bool = True,
) -> None:
    """Write data to worksheet - internal helper function"""
    try:
        if not data:
            raise DataError("No data provided to write")

        try:
            start_coords = parse_cell_range(start_cell)
            if not start_coords or not all(x is not None for x in start_coords[:2]):
                raise DataError(f"Invalid start cell reference: {start_cell}")
            start_row, start_col = start_coords[0], start_coords[1]
        except ValueError as e:
            raise DataError(f"Invalid start cell format: {str(e)}")

        # Validate data structure
        if not all(isinstance(row, dict) for row in data):
            raise DataError("All data rows must be dictionaries")

        # Write headers if requested
        headers = list(data[0].keys())
        if write_headers:
            for i, header in enumerate(headers):
                cell = worksheet.cell(row=start_row, column=start_col + i)
                cell.value = header
                cell.font = Font(bold=True)
            start_row += 1  # Move start row down if headers were written

        # Write data
        for i, row_dict in enumerate(data):
            if not all(h in row_dict for h in headers):
                raise DataError(f"Row {i+1} is missing required headers")
            for j, header in enumerate(headers):
                cell = worksheet.cell(row=start_row + i, column=start_col + j)
                cell.value = row_dict.get(header, "")
    except DataError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to write worksheet data: {e}")
        raise DataError(str(e))
