from __future__ import annotations

import argparse
import re
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter


ROOT = Path(__file__).resolve().parents[1]
INPUT = ROOT / "entrada"
OUTPUT = ROOT / "salida"


def normalize_login(value: object) -> str:
    text = "" if value is None else str(value).strip()
    return re.sub(r"W\d+$", "", text, flags=re.IGNORECASE).upper()


def read_catalog(path: Path) -> dict[str, str]:
    book = load_workbook(path, read_only=True, data_only=True)
    sheet = book["Internet"]
    mapping: dict[str, str] = {}
    for description, code in sheet.iter_rows(
        min_row=2, min_col=1, max_col=2, values_only=True
    ):
        if description and code:
            mapping[str(description).strip()] = str(code).strip().upper()
    book.close()
    return mapping


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Cruza grupos de internet con roles SGA."
    )
    parser.add_argument(
        "--fuera", type=Path, default=INPUT / "Fuera de su grupo - internet.xlsx"
    )
    parser.add_argument(
        "--catalogo", type=Path, default=INPUT / "Internet_codigos.xlsx"
    )
    parser.add_argument("--sga", type=Path, default=INPUT / "SGA - INTERNET.xlsx")
    parser.add_argument(
        "--salida", type=Path, default=OUTPUT / "analisis_internet.xlsx"
    )
    args = parser.parse_args()

    catalog = read_catalog(args.catalogo)

    sga_book = load_workbook(args.sga, read_only=True, data_only=True)
    sga_sheet = sga_book.active
    sga_by_login: dict[str, list[tuple[str, str]]] = {}
    for row in sga_sheet.iter_rows(min_row=2, values_only=True):
        if len(row) < 9 or not row[1]:
            continue
        login = str(row[1]).strip().lower()
        role = "" if row[8] is None else str(row[8]).strip().upper()
        name = "" if row[2] is None else str(row[2]).strip()
        sga_by_login.setdefault(login, []).append((role, name))
    sga_book.close()

    source_book = load_workbook(args.fuera, read_only=True, data_only=True)
    if "USUARIOS" not in source_book.sheetnames:
        raise ValueError(f"No existe la hoja 'USUARIOS' en {args.fuera}")
    source_sheet = source_book["USUARIOS"]
    results: list[list[object]] = []
    for row in source_sheet.iter_rows(min_row=2, values_only=True):
        if len(row) < 4 or not row[1]:
            continue
        number, original, ip, description = row[:4]
        original = str(original).strip()
        normalized = normalize_login(original)
        description = "" if description is None else str(description).strip()
        converted = catalog.get(description, "")
        matches = sga_by_login.get(normalized.lower(), [])
        if not matches:
            results.append(
                [
                    number,
                    original,
                    ip,
                    description,
                    normalized,
                    converted,
                    "",
                    "",
                    "NO ENCONTRADO EN SGA",
                ]
            )
            continue
        for role, _name in matches:
            expected = role
            status = "COINCIDE" if converted and role == converted else "NO COINCIDE"
            if not converted:
                status = "DESCRIPCIÓN SIN CÓDIGO"
            results.append(
                [
                    number,
                    original,
                    ip,
                    description,
                    normalized,
                    converted,
                    expected,
                    converted,
                    status,
                ]
            )
    source_book.close()

    args.salida.parent.mkdir(parents=True, exist_ok=True)
    out = Workbook()
    sheet = out.active
    sheet.title = "dispositivos-sga"
    sheet.append(
        [
            "N°",
            "Nombre sin normalizar",
            "Dirección IP",
            "Nombre de grupo",
            "Nombre normalizado",
            "Descripción convertida a código",
            "Valor esperado (rol SGA)",
            "Valor encontrado (código)",
            "Resultado",
        ]
    )
    for result in results:
        sheet.append(result)
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
    widths = [8, 25, 17, 68, 22, 28, 25, 25, 24]
    for col, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(col)].width = width
    out.save(args.salida)
    print(f"Generado: {args.salida} ({len(results)} filas)")


if __name__ == "__main__":
    main()
