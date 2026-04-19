from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, BinaryIO

import pandas as pd

from domain.models import NormalizedPHRow, ParseResult, ReferenceData
from services.normalization_service import NormalizationService, normalize_key


@dataclass(frozen=True)
class DayBlock:
    day_label: str | None
    date_value: date | datetime | None
    supplier_col: int
    qty_col: int | None
    weight_col: int | None
    transport_col: int | None


class MathieuParser:
    TECHNICAL_LABELS = {"total", "totaux"}

    def __init__(self, references: ReferenceData):
        self.references = references
        self.normalizer = NormalizationService(references)

    def parse(self, source: str | Path | BinaryIO, source_filename: str) -> ParseResult:
        self._rewind(source)
        sheets = pd.read_excel(source, sheet_name=None, header=None, engine="openpyxl")
        detected_sheets = list(sheets.keys())
        expected_sheets = self.references.plants_by_sheet
        processed_sheets = [name for name in detected_sheets if name in expected_sheets]
        ignored_sheets = [name for name in detected_sheets if name not in expected_sheets]

        rows: list[NormalizedPHRow] = []
        for sheet_name in processed_sheets:
            sheet_rows = self._parse_sheet(
                df=sheets[sheet_name],
                sheet_name=sheet_name,
                source_filename=source_filename,
            )
            rows.extend(sheet_rows)

        return ParseResult(
            rows=rows,
            detected_sheets=detected_sheets,
            processed_sheets=processed_sheets,
            ignored_sheets=ignored_sheets,
        )

    def _parse_sheet(
        self, df: pd.DataFrame, sheet_name: str, source_filename: str
    ) -> list[NormalizedPHRow]:
        header_idx = self._find_header_row(df)
        if header_idx is None:
            return []

        day_row_idx = max(0, header_idx - 1)
        header_row = df.iloc[header_idx]
        day_row = df.iloc[day_row_idx]
        blocks = self._build_day_blocks(header_row, day_row)
        plant = self.references.plants_by_sheet[sheet_name]

        rows: list[NormalizedPHRow] = []
        current_product: Any = None
        current_product_valid = False

        for row_idx in range(header_idx + 1, len(df)):
            row = df.iloc[row_idx]
            product_cell = self._cell(row, 0)

            if self._is_blank_row(row, blocks):
                current_product = None
                current_product_valid = False
                continue

            if self._is_total_label(product_cell):
                current_product = None
                current_product_valid = False
                continue

            if self._is_unit_row(row, blocks):
                current_product = None
                current_product_valid = False
                continue

            useful_blocks = [block for block in blocks if self._has_block_content(row, block)]
            if self._is_section_break(product_cell, useful_blocks):
                current_product = None
                current_product_valid = False
                continue

            inherited = False
            if self._has_text(product_cell):
                current_product = product_cell
                current_product_valid = True
            elif useful_blocks and current_product_valid:
                product_cell = current_product
                inherited = True

            if not useful_blocks:
                continue

            for block in useful_blocks:
                normalized = self._build_row(
                    source_filename=source_filename,
                    sheet_name=sheet_name,
                    excel_row_index=row_idx + 1,
                    plant_code=plant.plant_code,
                    plant_name=plant.plant_name,
                    day_block=block,
                    product_raw=product_cell,
                    row=row,
                    inherited_product=inherited,
                )
                rows.append(normalized)

        return rows

    def _find_header_row(self, df: pd.DataFrame) -> int | None:
        for idx, row in df.iterrows():
            normalized = [normalize_key(value) for value in row]
            if "fournisseur" in normalized and "nbre" in normalized:
                return int(idx)
        return None

    def _build_day_blocks(self, header_row: pd.Series, day_row: pd.Series) -> list[DayBlock]:
        blocks: list[DayBlock] = []
        for col_idx, value in enumerate(header_row):
            if normalize_key(value) != "fournisseur":
                continue
            qty_col = self._find_named_col(header_row, col_idx + 1, ["nbre"])
            weight_col = self._find_named_col(header_row, col_idx + 1, ["poids"])
            transport_col = self._find_named_col(header_row, col_idx + 1, ["trpt"])
            day_label, date_value = self._read_day_values(day_row, col_idx)
            blocks.append(
                DayBlock(
                    day_label=day_label,
                    date_value=date_value,
                    supplier_col=col_idx,
                    qty_col=qty_col,
                    weight_col=weight_col,
                    transport_col=transport_col,
                )
            )
        return blocks

    def _find_named_col(
        self, header_row: pd.Series, start_col: int, accepted_names: list[str]
    ) -> int | None:
        stop_col = min(start_col + 4, len(header_row))
        for col_idx in range(start_col, stop_col):
            if normalize_key(header_row.iloc[col_idx]) in accepted_names:
                return col_idx
        return None

    def _read_day_values(
        self, day_row: pd.Series, supplier_col: int
    ) -> tuple[str | None, date | datetime | None]:
        day_label = None
        date_value = None
        for col_idx in range(max(0, supplier_col - 1), min(supplier_col + 3, len(day_row))):
            value = day_row.iloc[col_idx]
            if isinstance(value, (datetime, date)):
                date_value = value
            elif self._has_text(value):
                day_label = str(value).strip()
        return day_label, date_value

    def _build_row(
        self,
        source_filename: str,
        sheet_name: str,
        excel_row_index: int,
        plant_code: str,
        plant_name: str,
        day_block: DayBlock,
        product_raw: Any,
        row: pd.Series,
        inherited_product: bool,
    ) -> NormalizedPHRow:
        supplier_raw = self._cell(row, day_block.supplier_col)
        qty_raw = self._cell(row, day_block.qty_col)
        weight_raw = self._cell(row, day_block.weight_col)
        transport_value = self._cell(row, day_block.transport_col)

        supplier, supplier_error = self.normalizer.resolve_supplier(supplier_raw)
        product, product_error = self.normalizer.resolve_product(product_raw)
        qty_value, qty_unit, qty_error, info_text = self._choose_quantity(qty_raw, weight_raw)

        reasons = [
            reason
            for reason in [supplier_error, product_error, qty_error]
            if reason is not None
        ]
        if inherited_product and product is None:
            reasons.append("Produit hérité non reconnu")

        return NormalizedPHRow(
            source_filename=source_filename,
            source_sheet=sheet_name,
            source_row_index=excel_row_index,
            buyer_code=self.references.buyer_code,
            plant_code=plant_code,
            plant_name=plant_name,
            date_source=day_block.date_value,
            day_label_source=day_block.day_label,
            supplier_raw=supplier_raw,
            supplier_id=supplier.supplier_id if supplier else None,
            supplier_name=supplier.supplier_name if supplier else None,
            product_raw=product_raw,
            product_id=product.product_id if product else None,
            product_name=product.product_name if product else None,
            qty_value=qty_value,
            qty_unit=qty_unit,
            qty_nbre_raw=qty_raw,
            qty_poids_raw=weight_raw,
            transport_value=transport_value,
            info_text=info_text,
            needs_review=bool(reasons),
            review_reason="; ".join(reasons) if reasons else None,
        )

    def _choose_quantity(
        self, qty_raw: Any, weight_raw: Any
    ) -> tuple[float | None, str | None, str | None, str | None]:
        weight = self._to_number(weight_raw)
        if weight is not None:
            return weight, "kg", None, None

        qty = self._to_number(qty_raw)
        if qty is not None:
            return qty, "pal", None, None

        text_values = [
            str(value).strip()
            for value in [weight_raw, qty_raw]
            if self._has_text(value)
        ]
        if text_values:
            info_text = " | ".join(text_values)
            return None, None, f"Quantité non numérique : {info_text}", info_text
        return None, None, "Quantité absente", None

    def _has_block_content(self, row: pd.Series, block: DayBlock) -> bool:
        values = [
            self._cell(row, block.supplier_col),
            self._cell(row, block.qty_col),
            self._cell(row, block.weight_col),
            self._cell(row, block.transport_col),
        ]
        return any(not self._is_empty(value) for value in values)

    def _is_blank_row(self, row: pd.Series, blocks: list[DayBlock]) -> bool:
        product_cell = self._cell(row, 0)
        if not self._is_empty(product_cell):
            return False
        return not any(self._has_block_content(row, block) for block in blocks)

    def _is_unit_row(self, row: pd.Series, blocks: list[DayBlock]) -> bool:
        product_cell = self._cell(row, 0)
        if not self._is_empty(product_cell):
            return False

        saw_unit = False
        for block in blocks:
            supplier = self._cell(row, block.supplier_col)
            qty = normalize_key(self._cell(row, block.qty_col))
            weight = self._cell(row, block.weight_col)
            transport = self._cell(row, block.transport_col)
            if not self._is_empty(supplier) or not self._is_empty(weight) or not self._is_empty(transport):
                return False
            if qty in {"pal", "kg"}:
                saw_unit = True
            elif qty:
                return False
        return saw_unit

    def _is_section_break(self, product_cell: Any, useful_blocks: list[DayBlock]) -> bool:
        if not self._has_text(product_cell):
            return False
        text = str(product_cell).strip()
        key = normalize_key(text)
        if key in self.TECHNICAL_LABELS:
            return True
        if not useful_blocks and text.upper() == text and len(text) > 2:
            return True
        return False

    def _is_total_label(self, value: Any) -> bool:
        return normalize_key(value) in self.TECHNICAL_LABELS

    def _to_number(self, value: Any) -> float | None:
        if self._is_empty(value):
            return None
        if isinstance(value, (int, float)) and not pd.isna(value):
            return float(value)
        if isinstance(value, str):
            text = value.strip().replace(",", ".")
            try:
                return float(text)
            except ValueError:
                return None
        return None

    def _cell(self, row: pd.Series, col_idx: int | None) -> Any:
        if col_idx is None or col_idx >= len(row):
            return None
        value = row.iloc[col_idx]
        return None if self._is_empty(value) else value

    def _is_empty(self, value: Any) -> bool:
        if value is None:
            return True
        try:
            return bool(pd.isna(value))
        except TypeError:
            return False

    def _has_text(self, value: Any) -> bool:
        return isinstance(value, str) and bool(value.strip())

    def _rewind(self, source: str | Path | BinaryIO) -> None:
        if hasattr(source, "seek"):
            source.seek(0)
