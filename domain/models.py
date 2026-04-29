from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
from typing import Any


@dataclass(frozen=True)
class SupplierRef:
    supplier_id: str
    supplier_name: str
    generate_bdc: bool = True
    ignore_if_detected: bool = False
    bdc_type: str | None = None
    delay_codes_by_plant: dict[str, str] = field(default_factory=dict)
    quantity_mode: str | None = None
    model_code: str | None = None
    email_to: str | None = None
    email_cc: str | None = None
    language: str | None = None


@dataclass(frozen=True)
class ProductRef:
    product_id: str
    product_name: str


@dataclass(frozen=True)
class PlantRef:
    sheet_name: str
    plant_code: str
    plant_name: str


@dataclass(frozen=True)
class ReferenceData:
    buyer_code: str
    buyer_name: str
    buyer_email: str | None
    format_code: str
    plants_by_sheet: dict[str, PlantRef]
    suppliers_by_id: dict[str, SupplierRef]
    supplier_aliases: dict[str, str]
    products_by_id: dict[str, ProductRef]
    product_aliases: dict[str, str]
    mail_templates_by_supplier: dict[str, "MailTemplateRef"] = field(default_factory=dict)
    default_mail_templates_by_language: dict[str, "MailTemplateRef"] = field(default_factory=dict)


@dataclass(frozen=True)
class MailTemplateRef:
    template_code: str
    supplier_id: str
    subject_template: str
    body_template: str


@dataclass
class NormalizedPHRow:
    source_filename: str
    source_sheet: str
    source_row_index: int
    buyer_code: str
    plant_code: str
    plant_name: str
    date_source: date | datetime | None
    day_label_source: str | None
    source_week_number: int | None
    supplier_raw: Any
    supplier_id: str | None
    supplier_name: str | None
    product_raw: Any
    product_id: str | None
    product_name: str | None
    qty_value: float | None
    qty_unit: str | None
    transport_value: Any
    qty_nbre_raw: Any = None
    qty_poids_raw: Any = None
    info_text: str | None = None
    needs_review: bool = False
    review_reason: str | None = None

    def as_dict(self) -> dict[str, Any]:
        return {
            "source_filename": self.source_filename,
            "source_sheet": self.source_sheet,
            "source_row_index": self.source_row_index,
            "buyer_code": self.buyer_code,
            "plant_code": self.plant_code,
            "plant_name": self.plant_name,
            "date_source": self.date_source,
            "source_week_number": self.source_week_number,
            "day_label_source": self.day_label_source,
            "supplier_raw": self.supplier_raw,
            "supplier_id": self.supplier_id,
            "supplier_name": self.supplier_name,
            "product_raw": self.product_raw,
            "product_id": self.product_id,
            "product_name": self.product_name,
            "qty_nbre_raw": self.qty_nbre_raw,
            "qty_poids_raw": self.qty_poids_raw,
            "qty_value": self.qty_value,
            "qty_unit": self.qty_unit,
            "transport_value": self.transport_value,
            "info_text": self.info_text,
            "needs_review": self.needs_review,
            "review_reason": self.review_reason,
        }


@dataclass
class ParseResult:
    rows: list[NormalizedPHRow] = field(default_factory=list)
    detected_sheets: list[str] = field(default_factory=list)
    processed_sheets: list[str] = field(default_factory=list)
    ignored_sheets: list[str] = field(default_factory=list)


@dataclass
class GeneratedBDCFile:
    supplier_id: str
    supplier_name: str
    filename: str
    content: bytes
    injected_rows: int
    skipped_rows: int
    week_number: int | None = None
    delivery_week_monday: date | None = None
    bdc_dates: list[date] = field(default_factory=list)


@dataclass
class BDCGenerationResult:
    files: list[GeneratedBDCFile] = field(default_factory=list)
    skipped_rows: list[NormalizedPHRow] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


@dataclass
class EmailAttachment:
    filename: str
    content: bytes
    mime_type: str
    path: str | None = None


@dataclass
class PreparedEmailPayload:
    supplier_id: str
    supplier_name: str
    to_recipients: str | None
    cc_recipients: str | None
    bcc_recipients: str | None
    subject: str | None
    body: str | None
    attachment: EmailAttachment | None
    needs_review: bool = False
    review_reason: str | None = None
