from __future__ import annotations

import re
import unicodedata
from typing import Any

from domain.models import ProductRef, ReferenceData, SupplierRef


def normalize_key(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(char for char in text if not unicodedata.combining(char))
    text = re.sub(r"\s+", " ", text)
    return text


class NormalizationService:
    def __init__(self, references: ReferenceData):
        self.references = references

    def resolve_supplier(self, raw_value: Any) -> tuple[SupplierRef | None, str | None]:
        key = normalize_key(raw_value)
        if not key:
            return None, "Fournisseur absent"

        direct_id = self._find_direct_supplier(key)
        supplier_id = direct_id or self.references.supplier_aliases.get(key)
        if supplier_id and supplier_id in self.references.suppliers_by_id:
            return self.references.suppliers_by_id[supplier_id], None
        return None, f"Fournisseur inconnu : {raw_value}"

    def resolve_product(self, raw_value: Any) -> tuple[ProductRef | None, str | None]:
        key = normalize_key(raw_value)
        if not key:
            return None, "Produit absent"

        direct_id = self._find_direct_product(key)
        product_id = direct_id or self.references.product_aliases.get(key)
        if product_id and product_id in self.references.products_by_id:
            return self.references.products_by_id[product_id], None
        return None, f"Produit inconnu : {raw_value}"

    def _find_direct_supplier(self, key: str) -> str | None:
        for supplier_id, supplier in self.references.suppliers_by_id.items():
            if key in {normalize_key(supplier_id), normalize_key(supplier.supplier_name)}:
                return supplier_id
        return None

    def _find_direct_product(self, key: str) -> str | None:
        for product_id, product in self.references.products_by_id.items():
            if key in {normalize_key(product_id), normalize_key(product.product_name)}:
                return product_id
        return None
