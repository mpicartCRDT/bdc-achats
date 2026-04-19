from __future__ import annotations

import unittest

from repositories.parametrage_repository import ParametrageRepository
from services.normalization_service import NormalizationService
from tests.helpers import make_parametrage_workbook


class NormalizationServiceTest(unittest.TestCase):
    def setUp(self):
        references = ParametrageRepository().load(make_parametrage_workbook())
        self.service = NormalizationService(references)

    def test_resolves_supplier_aliases(self):
        cases = {
            "Ag Prog": "AGRICOLA_PROGRES",
            "GSC": "GROUPE_SAVEUR_CLOS",
            "Roca": "ROCA_DISTRIBUTION",
        }
        for raw, expected in cases.items():
            supplier, error = self.service.resolve_supplier(raw)
            self.assertIsNone(error)
            self.assertEqual(supplier.supplier_id, expected)

    def test_resolves_product_aliases(self):
        cases = {
            "Frisée": "FRISEE",
            "F de Ch Rouge": "FDC_ROUGE",
            "Tomate pulpa": "TOMATE_PULPA",
        }
        for raw, expected in cases.items():
            product, error = self.service.resolve_product(raw)
            self.assertIsNone(error)
            self.assertEqual(product.product_id, expected)

    def test_unknown_values_return_review_errors(self):
        supplier, supplier_error = self.service.resolve_supplier("Mystere")
        product, product_error = self.service.resolve_product("Produit mystere")

        self.assertIsNone(supplier)
        self.assertIn("Fournisseur inconnu", supplier_error)
        self.assertIsNone(product)
        self.assertIn("Produit inconnu", product_error)


if __name__ == "__main__":
    unittest.main()
