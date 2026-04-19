from __future__ import annotations

import unittest

from repositories.parametrage_repository import ParametrageError, ParametrageRepository
from tests.helpers import make_parametrage_workbook


class ParametrageRepositoryTest(unittest.TestCase):
    def test_accepts_minimal_required_columns_without_optional_columns(self):
        references = ParametrageRepository().load(make_parametrage_workbook())

        self.assertEqual(references.buyer_code, "MATHIEU")
        self.assertEqual(references.buyer_name, "Mathieu Picart")
        self.assertEqual(references.buyer_email, "m.picart@example.test")
        self.assertEqual(references.format_code, "MATHIEU_V1")
        self.assertIn("SPI 1", references.plants_by_sheet)
        self.assertEqual(
            references.supplier_aliases["ag prog"], "AGRICOLA_PROGRES"
        )
        self.assertIn("FR", references.default_mail_templates_by_language)

    def test_blocks_only_when_required_column_is_missing(self):
        with self.assertRaises(ParametrageError) as ctx:
            ParametrageRepository().load(
                make_parametrage_workbook(include_required=False)
            )

        self.assertIn("Colonnes indispensables", str(ctx.exception))
        self.assertIn("code_format_ph", str(ctx.exception))

    def test_reads_optional_supplier_delay_codes_when_present(self):
        references = ParametrageRepository().load(
            make_parametrage_workbook(include_supplier_delay_columns=True)
        )

        supplier = references.suppliers_by_id["ROCA_DISTRIBUTION"]
        self.assertEqual(supplier.bdc_type, "Départ")
        self.assertEqual(supplier.delay_codes_by_plant["KB"], "A-B")
        self.assertEqual(supplier.delay_codes_by_plant["C9"], "A-C")
        self.assertEqual(supplier.delay_codes_by_plant["GS"], "A-D")
        self.assertEqual(supplier.delay_codes_by_plant["GNS"], "A-D")
        self.assertEqual(supplier.delay_codes_by_plant["VNC"], "A-B")
        self.assertEqual(supplier.quantity_mode, "standard")


if __name__ == "__main__":
    unittest.main()
