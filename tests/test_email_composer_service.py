from __future__ import annotations

import unittest

from domain.models import GeneratedBDCFile
from repositories.parametrage_repository import ParametrageRepository
from services.email_composer_service import EmailComposerService
from tests.helpers import make_parametrage_workbook


class EmailComposerServiceTest(unittest.TestCase):
    def setUp(self):
        self.references = ParametrageRepository().load(make_parametrage_workbook())
        supplier = self.references.suppliers_by_id["AGRICOLA_PROGRES"]
        self.references.suppliers_by_id["AGRICOLA_PROGRES"] = supplier.__class__(
            **{**supplier.__dict__, "email_to": "client@example.test", "email_cc": "cc@example.test"}
        )
        self.service = EmailComposerService(self.references)

    def test_composes_payload_with_template_variables_and_attachment(self):
        payload = self.service.compose_for_file(
            GeneratedBDCFile(
                supplier_id="AGRICOLA_PROGRES",
                supplier_name="Agricola Progres",
                filename="Agricola Progres S16.xlsx",
                content=b"xlsx",
                injected_rows=1,
                skipped_rows=0,
                week_number=16,
            )
        )

        self.assertFalse(payload.needs_review)
        self.assertEqual(payload.to_recipients, "client@example.test")
        self.assertEqual(payload.cc_recipients, "cc@example.test")
        self.assertEqual(payload.subject, "Programme semaine 16")
        self.assertIn("Agricola Progres", payload.body)
        self.assertIn("Mathieu Picart", payload.body)
        self.assertEqual(payload.attachment.filename, "Agricola Progres S16.xlsx")

    def test_flags_missing_email(self):
        payload = self.service.compose_for_file(
            GeneratedBDCFile(
                supplier_id="ROCA_DISTRIBUTION",
                supplier_name="Roca Distribution",
                filename="Roca Distribution S16.xlsx",
                content=b"xlsx",
                injected_rows=1,
                skipped_rows=0,
                week_number=16,
            )
        )

        self.assertTrue(payload.needs_review)
        self.assertIn("Email destinataire manquant", payload.review_reason)


if __name__ == "__main__":
    unittest.main()
