from __future__ import annotations

import unittest
from datetime import date

from domain.models import GeneratedBDCFile, MailTemplateRef
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

    def test_composes_payload_with_week_and_date_variables_from_generated_bdc(self):
        self.references.default_mail_templates_by_language["FR"] = MailTemplateRef(
            template_code="MT_DEFAULT_FR",
            supplier_id="*",
            subject_template="Programme S{{week_number_2digits}} du {{date_debut}}",
            body_template=(
                "Bonjour {{supplier_name}},\n"
                "Livraisons du {{date_lundi}} au {{date_samedi}}. "
                "BDC du {{date_samedi_s_moins_1}} au {{date_lundi_s_plus_1}}."
            ),
        )

        payload = self.service.compose_for_file(
            GeneratedBDCFile(
                supplier_id="AGRICOLA_PROGRES",
                supplier_name="Agricola Progres",
                filename="Agricola Progres S16.xlsx",
                content=b"xlsx",
                injected_rows=1,
                skipped_rows=0,
                week_number=16,
                delivery_week_monday=date(2026, 4, 13),
                bdc_dates=[
                    date(2026, 4, 11),
                    date(2026, 4, 13),
                    date(2026, 4, 14),
                    date(2026, 4, 15),
                    date(2026, 4, 16),
                    date(2026, 4, 17),
                    date(2026, 4, 18),
                    date(2026, 4, 20),
                ],
            )
        )

        self.assertEqual(payload.subject, "Programme S16 du 13/04/2026")
        self.assertIn("Livraisons du 13/04/2026 au 18/04/2026", payload.body)
        self.assertIn("BDC du 11/04/2026 au 20/04/2026", payload.body)

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
