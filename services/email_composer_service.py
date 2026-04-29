from __future__ import annotations

from datetime import date, timedelta
import re
from typing import Any

from domain.models import EmailAttachment, GeneratedBDCFile, PreparedEmailPayload, ReferenceData
from services.normalization_service import normalize_key


class EmailComposerService:
    DEFAULT_LANGUAGE = "FR"

    def __init__(self, references: ReferenceData):
        self.references = references

    def compose_for_files(self, files: list[GeneratedBDCFile]) -> list[PreparedEmailPayload]:
        return [self.compose_for_file(file) for file in files]

    def compose_for_file(self, generated_file: GeneratedBDCFile) -> PreparedEmailPayload:
        supplier = self.references.suppliers_by_id.get(generated_file.supplier_id)
        if supplier is None:
            return self._error_payload(
                generated_file,
                "Fournisseur introuvable dans le paramétrage.",
            )

        template = self._template_for_supplier(generated_file.supplier_id, supplier.language)
        reasons: list[str] = []
        if not supplier.email_to:
            reasons.append("Email destinataire manquant.")
        if template is None:
            reasons.append("Modèle mail manquant.")
        if not generated_file.content:
            reasons.append("Fichier BDC manquant.")

        variables = self._template_variables(generated_file)

        subject = self._render(template.subject_template, variables) if template else None
        body = self._render(template.body_template, variables) if template else None
        attachment = (
            EmailAttachment(
                filename=generated_file.filename,
                content=generated_file.content,
                mime_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                path=None,
            )
            if generated_file.content
            else None
        )

        return PreparedEmailPayload(
            supplier_id=generated_file.supplier_id,
            supplier_name=generated_file.supplier_name,
            to_recipients=supplier.email_to,
            cc_recipients=supplier.email_cc,
            bcc_recipients=None,
            subject=subject,
            body=body,
            attachment=attachment,
            needs_review=bool(reasons),
            review_reason="; ".join(reasons) if reasons else None,
        )

    def _template_for_supplier(self, supplier_id: str, language: str | None):
        if supplier_id in self.references.mail_templates_by_supplier:
            return self.references.mail_templates_by_supplier[supplier_id]

        lang = normalize_key(language or self.DEFAULT_LANGUAGE).upper()
        return (
            self.references.default_mail_templates_by_language.get(lang)
            or self.references.default_mail_templates_by_language.get(self.DEFAULT_LANGUAGE)
        )

    def _render(self, template: str, variables: dict[str, Any]) -> str:
        rendered = template
        for key, value in variables.items():
            rendered = rendered.replace("{{" + key + "}}", "" if value is None else str(value))
        return rendered

    def _template_variables(self, generated_file: GeneratedBDCFile) -> dict[str, Any]:
        week_number = generated_file.week_number or self._week_from_filename(generated_file.filename)
        delivery_monday = generated_file.delivery_week_monday
        delivery_saturday = (
            delivery_monday + timedelta(days=5)
            if delivery_monday is not None
            else None
        )
        bdc_dates = generated_file.bdc_dates

        variables: dict[str, Any] = {
            "week_number": week_number,
            "week_number_2digits": f"{week_number:02d}" if week_number is not None else None,
            "week": week_number,
            "semaine": week_number,
            "semaine_2_chiffres": f"{week_number:02d}" if week_number is not None else None,
            "buyer_name": self.references.buyer_name,
            "supplier_name": generated_file.supplier_name,
            "date": self._format_date(delivery_monday),
            "date_debut": self._format_date(delivery_monday),
            "date_fin": self._format_date(delivery_saturday),
            "date_livraison_debut": self._format_date(delivery_monday),
            "date_livraison_fin": self._format_date(delivery_saturday),
            "date_start": self._format_date(delivery_monday),
            "date_end": self._format_date(delivery_saturday),
            "delivery_start": self._format_date(delivery_monday),
            "delivery_end": self._format_date(delivery_saturday),
            "date_lundi": self._format_date(delivery_monday),
            "date_mardi": self._format_date_from_monday(delivery_monday, 1),
            "date_mercredi": self._format_date_from_monday(delivery_monday, 2),
            "date_jeudi": self._format_date_from_monday(delivery_monday, 3),
            "date_vendredi": self._format_date_from_monday(delivery_monday, 4),
            "date_samedi": self._format_date(delivery_saturday),
        }

        block_names = [
            "date_samedi_s_moins_1",
            "date_lundi",
            "date_mardi",
            "date_mercredi",
            "date_jeudi",
            "date_vendredi",
            "date_samedi",
            "date_lundi_s_plus_1",
        ]
        for key, value in zip(block_names, bdc_dates):
            variables[key] = self._format_date(value)
        return variables

    def _format_date_from_monday(self, monday: date | None, offset_days: int) -> str | None:
        if monday is None:
            return None
        return self._format_date(monday + timedelta(days=offset_days))

    def _format_date(self, value: date | None) -> str | None:
        if value is None:
            return None
        return value.strftime("%d/%m/%Y")

    def _week_from_filename(self, filename: str) -> int | None:
        match = re.search(r"(?:_|\s)S(\d+)", filename)
        return int(match.group(1)) if match else None

    def _error_payload(self, generated_file: GeneratedBDCFile, reason: str) -> PreparedEmailPayload:
        return PreparedEmailPayload(
            supplier_id=generated_file.supplier_id,
            supplier_name=generated_file.supplier_name,
            to_recipients=None,
            cc_recipients=None,
            bcc_recipients=None,
            subject=None,
            body=None,
            attachment=None,
            needs_review=True,
            review_reason=reason,
        )
