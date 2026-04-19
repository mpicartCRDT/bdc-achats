from __future__ import annotations

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

        variables = {
            "week_number": generated_file.week_number or self._week_from_filename(generated_file.filename),
            "buyer_name": self.references.buyer_name,
            "supplier_name": generated_file.supplier_name,
        }

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
