from __future__ import annotations

import json
from zipfile import ZIP_DEFLATED, ZipFile
from io import BytesIO

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

from parsers.factory import get_parser
from repositories.parametrage_repository import ParametrageError, ParametrageRepository
from services.bdc_generation_service import BDCGenerationError, BDCGenerationService
from services.email_composer_service import EmailComposerService


CONTROL_COLUMNS = [
    "source_sheet",
    "source_row_index",
    "plant_code",
    "plant_name",
    "date_source",
    "day_label_source",
    "supplier_raw",
    "supplier_id",
    "supplier_name",
    "product_raw",
    "product_id",
    "product_name",
    "qty_value",
    "qty_unit",
    "qty_nbre_raw",
    "qty_poids_raw",
    "transport_value",
    "info_text",
    "needs_review",
    "review_reason",
]


def main() -> None:
    st.set_page_config(page_title="Génération de BDC", layout="wide")
    st.title("Génération de BDC")
    st.caption("Lecture PH, génération BDC et préparation manuelle des emails.")

    parametrage_file = st.file_uploader(
        "Charger l'excel de paramétrage", type=["xlsx"], key="parametrage"
    )
    ph_file = st.file_uploader("Charger le programme", type=["xlsx"], key="ph")
    template_file = st.file_uploader(
        "Charger le BDC type",
        type=["xlsx"],
        key="template",
        help="Chargez le BDC type utilisé pour générer les fichiers.",
    )

    if parametrage_file is None or ph_file is None:
        st.info("Chargez l'excel de paramétrage puis le programme pour lancer le contrôle.")
        return

    try:
        references = ParametrageRepository().load(parametrage_file)
    except ParametrageError as exc:
        st.error(str(exc))
        return
    except Exception:
        st.error("Le fichier de paramétrage n'a pas pu être lu. Vérifiez le fichier Excel.")
        return

    try:
        parser = get_parser(references.format_code, references)
        result = parser.parse(ph_file, source_filename=ph_file.name)
    except Exception:
        st.error("La PH n'a pas pu être lue. Vérifiez qu'il s'agit bien du format Mathieu.")
        return

    st.subheader("Synthèse")
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Acheteur", references.buyer_name)
    col2.metric("Fichier", ph_file.name)
    col3.metric("Lignes extraites", len(result.rows))
    col4.metric("Lignes à revoir", sum(row.needs_review for row in result.rows))

    st.write("Feuilles détectées :", ", ".join(result.detected_sheets) or "Aucune")
    st.write("Feuilles exploitées :", ", ".join(result.processed_sheets) or "Aucune")
    st.write("Feuilles ignorées :", ", ".join(result.ignored_sheets) or "Aucune")

    rows_df = pd.DataFrame([row.as_dict() for row in result.rows])
    unknown_suppliers = _unique_values(rows_df, "supplier_raw", "supplier_id")
    unknown_products = _unique_values(rows_df, "product_raw", "product_id")

    st.subheader("Alertes")
    st.write("Fournisseurs inconnus :", ", ".join(unknown_suppliers) or "Aucun")
    st.write("Produits inconnus :", ", ".join(unknown_products) or "Aucun")

    st.subheader("Table de contrôle")
    if rows_df.empty:
        st.warning("Aucune ligne utile n'a été extraite.")
        return

    visible_columns = [col for col in CONTROL_COLUMNS if col in rows_df.columns]
    st.dataframe(rows_df[visible_columns], use_container_width=True, hide_index=True)

    st.subheader("Génération BDC")
    template_source = _resolve_template_source(template_file)
    if template_source is None:
        st.warning("Chargez le BDC type pour préparer les fichiers BDC.")
        return

    if st.button("Préparer les BDC Excel"):
        try:
            generation = BDCGenerationService(references).generate(
                template_source=template_source,
                rows=result.rows,
            )
        except BDCGenerationError as exc:
            st.error(str(exc))
            return
        except Exception:
            st.error("Les BDC n'ont pas pu être générés. Vérifiez le template Excel.")
            return

        if generation.warnings:
            for warning in generation.warnings:
                st.warning(warning)

        if not generation.files:
            st.info("Aucun fichier BDC généré.")
            return

        st.session_state["last_generation_files"] = generation.files
        st.session_state["last_prepared_emails"] = EmailComposerService(
            references
        ).compose_for_files(generation.files)

        _render_bdc_downloads(generation)

        if generation.skipped_rows:
            st.caption(
                "Les lignes en revue ou incomplètes ne sont pas injectées dans les BDC."
            )

    _render_email_preparation_section()


def _render_bdc_downloads(generation) -> None:
    st.write(f"Fichiers générés : {len(generation.files)}")
    summary_df = pd.DataFrame(
        [
            {
                "fichier": item.filename,
                "fournisseur": item.supplier_name,
                "lignes_injectees": item.injected_rows,
                "lignes_ignorees": item.skipped_rows,
            }
            for item in generation.files
        ]
    )
    st.dataframe(summary_df, use_container_width=True, hide_index=True)

    zip_bytes = BDCGenerationService.build_zip(generation.files)
    st.download_button(
        "Télécharger tous les BDC",
        data=zip_bytes,
        file_name="BDC_Mathieu_Lot2.zip",
        mime="application/zip",
    )

    st.write("Téléchargement individuel :")
    for generated_file in generation.files:
        st.download_button(
            label=f"Télécharger {generated_file.filename}",
            data=generated_file.content,
            file_name=generated_file.filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"download_{generated_file.filename}",
        )


def _render_email_preparation_section() -> None:
    payloads = st.session_state.get("last_prepared_emails")
    if not payloads:
        return

    st.subheader("Préparation des emails")
    st.caption(
        "Copiez les champs dans Outlook, puis téléchargez le BDC à joindre manuellement."
    )
    prepared_df = pd.DataFrame([_email_preview_row(payload) for payload in payloads])
    st.dataframe(prepared_df, use_container_width=True, hide_index=True)

    st.download_button(
        "Exporter le récapitulatif emails en CSV",
        data=prepared_df.to_csv(index=False).encode("utf-8-sig"),
        file_name="emails_prepares_mathieu.csv",
        mime="text/csv",
    )

    st.download_button(
        "Télécharger tous les BDC",
        data=_build_zip_from_payloads(payloads),
        file_name="BDC_Mathieu_Pieces_jointes.zip",
        mime="application/zip",
        key="email_download_all_bdc",
    )

    for payload in payloads:
        attachment_name = payload.attachment.filename if payload.attachment else "BDC manquant"
        with st.expander(f"{payload.supplier_name} - {attachment_name}"):
            if payload.needs_review:
                st.warning(payload.review_reason)

            st.write("Fichier BDC :", attachment_name)
            st.write("Pièce jointe :", _attachment_path_label(payload))

            col_left, col_right = st.columns([3, 1])
            with col_left:
                st.text_input("À", payload.to_recipients or "", key=f"to_{payload.supplier_id}")
            with col_right:
                _copy_button("Copier destinataires", payload.to_recipients or "", f"copy_to_{payload.supplier_id}")

            col_left, col_right = st.columns([3, 1])
            with col_left:
                st.text_input("Cc", payload.cc_recipients or "", key=f"cc_{payload.supplier_id}")
            with col_right:
                _copy_button("Copier cc", payload.cc_recipients or "", f"copy_cc_{payload.supplier_id}")

            col_left, col_right = st.columns([3, 1])
            with col_left:
                st.text_input("Objet", payload.subject or "", key=f"subject_{payload.supplier_id}")
            with col_right:
                _copy_button("Copier objet", payload.subject or "", f"copy_subject_{payload.supplier_id}")

            st.text_area(
                "Corps du mail",
                payload.body or "",
                key=f"body_{payload.supplier_id}",
                height=180,
            )
            _copy_button("Copier corps du mail", payload.body or "", f"copy_body_{payload.supplier_id}")

            if payload.attachment is not None:
                st.download_button(
                    f"Télécharger la pièce jointe BDC - {payload.attachment.filename}",
                    data=payload.attachment.content,
                    file_name=payload.attachment.filename,
                    mime=payload.attachment.mime_type,
                    key=f"email_bdc_download_{payload.supplier_id}",
                )


def _email_preview_row(payload) -> dict[str, str | None]:
    return {
        "fournisseur": payload.supplier_name,
        "email_a": payload.to_recipients,
        "email_cc": payload.cc_recipients,
        "objet": payload.subject,
        "corps": payload.body,
        "fichier_bdc": payload.attachment.filename if payload.attachment else None,
        "piece_jointe": _attachment_path_label(payload),
        "statut": "À compléter" if payload.needs_review else "Prêt à copier",
        "alerte": payload.review_reason,
    }


def _attachment_path_label(payload) -> str:
    if payload.attachment is None:
        return "BDC manquant"
    if payload.attachment.path:
        return payload.attachment.path
    return "Disponible au téléchargement dans Streamlit"


def _build_zip_from_payloads(payloads) -> bytes:
    stream = BytesIO()
    with ZipFile(stream, "w", ZIP_DEFLATED) as archive:
        for payload in payloads:
            if payload.attachment is not None:
                archive.writestr(payload.attachment.filename, payload.attachment.content)
    return stream.getvalue()


def _copy_button(label: str, text: str, key: str) -> None:
    safe_text = json.dumps(text or "")
    safe_label = json.dumps(label)
    components.html(
        f"""
        <button id="{key}" style="border:1px solid #ccc;border-radius:6px;padding:0.35rem 0.65rem;background:white;cursor:pointer;">
          {label}
        </button>
        <span id="{key}_status" style="font-family:sans-serif;font-size:12px;margin-left:8px;color:#555;"></span>
        <script>
        const btn = document.getElementById("{key}");
        const status = document.getElementById("{key}_status");
        btn.setAttribute("aria-label", {safe_label});
        btn.onclick = async () => {{
          try {{
            await navigator.clipboard.writeText({safe_text});
            status.innerText = "Copié";
          }} catch (err) {{
            status.innerText = "Copie impossible";
          }}
        }};
        </script>
        """,
        height=38,
    )


def _unique_values(df: pd.DataFrame, raw_col: str, id_col: str) -> list[str]:
    if df.empty or raw_col not in df or id_col not in df:
        return []
    values = df[df[id_col].isna()][raw_col].dropna().astype(str).str.strip()
    return sorted(value for value in values.unique() if value)


def _resolve_template_source(uploaded_file):
    return uploaded_file


if __name__ == "__main__":
    main()
