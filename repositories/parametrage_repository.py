from __future__ import annotations

from dataclasses import dataclass
from typing import BinaryIO

import pandas as pd

from domain.models import MailTemplateRef, PlantRef, ProductRef, ReferenceData, SupplierRef
from services.normalization_service import normalize_key


class ParametrageError(ValueError):
    """User-facing configuration error."""


@dataclass(frozen=True)
class SheetSpec:
    name: str
    required_columns: set[str]


class ParametrageRepository:
    BUYER_CODE = "MATHIEU"
    FORMAT_CODE = "MATHIEU_V1"

    REQUIRED_SHEETS = {
        "Acheteurs": SheetSpec(
            "Acheteurs", {"code_acheteur", "code_format_ph", "actif"}
        ),
        "Correspondance_Feuille_Usine": SheetSpec(
            "Correspondance_Feuille_Usine",
            {"code_format_ph", "nom_feuille", "code_usine", "nom_usine", "actif"},
        ),
        "Fournisseurs": SheetSpec(
            "Fournisseurs", {"code_fournisseur", "nom_fournisseur", "actif"}
        ),
        "Alias_Fournisseurs": SheetSpec(
            "Alias_Fournisseurs", {"alias", "code_fournisseur", "actif"}
        ),
        "Produits": SheetSpec(
            "Produits", {"code_produit", "nom_produit", "actif"}
        ),
        "Alias_Produits": SheetSpec(
            "Alias_Produits", {"alias", "code_produit", "actif"}
        ),
    }

    def load(self, source: BinaryIO) -> ReferenceData:
        self._ensure_file_like(source)
        self._rewind(source)
        sheets = pd.read_excel(source, sheet_name=None, header=None, engine="openpyxl")
        missing = [name for name in self.REQUIRED_SHEETS if name not in sheets]
        if missing:
            raise ParametrageError(
                "Feuille de paramétrage manquante : " + ", ".join(missing)
            )

        tables = {
            name: self._read_table(sheets[name], spec)
            for name, spec in self.REQUIRED_SHEETS.items()
        }
        mail_templates_by_supplier, default_mail_templates = self._load_mail_templates(sheets)

        buyers = self._active_rows(tables["Acheteurs"])
        buyer_rows = buyers[
            buyers["code_acheteur"].astype(str).str.upper() == self.BUYER_CODE
        ]
        if buyer_rows.empty:
            raise ParametrageError("Acheteur Mathieu introuvable ou inactif.")

        format_code = str(buyer_rows.iloc[0]["code_format_ph"]).strip()
        buyer_name = self._optional_text(buyer_rows.iloc[0], "nom_acheteur") or "Mathieu"
        buyer_email = self._optional_text(buyer_rows.iloc[0], "email_m365")
        if format_code != self.FORMAT_CODE:
            raise ParametrageError(
                f"Format PH attendu {self.FORMAT_CODE}, format trouvé {format_code}."
            )

        plants_by_sheet = self._load_plants(tables["Correspondance_Feuille_Usine"])
        suppliers_by_id = self._load_suppliers(tables["Fournisseurs"])
        supplier_aliases = self._load_aliases(
            tables["Alias_Fournisseurs"], "alias", "code_fournisseur", suppliers_by_id
        )
        products_by_id = self._load_products(tables["Produits"])
        product_aliases = self._load_aliases(
            tables["Alias_Produits"], "alias", "code_produit", products_by_id
        )

        return ReferenceData(
            buyer_code=self.BUYER_CODE,
            buyer_name=buyer_name,
            buyer_email=buyer_email,
            format_code=self.FORMAT_CODE,
            plants_by_sheet=plants_by_sheet,
            suppliers_by_id=suppliers_by_id,
            supplier_aliases=supplier_aliases,
            products_by_id=products_by_id,
            product_aliases=product_aliases,
            mail_templates_by_supplier=mail_templates_by_supplier,
            default_mail_templates_by_language=default_mail_templates,
        )

    def _read_table(self, raw_df: pd.DataFrame, spec: SheetSpec) -> pd.DataFrame:
        header_idx = self._find_header_row(raw_df, spec.required_columns)
        if header_idx is None:
            required = ", ".join(sorted(spec.required_columns))
            raise ParametrageError(
                f"Colonnes indispensables introuvables dans {spec.name} : {required}"
            )

        columns = [self._clean_column_name(value) for value in raw_df.iloc[header_idx]]
        df = raw_df.iloc[header_idx + 1 :].copy()
        df.columns = columns
        df = df.loc[:, [col for col in df.columns if col]]
        df = df.dropna(how="all")

        missing = sorted(spec.required_columns - set(df.columns))
        if missing:
            raise ParametrageError(
                f"Colonnes indispensables manquantes dans {spec.name} : "
                + ", ".join(missing)
            )
        return df

    def _find_header_row(self, df: pd.DataFrame, required_columns: set[str]) -> int | None:
        for idx, row in df.iterrows():
            names = {self._clean_column_name(value) for value in row}
            if required_columns.issubset(names):
                return int(idx)
        return None

    def _clean_column_name(self, value: object) -> str:
        if value is None or pd.isna(value):
            return ""
        return str(value).strip()

    def _active_rows(self, df: pd.DataFrame) -> pd.DataFrame:
        return df[df["actif"].map(self._is_active)].copy()

    def _is_active(self, value: object) -> bool:
        return normalize_key(value) in {"oui", "yes", "true", "1", "actif"}

    def _load_plants(self, df: pd.DataFrame) -> dict[str, PlantRef]:
        active = self._active_rows(df)
        active = active[active["code_format_ph"].astype(str).str.strip() == self.FORMAT_CODE]
        plants: dict[str, PlantRef] = {}
        for _, row in active.iterrows():
            sheet_name = str(row["nom_feuille"]).strip()
            plants[sheet_name] = PlantRef(
                sheet_name=sheet_name,
                plant_code=str(row["code_usine"]).strip(),
                plant_name=str(row["nom_usine"]).strip(),
            )
        if not plants:
            raise ParametrageError("Aucune feuille usine active trouvée pour Mathieu.")
        return plants

    def _load_suppliers(self, df: pd.DataFrame) -> dict[str, SupplierRef]:
        suppliers: dict[str, SupplierRef] = {}
        for _, row in self._active_rows(df).iterrows():
            supplier_id = str(row["code_fournisseur"]).strip()
            suppliers[supplier_id] = SupplierRef(
                supplier_id=supplier_id,
                supplier_name=str(row["nom_fournisseur"]).strip(),
                generate_bdc=self._optional_bool(row, "generer_bdc", True),
                ignore_if_detected=self._optional_bool(row, "ignorer_si_detecte", False),
                bdc_type=self._optional_text(row, "type_bdc"),
                delay_codes_by_plant=self._load_supplier_delay_codes(row),
                quantity_mode=self._optional_text(row, "mode_quantite_bdc"),
                model_code=self._optional_text(row, "code_modele"),
                email_to=self._optional_text(row, "email_a"),
                email_cc=self._optional_text(row, "email_cc"),
                language=self._optional_text(row, "langue_mail"),
            )
        return suppliers

    def _load_products(self, df: pd.DataFrame) -> dict[str, ProductRef]:
        products: dict[str, ProductRef] = {}
        for _, row in self._active_rows(df).iterrows():
            product_id = str(row["code_produit"]).strip()
            products[product_id] = ProductRef(
                product_id=product_id,
                product_name=str(row["nom_produit"]).strip(),
            )
        return products

    def _load_aliases(
        self,
        df: pd.DataFrame,
        alias_col: str,
        code_col: str,
        known_refs: dict[str, object],
    ) -> dict[str, str]:
        aliases: dict[str, str] = {}
        for _, row in self._active_rows(df).iterrows():
            alias = normalize_key(row[alias_col])
            code = str(row[code_col]).strip()
            if alias and code in known_refs:
                aliases[alias] = code
        return aliases

    def _ensure_file_like(self, source: object) -> None:
        if isinstance(source, (str, bytes)) or not hasattr(source, "read"):
            raise ParametrageError(
                "Le fichier de paramétrage doit être chargé depuis l'interface."
            )

    def _rewind(self, source: BinaryIO) -> None:
        if hasattr(source, "seek"):
            source.seek(0)

    def _optional_text(self, row: pd.Series, column: str) -> str | None:
        if column not in row.index or pd.isna(row[column]):
            return None
        value = str(row[column]).strip()
        return value or None

    def _optional_bool(self, row: pd.Series, column: str, default: bool) -> bool:
        if column not in row.index or pd.isna(row[column]):
            return default
        key = normalize_key(row[column])
        if key in {"oui", "yes", "true", "1"}:
            return True
        if key in {"non", "no", "false", "0"}:
            return False
        return default

    def _load_supplier_delay_codes(self, row: pd.Series) -> dict[str, str]:
        plant_columns = {
            "KB": ["code_delai_KB"],
            "C9": ["code_delai_C9"],
            "GS": ["code_delai_GS", "code_delai_GNS"],
            "GNS": ["code_delai_GNS", "code_delai_GS"],
            "VNC": ["code_delai_VNC"],
            "PERPI": ["code_delai_PERPI"],
            "HTG": ["code_delai_HTG"],
        }
        delays: dict[str, str] = {}
        for plant_code, columns in plant_columns.items():
            for column in columns:
                value = self._optional_text(row, column)
                if value:
                    delays[plant_code] = value
                    break
        return delays

    def _load_mail_templates(
        self, sheets: dict[str, pd.DataFrame]
    ) -> tuple[dict[str, MailTemplateRef], dict[str, MailTemplateRef]]:
        if "Modeles_Mails" not in sheets:
            return {}, {}

        spec = SheetSpec(
            "Modeles_Mails",
            {"code_modele_mail", "code_fournisseur", "objet_modele", "corps_modele", "actif"},
        )
        table = self._read_table(sheets["Modeles_Mails"], spec)
        by_supplier: dict[str, MailTemplateRef] = {}
        by_language: dict[str, MailTemplateRef] = {}
        for _, row in self._active_rows(table).iterrows():
            template = MailTemplateRef(
                template_code=str(row["code_modele_mail"]).strip(),
                supplier_id=str(row["code_fournisseur"]).strip(),
                subject_template=str(row["objet_modele"]).strip(),
                body_template=str(row["corps_modele"]),
            )
            if template.supplier_id == "*":
                by_language[self._language_from_template_code(template.template_code)] = template
            else:
                by_supplier[template.supplier_id] = template
        return by_supplier, by_language

    def _language_from_template_code(self, template_code: str) -> str:
        suffix = template_code.rsplit("_", 1)[-1].strip().upper()
        return suffix or "FR"
