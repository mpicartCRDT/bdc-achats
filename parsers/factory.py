from __future__ import annotations

from domain.models import ReferenceData
from parsers.mathieu_parser import MathieuParser


def get_parser(format_code: str, references: ReferenceData):
    if format_code == "MATHIEU_V1":
        return MathieuParser(references)
    raise ValueError(f"Format PH non pris en charge dans ce lot : {format_code}")
