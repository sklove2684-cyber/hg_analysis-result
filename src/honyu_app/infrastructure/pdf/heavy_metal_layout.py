from __future__ import annotations

from dataclasses import dataclass
import re


ELEMENT_SYMBOLS = frozenset(
    "H He Li Be B C N O F Ne Na Mg Al Si P S Cl Ar K Ca Sc Ti V Cr Mn Fe Co Ni "
    "Cu Zn Ga Ge As Se Br Kr Rb Sr Y Zr Nb Mo Tc Ru Rh Pd Ag Cd In Sn Sb Te I "
    "Xe Cs Ba La Ce Pr Nd Pm Sm Eu Gd Tb Dy Ho Er Tm Yb Lu Hf Ta W Re Os Ir Pt "
    "Au Hg Tl Pb Bi Po At Rn Fr Ra Ac Th Pa U Np Pu Am Cm Bk Cf Es Fm Md No Lr "
    "Rf Db Sg Bh Hs Mt Ds Rg Cn Nh Fl Mc Lv Ts Og".split()
)


@dataclass(frozen=True, slots=True)
class HeavyMetalTableLayout:
    header_row: int
    sample_name_column: int
    element_columns: tuple[tuple[str, int], ...]


def normalize_header(value: object | None) -> str:
    return " ".join(str(value or "").split())


def element_symbol_from_header(value: object | None) -> str | None:
    text = normalize_header(value)
    quant = re.search(r"\bQuant\b", text, re.IGNORECASE)
    average = re.search(r"\bAverage\b", text, re.IGNORECASE)
    if quant is None or average is None:
        return None
    candidates = re.findall(r"\b[A-Z][a-z]?\b", text[:quant.start()])
    symbols = [candidate for candidate in candidates if candidate in ELEMENT_SYMBOLS]
    return symbols[-1] if len(set(symbols)) == 1 else None


def find_heavy_metal_table_layout(
    table: list[list[object | None]], *, min_elements: int = 1
) -> HeavyMetalTableLayout | None:
    if not table:
        return None
    column_count = max((len(row) for row in table), default=0)
    for header_row in range(min(7, len(table))):
        headers = []
        for column in range(column_count):
            headers.append(" ".join(
                normalize_header(row[column])
                for row in table[:header_row + 1]
                if column < len(row) and normalize_header(row[column])
            ))
        sample_columns = [
            index for index, header in enumerate(headers)
            if re.search(r"\bSample\s*Name\b", header, re.IGNORECASE)
        ]
        element_columns = tuple(
            (symbol, index)
            for index, header in enumerate(headers)
            if (symbol := element_symbol_from_header(header)) is not None
        )
        if len(sample_columns) == 1 and len(element_columns) >= min_elements:
            symbols = [symbol for symbol, _ in element_columns]
            if len(symbols) != len(set(symbols)):
                raise ValueError("duplicated element header")
            return HeavyMetalTableLayout(
                header_row, sample_columns[0], element_columns
            )
    return None
