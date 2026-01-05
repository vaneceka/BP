from dataclasses import dataclass
from typing import Dict, List

@dataclass
class StyleSpec:
    TAB_TOLERANCE = 5 
    SPACE_TOLERANCE = 20    
    INDENT_TOLERANCE = 20
    name: str
    font: str | None = None
    size: int | None = None
    bold: bool | None = None
    italic: bool | None = None
    underline: bool | None = None
    allCaps: bool | None = None
    color: str | None = None
    alignment: str | None = None
    lineHeight: float | None = None
    pageBreakBefore: bool | None = None
    isNumbered: bool | None = None
    numLevel: int | None = None
    basedOn: str | None = None
    spaceBefore: int | None = None 
    indentLeft: int | None = None
    indentRight: int | None = None
    indentFirstLine: int | None = None
    indentHanging: int | None = None
    tabs: list[tuple[str, int]] | None = None 
    
    def _int_close(self, a: int | None, b: int | None, tol: int) -> bool:
        if a is None and b is None:
            return True
        if a is None or b is None:
            return False
        return abs(a - b) <= tol


    def _tabs_close(
        self,
        actual: list[tuple[str, int]] | None,
        expected: list[tuple[str, int]] | None,
        tol: int = TAB_TOLERANCE,
    ) -> bool:
        if actual is None and expected is None:
            return True
        if actual is None or expected is None:
            return False
        if len(actual) != len(expected):
            return False

        for (a_align, a_pos), (e_align, e_pos) in zip(actual, expected):
            if a_align != e_align:
                return False
            if abs(a_pos - e_pos) > tol:
                return False

        return True

    def diff(
        self,
        expected: "StyleSpec",
        *,
        doc_default_size: int | None = None,
        ignore_fields: set[str] | None = None,
    ) -> list[str]:
        if expected is None:
            return []

        ignore = {"name"} if ignore_fields is None else ignore_fields
        diffs: list[str] = []

        for field, expected_value in vars(expected).items():
            if field in ignore:
                continue
            if expected_value is None:
                continue

            actual_value = getattr(self, field, None)

            if field == "spaceBefore":
                if not self._int_close(actual_value, expected_value, self.SPACE_TOLERANCE):
                    diffs.append(
                        f"spaceBefore: očekáváno {expected_value}, nalezeno {actual_value}"
                    )
                continue

            if field.startswith("indent"):
                if not self._int_close(actual_value, expected_value, self.SPACE_TOLERANCE):
                    diffs.append(
                        f"{field}: očekáváno {expected_value}, nalezeno {actual_value}"
                    )
                continue

            if field == "tabs":
                if not self._tabs_close(actual_value, expected_value):
                    diffs.append(
                        f"tabs: očekáváno {expected_value}, nalezeno {actual_value}"
                    )
                continue

            if field == "size":
                if actual_value is None:
                    if doc_default_size != expected_value:
                        diffs.append(
                            f"size: očekáváno {expected_value}, default dokumentu je {doc_default_size}"
                        )
                elif actual_value != expected_value:
                    diffs.append(f"size: očekáváno {expected_value}, nalezeno {actual_value}")
                continue

            if field == "alignment":
                norm = {
                    "left": "start",
                    "right": "end",
                }
                act = norm.get(actual_value, actual_value)
                exp = norm.get(expected_value, expected_value)

                if act != exp:
                    diffs.append(
                        f"alignment: očekáváno {expected_value}, nalezeno {actual_value}"
                    )
                continue

            if actual_value != expected_value:
                diffs.append(f"{field}: očekáváno {expected_value}, nalezeno {actual_value}")

        return diffs

    def matches(self, expected: "StyleSpec", *, doc_default_size: int | None = None, strict: bool = False) -> bool:
        return len(self.diff(expected, doc_default_size=doc_default_size)) == 0
    

@dataclass
class WordAssignment:
    styles: Dict[str, StyleSpec]
    headlines: List[dict]
    objects: List[dict]
    bibliography: List[dict]