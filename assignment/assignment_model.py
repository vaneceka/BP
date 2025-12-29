from dataclasses import dataclass

@dataclass
class StyleSpec:
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

    indentLeft: int | None = None
    indentRight: int | None = None
    indentFirstLine: int | None = None
    indentHanging: int | None = None
    tabs: list[tuple[str, int]] | None = None 
    

    def diff(
        self,
        expected: "StyleSpec",
        *,
        doc_default_size: int | None = None,
        ignore_fields: set[str] | None = None,
        strict: bool = False,
    ) -> list[str]:
        if expected is None:
            return []

        ignore = {"name"} if ignore_fields is None else ignore_fields
        diffs: list[str] = []

        # 1) "musí odpovídat" – jen pole, která jsou v expected vyplněná
        for field, expected_value in vars(expected).items():
            if field in ignore:
                continue
            if expected_value is None:
                continue

            actual_value = getattr(self, field, None)

            if field == "size":
                if actual_value is None:
                    if doc_default_size != expected_value:
                        diffs.append(
                            f"size: očekáváno {expected_value}, default dokumentu je {doc_default_size}"
                        )
                elif actual_value != expected_value:
                    diffs.append(f"size: očekáváno {expected_value}, nalezeno {actual_value}")
                continue

            if actual_value != expected_value:
                diffs.append(f"{field}: očekáváno {expected_value}, nalezeno {actual_value}")

        # 2) STRICT: nic navíc nesmí být zapnuto (jen když expected to nepožaduje)
        if strict:
            # jen bool pole (u čísel/fontů strict typicky nedává smysl bez dalších pravidel)
            bool_fields = ["bold", "italic","underline", "allCaps", "pageBreakBefore"]
            for field in bool_fields:
                if field in ignore:
                    continue
                exp = getattr(expected, field, None)
                act = getattr(self, field, None)

                # pokud expected nic neříká, bereme "nemá být zapnuto"
                if exp is None and act is True:
                    diffs.append(f"{field}: nemá být zapnuto, ale je True")
            
            indent_fields = [
                "indentLeft",
                "indentRight",
                "indentFirstLine",
                "indentHanging",
            ]
                    
            for field in indent_fields:
                exp = getattr(expected, field, None)
                act = getattr(self, field, None)
                if exp is None and act is not None:
                    diffs.append(f"{field}: nemá být nastaveno, ale je {act}")
            
            # Zadaní řeší číslování
            if expected.numLevel is not None:
                if not self.isNumbered:
                    diffs.append("isNumbered: má být číslovaný, ale není")
            else:
                # Zadaní neřeší číslování
                if self.isNumbered:
                    diffs.append("isNumbered: nemá být číslovaný, ale je")
        return diffs

    def matches(self, expected: "StyleSpec", *, doc_default_size: int | None = None, strict: bool = False) -> bool:
        return len(self.diff(expected, doc_default_size=doc_default_size, strict=strict)) == 0