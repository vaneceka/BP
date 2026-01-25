from checks.base_check import BaseCheck, CheckResult


# class UnnumberedSpecialHeadingsCheck(BaseCheck):
#     name = "Styly speciálních kapitol nesmí být číslované"
#     penalty = -1

#     SPECIAL_STYLES = [
#         "obsah",
#         "Nadpisobsahu",
#         "content heading",
#         "Seznamobrzk",
#         "seznam obrázků",
#         "seznam tabulek",
#         "bibliografie",
#         "Vrazncitt"
#     ]

#     def run(self, document, assignment=None):
#         errors = []

#         for style_name in self.SPECIAL_STYLES:
#             style = document.get_style_by_any_name([style_name])
#             if style and style.isNumbered:
#                 errors.append(style.name)

#         if errors:
#             return CheckResult(
#                 False,
#                 "Následující styly mají zapnuté číslování:\n"
#                 + "\n".join(f"- {s}" for s in errors),
#                 self.penalty,
#             )

#         return CheckResult(True, "Styly speciálních kapitol nejsou číslované.", 0)

class UnnumberedSpecialHeadingsCheck(BaseCheck):
    name = "Speciální kapitoly nesmí být číslované"
    penalty = -1

    SPECIAL_TITLES = {
        "obsah",
        "bibliografie",
        "seznam obrázků",
        "seznam tabulek",
    }

    def run(self, document, assignment=None):
        errors = set()

        for text, level in document.iter_headings():
            if text.strip().lower() in self.SPECIAL_TITLES:
                if document.heading_level_is_numbered(level):
                    errors.add(f"{text} (úroveň {level})")

        for p in document.iter_paragraphs():
            style = (document.paragraph_style_name(p) or "").lower()
            if not style.endswith("1"):
                continue  

            text = (document.paragraph_text(p) or "").strip()
            if text.lower() in self.SPECIAL_TITLES:
                if document.heading_level_is_numbered(1):
                    errors.add(f"{text} (úroveň 1)")

        if errors:
            return CheckResult(
                False,
                "Následující speciální kapitoly jsou číslované:\n"
                + "\n".join(f"- {e}" for e in errors),
                self.penalty,
            )

        return CheckResult(
            True,
            "Speciální kapitoly nejsou číslované.",
            0
        )