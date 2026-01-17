# from checks.base_check import BaseCheck, CheckResult


# class CoverStylesCheck(BaseCheck):
#     name = "Styly pro desky práce"
#     penalty = -5

#     def run(self, document, assignment=None):
#         if assignment is None:
#             return CheckResult(True, "Zadání nebylo předáno – kontrola přeskočena.", 0)
#         errors = []

#         required_styles = {
#             "desky-fakulta": assignment.styles.get("desky-fakulta"),
#             "desky-nazev-prace": assignment.styles.get("desky-nazev-prace"),
#             "desky-rok-a-jmeno": assignment.styles.get("desky-rok-a-jmeno"),
#         }

#         for style_name, expected in required_styles.items():
#             if expected is None:
#                 continue

#             actual = document.get_custom_style(style_name)
#             if actual is None:
#                 errors.append(f"Styl „{style_name}“ v dokumentu neexistuje.")
#                 continue

#             diffs = actual.diff(expected)
#             if diffs:
#                 errors.append(
#                     f"Styl „{style_name}“ neodpovídá zadání:\n"
#                     + "\n".join(f"  - {d}" for d in diffs)
#                 )

#         if errors:
#             return CheckResult(
#                 False,
#                 "Chyby ve stylech desek:\n" + "\n".join(errors),
#                 self.penalty,
#             )

#         return CheckResult(True, "Styly desek odpovídají zadání.", 0)

from checks.base_check import BaseCheck, CheckResult


class CoverStylesCheck(BaseCheck):
    name = "Styly pro desky práce"
    penalty = -5

    def run(self, document, assignment=None):
        if assignment is None:
            return CheckResult(True, "Zadání nebylo předáno – kontrola přeskočena.", 0)

        errors = []

        required = {
            "desky-fakulta": assignment.styles.get("desky-fakulta"),
            "desky-nazev-prace": assignment.styles.get("desky-nazev-prace"),
            "desky-rok-a-jmeno": assignment.styles.get("desky-rok-a-jmeno"),
        }

        for key, expected in required.items():
            if expected is None:
                continue

            actual = document.get_cover_style(key)

            if actual is None:
                errors.append(
                    f"Styl pro „{key}“ nebyl v dokumentu nalezen."
                )
                continue

            diffs = actual.diff(
                expected,
                doc_default_size=document.get_doc_default_font_size()
            )

            if diffs:
                errors.append(
                    f"Styl „{key}“ neodpovídá zadání:\n"
                    + "\n".join(f"  - {d}" for d in diffs)
                )

        if errors:
            return CheckResult(
                False,
                "Chyby ve stylech desek:\n" + "\n".join(errors),
                self.penalty,
            )

        return CheckResult(True, "Styly desek odpovídají zadání.", 0)