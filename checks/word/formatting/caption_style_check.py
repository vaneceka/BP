# from checks.base_check import BaseCheck, CheckResult

# class CaptionStyleCheck(BaseCheck):
#     name = "Styl Titulek není změněn."
#     penalty = -5

#     def run(self, document, assignment=None):
#         expected = assignment.styles.get("Caption")
#         if expected is None:
#             return CheckResult(True, "Zadání styl Titulek/Caption neřeší.", 0)

#         actual = document.get_style_by_any_name(["Caption", "Titulek"], default_alignment="start")
#         if actual is None:
#             return CheckResult(False, "Styl Caption/Titulek nebyl ve styles.xml nalezen.", self.penalty)

#         diffs = actual.diff(expected, doc_default_size=document.get_doc_default_font_size())
#         if diffs:
#             msg = "Styl Titulek (Caption) neodpovídá zadání:\n" + "\n".join(f"- {d}" for d in diffs)
#             return CheckResult(False, msg, self.penalty)

#         return CheckResult(True, "Styl Titulek (Caption) odpovídá zadání.", 0)

from checks.base_check import BaseCheck, CheckResult


class CaptionStyleCheck(BaseCheck):
    name = "Styl Titulek (Caption) není změněn."
    penalty = -5

    def run(self, document, assignment=None):
        expected = assignment.styles.get("Caption")
        if expected is None:
            return CheckResult(True, "Zadání styl Titulek/Caption neřeší.", 0)

        actual = document.get_style_by_any_name(
            ["Caption", "Titulek"],
            default_alignment="start"
        )

        if actual is None:
            return CheckResult(
                True,
                "Styl Caption nebyl v dokumentu upraven – používá se výchozí styl LibreOffice.",
                0
            )

        diffs = actual.diff(
            expected,
            doc_default_size=document.get_doc_default_font_size()
        )

        if diffs:
            msg = (
                "Styl Titulek (Caption) neodpovídá zadání:\n"
                + "\n".join(f"- {d}" for d in diffs)
            )
            return CheckResult(False, msg, self.penalty)

        return CheckResult(True, "Styl Titulek (Caption) odpovídá zadání.", 0)