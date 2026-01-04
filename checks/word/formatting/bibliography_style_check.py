from checks.base_check import BaseCheck, CheckResult

class BibliographyStyleCheck(BaseCheck):
    name = "Styl Bibliografie není změněn."
    penalty = -5

    def run(self, document, assignment=None):
        expected = assignment.styles.get("Bibliography")
        if expected is None:
            return CheckResult(True, "Zadání styl Bibliografie neřeší.", 0)

        actual = document.get_style_by_any_name(["Bibliography", "Bibliografie"], default_alignment="start")
        if actual is None:
            return CheckResult(False, "Styl Bibliografie nebyl ve styles.xml nalezen.", self.penalty)

        diffs = actual.diff(expected, doc_default_size=document.get_doc_default_font_size())
        if diffs:
            msg = "Styl Bibliografie neodpovídá zadání:\n" + "\n".join(f"- {d}" for d in diffs)
            return CheckResult(False, msg, self.penalty)

        return CheckResult(True, "Styl Bibliografie odpovídá zadání.", 0)