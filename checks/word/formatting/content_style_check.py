from checks.word.base_check import BaseCheck, CheckResult


class ContentHeadingStyleCheck(BaseCheck):
    name = "Styl Nadpis obsahu není změněn dle zadání"
    penalty = -5

    def run(self, document, assignment=None):
        expected = assignment.styles.get("Content Heading")
        if expected is None:
            return CheckResult(True, "Zadání styl Content Heading neřeší.", 0)

        actual = document.get_style_by_any_name(
            ["Content Heading", "TOC Heading", "Nadpisobsahu", "Nadpis obsahu"],
            default_alignment="start",
        )
        if actual is None:
            return CheckResult(
                False,
                "Styl Content Heading nebyl ve styles.xml nalezen.",
                self.penalty
            )

        diffs = actual.diff(expected, strict=True)

        if diffs:
            message = (
                "Styl Content Heading neodpovídá zadání:\n"
                + "\n".join(f"- {d}" for d in diffs)
            )
            return CheckResult(False, message, self.penalty)

        return CheckResult(True, "Styl Content Heading odpovídá zadání.", 0)