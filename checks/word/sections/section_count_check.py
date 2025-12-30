from ...base_check import BaseCheck, CheckResult


class SectionCountCheck(BaseCheck):
    name = "Počet oddílů"
    penalty = -10

    def run(self, document, assignment=None):
        count = document.section_count()

        if count <= 1:
            return CheckResult(
                False,
                "Oddíly zcela chybí (nalezen pouze 1 oddíl).",
                -100,
                fatal=True
            )

        if count < 3:
            return CheckResult(
                False,
                f"Nalezeno {count} oddílů, požadovány 3.",
                self.penalty
            )

        return CheckResult(True, f"Počet oddílů OK ({count})", 0)