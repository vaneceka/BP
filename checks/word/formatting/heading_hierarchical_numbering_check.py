from checks.base_check import BaseCheck, CheckResult

class HeadingHierarchicalNumberingCheck(BaseCheck):
    name = (
        "Nadpisy nemají aktivováno automatické "
        "hierarchické číslování (např. 1.1.1)"
    )
    penalty = -5

    def run(self, document, assignment=None):
        broken = []

        for level in (1, 2, 3):
            expected = assignment.styles.get(f"Heading {level}")
            actual = document.get_heading_style(level)

            if expected is None or actual is None:
                continue

            if not actual.isNumbered:
                broken.append(
                    f"Nadpis {level}: není číslovaný"
                )
                continue

            if actual.numLevel != level - 1:
                broken.append(
                    f"Nadpis {level}: špatná úroveň "
                    f"(očekáváno {level-1}, nalezeno {actual.numLevel})"
                )

        if broken:
            message = (
                "Nadpisy nemají správně nastavené "
                "automatické hierarchické číslování:\n"
                + "\n".join(f"- {b}" for b in broken)
            )
            return CheckResult(False, message, self.penalty)

        return CheckResult(
            True,
            "Automatické hierarchické číslování nadpisů je nastaveno správně.",
            0,
        )