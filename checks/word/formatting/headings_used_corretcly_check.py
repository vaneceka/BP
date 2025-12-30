from checks.word.base_check import BaseCheck, CheckResult

class HeadingsUsedCorrectlyCheck(BaseCheck):
    name = "Nadpisy nejsou použity správně dle seznamu"
    penalty = -5

    def run(self, document, assignment=None):
        expected = [
            (h["text"].strip(), int(h["level"]))
            for h in assignment.headlines
        ]

        actual = document.iter_headings()

        def count(items):
            d = {}
            for i in items:
                d[i] = d.get(i, 0) + 1
            return d

        expected_counts = count(expected)
        actual_counts = count(actual)

        missing = []
        for k, v in expected_counts.items():
            missing.extend([k] * max(0, v - actual_counts.get(k, 0)))

        extra = []
        # k -> nadpis, v -> pocet
        for k, v in actual_counts.items():
            extra.extend([k] * max(0, v - expected_counts.get(k, 0)))

        if not missing and not extra:
            return CheckResult(True, "Všechny nadpisy odpovídají.", 0)

        lines = []

        if missing:
            lines.append("Chybí nadpisy:")
            lines += [f"- {t} (H{lvl})" for t, lvl in missing]

        if extra:
            lines.append("Navíc / špatné nadpisy:")
            lines += [f"- {t} (H{lvl})" for t, lvl in extra]

        penalty = self.penalty * (len(missing) + len(extra))

        return CheckResult(False, "\n".join(lines), penalty)