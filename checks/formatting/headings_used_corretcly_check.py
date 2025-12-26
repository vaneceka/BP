from collections import Counter
from checks.base_check import BaseCheck, CheckResult

class HeadingsUsedCorrectlyCheck(BaseCheck):
    name = "Nadpisy nejsou použity správně dle seznamu"
    points_per_error = -5

    def run(self, document, assignment=None):
        expected_list = assignment.headlines
        expected = Counter(
            (h["text"].strip(), int(h["level"])) for h in expected_list
        )

        actual_list = document.iter_headings()
        actual = Counter(actual_list)

        missing = list((expected - actual).elements())
        extra = list((actual - expected).elements())

        error_count = len(missing) + len(extra)

        if error_count == 0:
            return CheckResult(
                True,
                "Všechny nadpisy odpovídají textům i úrovním.",
                0,
            )

        penalty = self.points_per_error * error_count

        lines = [
            f"Nalezeno {error_count} chyb v použití nadpisů "
            f"({abs(self.points_per_error)} bodů za každou):"
        ]

        if missing:
            lines.append("\nChybí nadpisy:")
            lines += [f"- {t} (H{lvl})" for (t, lvl) in missing]

        if extra:
            lines.append("\nNavíc / špatně zařazené nadpisy:")
            lines += [f"- {t} (H{lvl})" for (t, lvl) in extra]

        lines.append(f"\nCelková penalizace: {penalty} bodů")

        return CheckResult(False, "\n".join(lines), penalty)