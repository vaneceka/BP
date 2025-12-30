from checks.base_check import BaseCheck, CheckResult

class DocumentStructureCheck(BaseCheck):
    name = "Chybná struktura dokumentu"
    penalty = -5  # násobí se

    def run(self, document, assignment=None):
        headings = document.iter_headings()
        for h in headings:
            print(h)
        if not headings:
            return CheckResult(True, "Dokument neobsahuje nadpisy.", 0)

        errors = []
        last_level = None

        for text, level in headings:
            # dokument nezačíná Nadpisem 1
            if last_level is None and level != 1:
                errors.append(
                    f"Nadpis „{text}“ je úrovně {level}, dokument musí začínat Nadpisem 1."
                )

            # přeskok úrovně (např. H1 -> H3)
            if last_level is not None and level > last_level + 1:
                errors.append(
                    f"Nadpis „{text}“ (úroveň {level}) přeskakuje úroveň "
                    f"(předchozí byla {last_level})."
                )

            last_level = level

        if errors:
            return CheckResult(
                False,
                "Chybná hierarchie nadpisů:\n" + "\n".join(errors),
                self.penalty * len(errors),
            )

        return CheckResult(True, "Hierarchie nadpisů je správná.", 0)