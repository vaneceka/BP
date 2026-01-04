from checks.base_check import BaseCheck, CheckResult


class ChapterNumberingContinuityCheck(BaseCheck):
    name = "Číslování kapitol nepokračuje mezi oddíly"
    penalty = -5

    def run(self, document, assignment=None):
        if document.section_count() < 3:
            return CheckResult(
                True,
                "Dokument nemá třetí oddíl – nelze ověřit kontinuitu.",
                0,
            )

        h1_sec2 = document.first_heading_in_section(1, level=1)
        h1_sec3 = document.first_heading_in_section(2, level=1)

        if h1_sec2 is None or h1_sec3 is None:
            return CheckResult(
                True,
                "Alespoň jeden z oddílů neobsahuje kapitoly – kontinuita se nekontroluje.",
                0,
            )

        num2 = document.get_heading_num_id(h1_sec2)
        num3 = document.get_heading_num_id(h1_sec3)

        if num2 is None or num3 is None:
            return CheckResult(
                False,
                "Kapitoly nejsou číslované – nelze ověřit kontinuitu.",
                self.penalty,
            )

        if num2 != num3:
            return CheckResult(
                False,
                "Číslování kapitol ve třetím oddílu nepokračuje z předchozího oddílu.",
                self.penalty,
            )

        return CheckResult(
            True,
            "Číslování kapitol mezi oddíly správně pokračuje.",
            0,
        )