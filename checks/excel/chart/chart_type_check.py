from checks.base_check import BaseCheck, CheckResult


class ChartTypeCheck(BaseCheck):
    name = "Nevhodný typ grafu"
    penalty = -5
    SHEET = "data"

   

    def run(self, document, assignment=None):
        if assignment is None:
            return CheckResult(False, "Assigment chybí - přeskakuji", self.penalty)
        
        if self.SHEET not in document.wb.sheetnames:
            return CheckResult(
                False,
                f'Chybí list "{self.SHEET}".',
                self.penalty,
                fatal=True
            )

        ws = document.wb[self.SHEET]

        if not ws._charts:
            return CheckResult(
                False,
                "Graf neexistuje",               
                self.penalty)                           
        
        expected_type = assignment.chart.get("type")
        
        actual_chart = ws._charts[0]
        actual_type = actual_chart.tagname

        if actual_type.lower() != expected_type.lower():
            return CheckResult(
                False,
                f"Nevhodný typ grafu:\n"
                f"– očekáváno: {expected_type}\n"
                f"– nalezeno: {actual_type}",
                self.penalty,
                fatal=True
            )

        return CheckResult(
            True,
            "Typ grafu odpovídá zadání.",
            0
        )

