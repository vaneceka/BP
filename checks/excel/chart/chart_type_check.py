from checks.base_check import BaseCheck, CheckResult


# class ChartTypeCheck(BaseCheck):
#     name = "Nevhodný typ grafu"
#     penalty = -5
#     SHEET = "data"

   

#     def run(self, document, assignment=None):
#         if assignment is None:
#             return CheckResult(False, "Assigment chybí - přeskakuji", self.penalty)
        
#         if self.SHEET not in document.wb.sheetnames:
#             return CheckResult(
#                 False,
#                 f'Chybí list "{self.SHEET}".',
#                 self.penalty,
#                 fatal=True
#             )

#         ws = document.wb[self.SHEET]

#         if not ws._charts:
#             return CheckResult(
#                 False,
#                 "Graf neexistuje",               
#                 self.penalty)                           
        
#         expected_type = assignment.chart.get("type")
        
#         actual_chart = ws._charts[0]
#         actual_type = actual_chart.tagname

#         if actual_type.lower() != expected_type.lower():
#             return CheckResult(
#                 False,
#                 f"Nevhodný typ grafu:\n"
#                 f"– očekáváno: {expected_type}\n"
#                 f"– nalezeno: {actual_type}",
#                 self.penalty,
#                 fatal=True
#             )

#         return CheckResult(
#             True,
#             "Typ grafu odpovídá zadání.",
#             0
#         )

from checks.base_check import BaseCheck, CheckResult


class ChartTypeCheck(BaseCheck):
    name = "Nevhodný typ grafu"
    penalty = -5
    SHEET = "data"

    def normalize_chart_type(self, t: str | None) -> str | None:
        if not t:
            return None

        t = t.lower()

        mapping = {
            # Excel
            "barchart": "bar",
            "linechart": "line",
            "piechart": "pie",

            # ODS
            "bar": "bar",
            "line": "line",
            "pie": "pie",
        }

        return mapping.get(t, t)

    def run(self, document, assignment=None):
        if assignment is None or not assignment.chart:
            return CheckResult(True, "Zadání graf nevyžaduje.", 0)

        expected_raw = assignment.chart.get("type")
        expected = self.normalize_chart_type(expected_raw)

        actual_raw = document.chart_type(self.SHEET)
        actual = self.normalize_chart_type(actual_raw)

        if actual is None:
            return CheckResult(
                False,
                "Graf neexistuje.",
                self.penalty,
                fatal=True,
            )

        if expected != actual:
            return CheckResult(
                False,
                "Nevhodný typ grafu:\n"
                f"– očekáváno: {expected_raw}\n"
                f"– nalezeno: {actual_raw}",
                self.penalty,
                fatal=True,
            )

        return CheckResult(True, "Typ grafu odpovídá zadání.", 0)