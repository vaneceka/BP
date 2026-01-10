# from checks.base_check import BaseCheck, CheckResult

# class ThreeDChartCheck(BaseCheck):
#     name = "Použit 3D graf"
#     penalty = -2
#     SHEET = "data"

#     def run(self, document, assignment=None):

#         if self.SHEET not in document.wb.sheetnames:
#             return CheckResult(
#                 False,
#                 f'Chybí list "{self.SHEET}".',
#                 self.penalty,
#                 fatal=True
#             )

#         ws = document.wb[self.SHEET]

#         charts = getattr(ws, "_charts", [])

#         if not charts:
#             return CheckResult(True, "Graf nebyl nalezen.", 0)

#         bad = []

#         for chart in charts:
#             chart_class = chart.tagname

#             if "3D" in chart_class:
#                 bad.append(chart_class)

#         if bad:
#             return CheckResult(
#                 False,
#                 "Použit 3D graf:\n– " + "\n– ".join(bad),
#                 self.penalty,
#                 fatal=True
#             )

#         return CheckResult(
#             True,
#             "Použit 2D graf.",
#             0
#         )

from checks.base_check import BaseCheck, CheckResult


class ThreeDChartCheck(BaseCheck):
    name = "Použit 3D graf"
    penalty = -2
    SHEET = "data"

    def run(self, document, assignment=None):

        # pokud graf vůbec není → OK
        if not document.has_chart(self.SHEET):
            return CheckResult(True, "Graf nebyl nalezen.", 0)

        if document.has_3d_chart(self.SHEET):
            return CheckResult(
                False,
                "Použit 3D graf.",
                self.penalty,
                fatal=True,
            )

        return CheckResult(
            True,
            "Použit 2D graf.",
            0,
        )