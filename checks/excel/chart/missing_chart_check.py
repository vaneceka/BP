from checks.base_check import BaseCheck, CheckResult


# class MissingChartCheck(BaseCheck):
#     name = "Požadovaný graf zcela chybí"
#     penalty = -100
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

#         if not ws._charts:
#             return CheckResult(
#                 False,
#                 "Požadovaný graf zcela chybí.",
#                 self.penalty,
#                 fatal=True
#             )

#         return CheckResult(
#             True,
#             "Požadovaný graf je v sešitu přítomen.",
#             0
#         )

from checks.base_check import BaseCheck, CheckResult


class MissingChartCheck(BaseCheck):
    name = "Požadovaný graf zcela chybí"
    penalty = -100
    SHEET = "data"

    def run(self, document, assignment=None):

        # pokud zadání graf nevyžaduje → OK
        if assignment is None or not assignment.chart:
            return CheckResult(True, "Zadání graf nevyžaduje.", 0)

        if not document.has_chart(self.SHEET):
            return CheckResult(
                False,
                "Požadovaný graf zcela chybí.",
                self.penalty,
                fatal=True,
            )

        return CheckResult(
            True,
            "Požadovaný graf je v sešitu přítomen.",
            0,
        )