# from checks.base_check import BaseCheck, CheckResult


# class ConditionalFormattingExistsCheck(BaseCheck):
#     name = "Chybí podmíněné formátování"
#     penalty = -5
#     SHEET = "data"

#     def run(self, document, assignment=None):

#         if self.SHEET not in document.wb.sheetnames:
#             return CheckResult(False, f'Chybí list "{self.SHEET}".', self.penalty, fatal=True)

#         ws = document.wb[self.SHEET]

#         if not ws.conditional_formatting._cf_rules:
#             return CheckResult(
#                 False,
#                 "V sešitu není nastaveno žádné podmíněné formátování.",
#                 self.penalty,
#                 fatal=True
#             )

#         return CheckResult(
#             True,
#             "Podmíněné formátování je v sešitu přítomno.",
#             0
#         )

from checks.base_check import BaseCheck, CheckResult


class ConditionalFormattingExistsCheck(BaseCheck):
    name = "Chybí podmíněné formátování"
    penalty = -5
    SHEET = "data"

    def run(self, document, assignment=None):

        if self.SHEET not in document.sheet_names():
            return CheckResult(
                False,
                f'Chybí list "{self.SHEET}".',
                self.penalty,
                fatal=True,
            )

        if not document.has_conditional_formatting(self.SHEET):
            return CheckResult(
                False,
                "V sešitu není nastaveno žádné podmíněné formátování.",
                self.penalty,
                fatal=True,
            )

        return CheckResult(
            True,
            "Podmíněné formátování je v sešitu přítomno.",
            0,
        )