from checks.base_check import BaseCheck, CheckResult
import re


# class MissingDescriptiveStatisticsCheck(BaseCheck):
#     name = "Zcela chybí popisná charakteristika"
#     penalty = -100

#     REQUIRED_FUNCTIONS = {
#         # "COUNT",
#         "MIN",
#         "MAX",
#         "AVERAGE",
#         "MEDIAN",
#         # "SUM",
#     }

#     FUNC_RE = re.compile(r"=([A-Z]+)\(", re.IGNORECASE)

#     def run(self, document, assignment=None):

#         if "data" not in document.sheet_names():
#             return CheckResult(
#                 False,
#                 'List "data" chybí – nelze vytvořit popisnou charakteristiku.',
#                 self.penalty,
#                 fatal=True,
#             )

#         found = set()

#         for cell in document.cells_with_formulas():
#             if cell["sheet"] != "data":
#                 continue

#             formula = cell["formula"]
#             if not isinstance(formula, str):
#                 continue

#             m = self.FUNC_RE.match(formula)
#             if not m:
#                 continue

#             funcs = self.FUNC_RE.findall(formula)
#             for func in funcs:
#                 func = func.upper()
#                 if func in self.REQUIRED_FUNCTIONS:
#                     found.add(func)

#         missing = self.REQUIRED_FUNCTIONS - found

#         if missing:
#             return CheckResult(
#                 False,
#                 "Chybí popisná charakteristika – nebyly nalezeny funkce: "
#                 + ", ".join(sorted(missing)),
#                 self.penalty,
#                 fatal=True,
#             )

#         return CheckResult(
#             True,
#             "Popisná charakteristika je přítomna.",
#             0,
#         )

from checks.base_check import BaseCheck, CheckResult
import re


class MissingDescriptiveStatisticsCheck(BaseCheck):
    name = "Zcela chybí popisná charakteristika"
    penalty = -100

    REQUIRED_FUNCTIONS = {
        "MIN",
        "MAX",
        "AVERAGE",
        "MEDIAN",
    }

    FUNC_RE = re.compile(r"=([A-Z]+)\(", re.IGNORECASE)

    def run(self, document, assignment=None):

        if "data" not in document.sheet_names():
            return CheckResult(
                False,
                'List "data" chybí – nelze vytvořit popisnou charakteristiku.',
                self.penalty,
                fatal=True,
            )

        found = set()

        for item in document.iter_formulas():
            if item["sheet"] != "data":
                continue

            formula = item["formula"]
            if not isinstance(formula, str):
                continue

            for func in self.FUNC_RE.findall(formula):
                func = func.upper()
                if func in self.REQUIRED_FUNCTIONS:
                    found.add(func)

        missing = self.REQUIRED_FUNCTIONS - found

        if missing:
            return CheckResult(
                False,
                "Chybí popisná charakteristika – nebyly nalezeny funkce: "
                + ", ".join(sorted(missing)),
                self.penalty,
                fatal=True,
            )

        return CheckResult(
            True,
            "Popisná charakteristika je přítomna.",
            0,
        )