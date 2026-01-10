from checks.base_check import BaseCheck, CheckResult

# class WrapTextCheck(BaseCheck):
#     name = "Není zalamování textu v buňce"
#     penalty = -1
#     SHEET = "data"

#     MIN_TEXT_LENGTH = 20

#     def run(self, document, assignment=None):

#         if self.SHEET not in document.wb.sheetnames:
#             return CheckResult(
#                 False,
#                 f'Chybí list "{self.SHEET}".',
#                 self.penalty,
#                 fatal=True
#             )

#         ws = document.wb[self.SHEET]
#         problems = []

#         for row in ws.iter_rows():
#             for cell in row:
#                 value = cell.value

#                 if not isinstance(value, str):
#                     continue

#                 text = value.strip()

#                 if len(text) < self.MIN_TEXT_LENGTH or text.startswith("="):
#                     continue

#                 wrap = cell.alignment.wrap_text

#                 if wrap is not True:
#                     problems.append(
#                         f"{cell.coordinate}: chybí zalamování textu"
#                     )

#         if problems:
#             return CheckResult(
#                 False,
#                 "Chybí zalamování textu v buňkách:\n"
#                 + "\n".join("– " + p for p in problems[:10]),
#                 self.penalty,
#                 fatal=True
#             )

#         return CheckResult(
#             True,
#             "Zalamování textu je nastaveno správně.",
#             0
#         )

class WrapTextCheck(BaseCheck):
    name = "Není zalamování textu v buňce"
    penalty = -1
    SHEET = "data"
    MIN_TEXT_LENGTH = 20

    def run(self, document, assignment=None):

        problems = []

        for addr in document.iter_cells(self.SHEET):
            value = document.get_cell_value(self.SHEET, addr)

            if not isinstance(value, str):
                continue
            if document.has_formula(self.SHEET, addr):
                continue

            text = value.strip()
            if len(text) < self.MIN_TEXT_LENGTH:
                continue

            style = document.get_cell_style(self.SHEET, addr)
            if not style or style.get("wrap") is not True:
                problems.append(f"{addr}: chybí zalamování textu")

        if problems:
            return CheckResult(
                False,
                "Chybí zalamování textu v buňkách:\n"
                + "\n".join("– " + p for p in problems[:10]),
                self.penalty,
                fatal=True
            )

        return CheckResult(True, "Zalamování textu je nastaveno správně.", 0)