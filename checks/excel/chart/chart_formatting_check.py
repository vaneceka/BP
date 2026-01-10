from checks.base_check import BaseCheck, CheckResult
from openpyxl.chart.bar_chart import BarChart

# class ChartFormattingCheck(BaseCheck):
#     name = (
#         "V grafu chybí název, popisy os nebo popisky dat"
#     )
#     penalty = -2
#     SHEET = "data"

#     def _chart_title_text(self, title):
#         if not title or not title.tx or not title.tx.rich:
#             return None
#         return title.tx.rich.p[0].r[0].t.strip()


#     def run(self, document, assignment=None):

#         if assignment is None or not assignment.chart:
#             return CheckResult(True, "Zadání neobsahuje graf.", 0)

#         if self.SHEET not in document.wb.sheetnames:
#             return CheckResult(
#                 False,
#                 f'Chybí list "{self.SHEET}".',
#                 self.penalty,
#                 fatal=True,
#             )

#         ws = document.wb[self.SHEET]

#         if not ws._charts:
#             return CheckResult(
#                 False,
#                 "Graf v sešitu chybí.",
#                 self.penalty,
#                 fatal=True,
#             )

#         chart: BarChart = ws._charts[0]
#         expected = assignment.chart

#         errors = []

#         expected_title = expected.get("title")
#         if expected_title:
#             actual_title = self._chart_title_text(chart.title)
#             if actual_title != expected_title:
#                 errors.append(
#                     f'název grafu (oček. „{expected_title}“, nalezeno „{actual_title}“) '
#                 )

#         expected_x = expected.get("xAxisLabel")
#         if expected_x:
#             actual_x = None

#             if chart.x_axis and chart.x_axis.title:
#                 actual_x = chart.x_axis.title.tx.rich.p[0].r[0].t
            
#             if actual_x != expected_x:
#                 errors.append(
#                     f'osa X (oček. „{expected_x}“, nalezeno „{actual_x}“) '
#                 )

#         expected_y = expected.get("yAxisLabel")
#         if expected_y:
#             actual_y = None
                
#             if chart.y_axis and chart.y_axis.title:
#                 actual_y = chart.y_axis.title.tx.rich.p[0].r[0].t

#             if actual_y != expected_y:
#                 errors.append(
#                     f'osa Y (oček. „{expected_y}“, nalezeno „{actual_y}“) '
#                 )

#         has_labels = False
#         for s in chart.series:
#             if s.dLbls and s.dLbls.showVal:
#                 has_labels = True
#                 break
        
#         if not has_labels:
#             errors.append("chybí popisky dat")

#         if errors:
#             return CheckResult(
#                 False,
#                 "V grafu chybí:\n– " + "\n– ".join(errors),
#                 self.penalty * len(errors),
#                 fatal=True,
#             )

#         return CheckResult(
#             True,
#             "Graf má název, popisy os i popisky dat.",
#             0,
#         )

class ChartFormattingCheck(BaseCheck):
    name = "V grafu chybí název, popisy os nebo popisky dat"
    penalty = -2
    SHEET = "data"

    def run(self, document, assignment=None):

        if not assignment or not assignment.chart:
            return CheckResult(True, "Zadání neobsahuje graf.", 0)

        if not document.has_chart(self.SHEET):
            return CheckResult(
                False,
                "Graf v sešitu chybí.",
                self.penalty,
                fatal=True,
            )

        expected = assignment.chart
        errors = []

        if expected.get("title"):
            if document.chart_title(self.SHEET) != expected["title"]:
                errors.append("název grafu")

        if expected.get("xAxisLabel"):
            if document.chart_x_label(self.SHEET) != expected["xAxisLabel"]:
                errors.append("osa X")

        if expected.get("yAxisLabel"):
            if document.chart_y_label(self.SHEET) != expected["yAxisLabel"]:
                errors.append("osa Y")

        if not document.chart_has_data_labels(self.SHEET):
            errors.append("popisky dat")

        if errors:
            return CheckResult(
                False,
                "V grafu chybí:\n– " + "\n– ".join(errors),
                self.penalty * len(errors),
                fatal=True,
            )

        return CheckResult(True, "Graf má správné formátování.", 0)