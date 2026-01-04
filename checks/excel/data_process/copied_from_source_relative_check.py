import re
from checks.base_check import BaseCheck, CheckResult
from openpyxl.utils import range_boundaries


class CopiedFromSourceByRelativeRefCheck(BaseCheck):
    name = "Záhlaví nebo hodnoty na list „data“ nejsou zkopírovány odkazem"
    penalty = -5

    DIRECT_REF_RE = re.compile(r"^=\s*'?zdroj'?!([A-Z]{1,3}[0-9]+)\s*$", re.IGNORECASE)

    def run(self, document, assignment=None):

        if "zdroj" not in document.sheet_names() or "data" not in document.sheet_names():
            return CheckResult(True, "Chybí list 'zdroj' nebo 'data' – check přeskočen.", 0)

        ws_src = document.wb["zdroj"]
        ws_data = document.wb["data"]

        dim = ws_src.calculate_dimension()
        min_col, min_row, max_col, max_row = range_boundaries(dim)

        problems = []

        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                src_cell = ws_src.cell(row=r, column=c)
                data_cell = ws_data.cell(row=r, column=c)

                if src_cell.value is None:
                    continue

                if data_cell.data_type != "f" or not isinstance(data_cell.value, str):
                    problems.append(f"data!{data_cell.coordinate}: hodnota není převzatá odkazem")
                    continue

                formula = data_cell.value.strip()

                m = self.DIRECT_REF_RE.match(formula)
                if not m:
                    problems.append(f"data!{data_cell.coordinate}: není přímý odkaz na zdroj ({formula})")
                    continue

                ref_addr = m.group(1).upper()
                if ref_addr != data_cell.coordinate.upper():
                    problems.append(
                        f"data!{data_cell.coordinate}: odkazuje na jinou buňku zdroje (={ref_addr})"
                    )
                    continue

                if "$" in formula:
                    problems.append(f"data!{data_cell.coordinate}: odkaz není relativní ({formula})")

        if problems:
            return CheckResult(
                False,
                "Záhlaví nebo hodnoty na list „data“ nejsou zkopírovány odkazem:\n"
                + "\n".join(f"- {p}" for p in problems[:10]),
                self.penalty,
            )

        return CheckResult(True, "Záhlaví i hodnoty jsou zkopírovány odkazem (relativně) ze 'zdroj'.", 0)