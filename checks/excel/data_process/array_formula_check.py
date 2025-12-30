from checks.base_check import BaseCheck, CheckResult


class ArrayFormulaCheck(BaseCheck):
    name = "Použit maticový vzorec"
    penalty = -10

    def run(self, document, assignment=None):

        array_cells = []

        for sheet_name, data in document.sheets.items():
            xml = data["xml"]

            for c in xml.findall(".//main:c", document.NS):
                f = c.find("main:f", document.NS)
                if f is not None and f.attrib.get("t") == "array":
                    addr = c.attrib.get("r")
                    array_cells.append(f"{sheet_name}!{addr}")

        if array_cells:
            lines = "\n".join(f"- {c}" for c in array_cells[:5])
            return CheckResult(
                False,
                "Není použit klasický vzorec, ale maticový:\n" + lines,
                self.penalty,
            )

        return CheckResult(
            True,
            "Vzorce nejsou maticové.",
            0,
        )