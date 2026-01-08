from checks.base_check import BaseCheck, CheckResult


# NOTE predelat i pro ods libre
class RedundantAbsoluteReferenceCheck(BaseCheck):
    name = "Nadbytečné použití absolutních adres"
    penalty = -10
    
    def run(self, document, assignment=None):

        if assignment is None or not hasattr(assignment, "cells"):
            return CheckResult(
                True,
                "Chybí excel assignment – check přeskočen.",
                0,
            )

        problems = []

        for addr, spec in assignment.cells.items():
            expected = spec.expression
            if not expected:
                continue

            sheet = "data"
            cell = document.get_cell(f"{sheet}!{addr}")
            if cell is None:
                continue

            student = cell["formula"]
            if not isinstance(student, str):
                continue

            if student.replace("$", "") == expected.replace("$", ""):
                if "$" in student and "$" not in expected:
                    problems.append(
                        f"{cell['sheet']}!{addr}: {student}"
                    )

        if problems:
            return CheckResult(
                False,
                "Nadbytečně použité absolutní adresy:\n"
                + "\n".join(f"- {p}" for p in problems[:5]),
                self.penalty,
            )

        return CheckResult(
            True,
            "Absolutní adresy jsou použity správně.",
            0,
        )