from ...base_check import BaseCheck, CheckResult

class Section3TableListCheck(BaseCheck):
    name = "Seznam tabulek ve 3. oddílu"
    penalty = -5

    def run(self, document, assignment=None):
        has_tables = False

        for obj in document.iter_objects():
            if obj["type"] == "table":
                has_tables = True
                break

        if not has_tables:
            return CheckResult(
                True,
                "Dokument neobsahuje žádné tabulky – seznam tabulek není vyžadován.",
                0,
            )

        if document.has_list_of_tables_in_section(2):
            return CheckResult(True, "Seznam tabulek nalezen.", 0)

        return CheckResult(
            False,
            "Dokument obsahuje tabulky, ale ve třetím oddílu chybí seznam tabulek.",
            self.penalty,
        )