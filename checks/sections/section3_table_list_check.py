from ..base_check import BaseCheck, CheckResult

class Section3TableListCheck(BaseCheck):
    name = "Seznam tabulek ve 3. oddílu"
    penalty = -5

    def run(self, document, assignment=None):

        # 1️⃣ zjisti, zda dokument vůbec obsahuje tabulky
        has_tables = any(
            obj["type"] == "table"
            for obj in document.iter_objects()
        )

        # ❗ žádné tabulky → seznam se NEVYŽADUJE
        if not has_tables:
            return CheckResult(
                True,
                "Dokument neobsahuje žádné tabulky – seznam tabulek není vyžadován.",
                0,
            )

        # 2️⃣ tabulky existují → seznam MUSÍ být
        if document.has_list_of_tables_in_section(2):
            return CheckResult(True, "Seznam tabulek nalezen.", 0)

        return CheckResult(
            False,
            "Dokument obsahuje tabulky, ale ve třetím oddílu chybí seznam tabulek.",
            self.penalty,
        )