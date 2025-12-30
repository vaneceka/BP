from checks.word.base_check import BaseCheck, CheckResult


class MissingListOfFiguresCheck(BaseCheck):
    name = "Chybí seznam obrázků"
    penalty = -100

    def run(self, document, assignment=None):
        images = [o for o in document.iter_objects() if o["type"] == "image"]

        if not images:
            return CheckResult(True, "Dokument neobsahuje obrázky.", 0)

        for i in range(document.section_count()):
            if document.has_list_of_figures_in_section(i):
                return CheckResult(True, "Seznam obrázků existuje.", 0)

        return CheckResult(False, "Seznam obrázků zcela chybí.", self.penalty)