from checks.base_check import BaseCheck, CheckResult


class ListOfFiguresNotUpdatedCheck(BaseCheck):
    name = "Seznam obrázků není aktuální"
    penalty = -5

    def run(self, document, assignment=None):

        captions = document.count_figure_captions()
        if captions == 0:
            return CheckResult(True, "Dokument neobsahuje obrázky.", 0)

        toc_items = document.count_list_of_figures_items()
        if toc_items == 0:
            return CheckResult(True, "Seznam obrázků neexistuje.", 0)

        if captions != toc_items:
            return CheckResult(
                False,
                f"Seznam obrázků není aktuální "
                f"(titulky: {captions}, položky v seznamu: {toc_items}).",
                self.penalty,
            )

        return CheckResult(True, "Seznam obrázků je aktuální.", 0)