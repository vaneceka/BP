from ..base_check import BaseCheck, CheckResult

class Section3FigureListCheck(BaseCheck):
    name = "Seznam obrázků ve 3. oddílu"
    penalty = -5

    def run(self, document, assignment=None):
        if document.has_list_of_figures_in_section(2):
            return CheckResult(True, "Seznam obrázků nalezen.", 0)
        return CheckResult(False, "Ve třetím oddílu chybí seznam obrázků.", self.penalty)