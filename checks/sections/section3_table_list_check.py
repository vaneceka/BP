from ..base_check import BaseCheck, CheckResult

class Section3TableListCheck(BaseCheck):
    name = "Seznam tabulek ve 3. oddílu"
    penalty = -5

    def run(self, document, assignment=None):
        if document.has_list_of_tables_in_section(2):
            return CheckResult(True, "Seznam tabulek nalezen.", 0)
        return CheckResult(False, "Ve třetím oddílu chybí seznam tabulek.", self.penalty)