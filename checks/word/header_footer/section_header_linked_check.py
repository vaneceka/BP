from checks.word.base_check import BaseCheck, CheckResult


class HeaderNotLinkedToPreviousCheck(BaseCheck):
    penalty = -2

    def __init__(self, section_number: int):
        self.section_number = section_number
        self.section_index = section_number - 1

        self.name = (
            f"{section_number}. oddíl není propojen s předchozím oddílem v záhlaví"
        )

    def run(self, document, assignment=None):

        # první oddíl nemá předchozí -> vždy OK
        if self.section_index == 0:
            return CheckResult(True, "První oddíl nemá předchozí oddíl.", 0)

        if document.section_count() <= self.section_index:
            return CheckResult(True, "Oddíl v dokumentu neexistuje.", 0)

        sect_pr = document.section_properties(self.section_index)
        if sect_pr is None:
            return CheckResult(
                False,
                "Oddíl nemá nastavení záhlaví.",
                self.penalty
            )

        header_refs = sect_pr.findall("w:headerReference", document.NS)

        # žádný headerReference = implicitně zděděné
        if not header_refs:
            return CheckResult(
                False,
                "Záhlaví je převzaté z předchozího oddílu.",
                self.penalty
            )

        for ref in header_refs:
            if ref.find("w:linkToPrevious", document.NS) is not None:
                return CheckResult(
                    False,
                    "Záhlaví je propojené s předchozím oddílem.",
                    self.penalty
                )

        return CheckResult(
            True,
            "Záhlaví není propojené s předchozím oddílem.",
            0
        )