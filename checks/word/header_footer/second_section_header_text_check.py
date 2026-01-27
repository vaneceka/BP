# from checks.base_check import BaseCheck, CheckResult


# class SecondSectionHeaderHasTextCheck(BaseCheck):
#     name = "Druhý oddíl má text v záhlaví"
#     penalty = -2

#     def run(self, document, assignment=None):

#         if document.section_count() < 2:
#             return CheckResult(True, "Dokument má méně než dva oddíly.", 0)

#         sect_pr = document.section_properties(1)
#         if sect_pr is None:
#             return CheckResult(False, "Druhý oddíl nemá nastavení oddílu.", self.penalty)

#         header_refs = sect_pr.findall("w:headerReference", document.NS)
#         if not header_refs:
#             return CheckResult(False, "Druhý oddíl nemá žádné záhlaví.", self.penalty)

#         for ref in header_refs:
#             r_id = ref.attrib.get(f"{{{document.NS['r']}}}id")
#             if not r_id:
#                 continue

#             part_path = document.resolve_part_target(r_id)
#             if not part_path:
#                 continue

#             try:
#                 header_xml = document._load(part_path)
#             except KeyError:
#                 continue

#             for t in header_xml.findall(".//w:t", document.NS):
#                 if t.text and t.text.strip():
#                     return CheckResult(True, "Záhlaví druhého oddílu obsahuje text.", 0)

#         return CheckResult(False, "Záhlaví druhého oddílu je prázdné.", self.penalty)

from checks.base_check import BaseCheck, CheckResult


class SecondSectionHeaderHasTextCheck(BaseCheck):
    name = "Druhý oddíl má text v záhlaví"
    penalty = -2

    def run(self, document, assignment=None):

        if document.section_count() < 2:
            return CheckResult(True, "Dokument má méně než dva oddíly.", 0)

        if document.section_has_header_text(1):
            return CheckResult(
                True,
                "Záhlaví druhého oddílu obsahuje text.",
                0,
            )

        return CheckResult(
            False,
            "Záhlaví druhého oddílu je prázdné nebo neexistuje.",
            self.penalty,
        )