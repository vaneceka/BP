from checks.base_check import BaseCheck, CheckResult


# class FooterLinkedToPreviousCheck(BaseCheck):
#     penalty = -2

#     def __init__(self, section_number: int):
#         self.section_number = section_number
#         self.section_index = section_number - 1

#         self.name = (
#             f"{section_number}. oddíl je propojen s předchozím oddílem v zápatí"
#         )

#     def run(self, document, assignment=None):
#         if self.section_index == 0:
#             return CheckResult(True, "První oddíl nemá předchozí oddíl.", 0)

#         if document.section_count() <= self.section_index:
#             return CheckResult(True, "Oddíl v dokumentu neexistuje.", 0)

#         sect_prev = document.section_properties(self.section_index - 1)
#         sect_curr = document.section_properties(self.section_index)

#         if sect_prev is None or sect_curr is None:
#             return CheckResult(True, "Oddíly nemají nastavení.", 0)

#         footers_prev = sect_prev.findall("w:footerReference", document.NS)
#         footers_curr = sect_curr.findall("w:footerReference", document.NS)

#         if not footers_curr:
#             return CheckResult(
#                 False,
#                 "Zápatí oddílu je převzaté z předchozího oddílu.",
#                 self.penalty,
#             )

#         def footer_map(footers):
#             result = {}
#             for f in footers:
#                 f_type = f.attrib.get(f"{{{document.NS['w']}}}type", "default")
#                 r_id = f.attrib.get(f"{{{document.NS['r']}}}id")
#                 if r_id:
#                     result[f_type] = r_id
#             return result

#         map_prev = footer_map(footers_prev)
#         map_curr = footer_map(footers_curr)

#         for f_type in map_prev:
#             if f_type in map_curr and map_prev[f_type] == map_curr[f_type]:
#                 return CheckResult(
#                     False,
#                     "Zápatí oddílu je propojené s předchozím oddílem.",
#                     self.penalty,
#                 )

#         return CheckResult(
#             True,
#             "Zápatí oddílu není propojené s předchozím oddílem.",
#             0,
#         )

class FooterLinkedToPreviousCheck(BaseCheck):
    penalty = -2

    def __init__(self, section_number: int):
        self.section_number = section_number
        self.section_index = section_number - 1
        self.name = (
            f"{section_number}. oddíl je propojen s předchozím oddílem v zápatí"
        )

    def run(self, document, assignment=None):
        if self.section_index == 0:
            return CheckResult(True, "První oddíl nemá předchozí oddíl.", 0)

        if document.section_count() <= self.section_index:
            return CheckResult(True, "Oddíl v dokumentu neexistuje.", 0)

        linked = document.footer_is_linked_to_previous(self.section_index)

        if linked is None:
            return CheckResult(
                True,
                "Propojení zápatí mezi oddíly není v tomto formátu podporováno.",
                0,
            )

        if linked:
            return CheckResult(
                False,
                "Zápatí oddílu je propojené s předchozím oddílem.",
                self.penalty,
            )

        return CheckResult(
            True,
            "Zápatí oddílu není propojené s předchozím oddílem.",
            0,
        )