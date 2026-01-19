# from checks.base_check import BaseCheck, CheckResult

# class HeadingHierarchicalNumberingCheck(BaseCheck):
#     name = (
#         "Nadpisy nemají aktivováno automatické "
#         "hierarchické číslování (např. 1.1.1)"
#     )
#     penalty = -5

#     def run(self, document, assignment=None):
#         broken = []

#         for level in (1, 2, 3):
#             expected = assignment.styles.get(f"Heading {level}")
#             actual = document.get_heading_style(level)

#             if expected is None or actual is None:
#                 continue

#             if not actual.isNumbered:
#                 broken.append(
#                     f"Nadpis {level}: není číslovaný"
#                 )
#                 continue

#             if actual.numLevel != level - 1:
#                 broken.append(
#                     f"Nadpis {level}: špatná úroveň "
#                     f"(očekáváno {level-1}, nalezeno {actual.numLevel})"
#                 )

#         if broken:
#             message = (
#                 "Nadpisy nemají správně nastavené "
#                 "automatické hierarchické číslování:\n"
#                 + "\n".join(f"- {b}" for b in broken)
#             )
#             return CheckResult(False, message, self.penalty)

#         return CheckResult(
#             True,
#             "Automatické hierarchické číslování nadpisů je nastaveno správně.",
#             0,
#         )

from checks.base_check import BaseCheck, CheckResult

class HeadingHierarchicalNumberingCheck(BaseCheck):
    name = (
        "Nadpisy nemají aktivováno automatické "
        "hierarchické číslování (např. 1.1.1)"
    )
    penalty = -5

    def run(self, document, assignment=None):
        broken = []

        for level in (1, 2, 3):
            is_numbered, is_hierarchical, num_level = \
                document.get_heading_numbering_info(level)

            if not is_numbered:
                broken.append(f"Nadpis {level}: není číslovaný")
                continue

            if not is_hierarchical:
                broken.append(
                    f"Nadpis {level}: není plně hierarchické číslování "
                    f"(např. 1.1.{'1' if level == 3 else ''})"
                )
                continue

            if num_level != level - 1:
                broken.append(
                    f"Nadpis {level}: špatná úroveň "
                    f"(očekáváno {level-1}, nalezeno {num_level})"
                )

        if broken:
            return CheckResult(
                False,
                "Nadpisy nemají správně nastavené "
                "automatické hierarchické číslování:\n"
                + "\n".join(f"- {b}" for b in broken),
                self.penalty,
            )

        return CheckResult(
            True,
            "Automatické hierarchické číslování nadpisů je nastaveno správně.",
            0,
        )