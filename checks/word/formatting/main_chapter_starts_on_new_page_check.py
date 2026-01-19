# from checks.base_check import BaseCheck, CheckResult

# class MainChapterStartsOnNewPageCheck(BaseCheck):
#     name = "Hlavní kapitola nezačíná na nové straně"
#     penalty = -2 

#     def run(self, document, assignment=None):
#         errors = []
#         total_penalty = 0

#         for p in document.iter_paragraphs():
#             text = document._paragraph_text(p)
#             if not text:
#                 continue

#             style_id = document._paragraph_style_id(p)
#             if not style_id:
#                 continue

#             level = document._style_level_from_styles_xml(style_id)
#             if level != 1:
#                 continue 

#             if document.paragraph_has_page_break(p):
#                 continue

#             if document.style_has_page_break(style_id):
#                 continue

#             errors.append(
#                 f"Hlavní kapitola „{text}“ nezačíná na nové straně."
#             )
#             total_penalty += self.penalty

#         if errors:
#             return CheckResult(
#                 False,
#                 "\n".join(errors),
#                 total_penalty,
#             )

#         return CheckResult(
#             True,
#             "Všechny hlavní kapitoly začínají na nové straně.",
#             0,
#         )

from checks.base_check import BaseCheck, CheckResult


class MainChapterStartsOnNewPageCheck(BaseCheck):
    name = "Hlavní kapitola nezačíná na nové straně"
    penalty = -2

    def run(self, document, assignment=None):
        errors = []
        total_penalty = 0

        for h in document.iter_main_headings():
            text = document.get_visible_text(h)

            if document.heading_starts_on_new_page(h):
                continue

            errors.append(
                f"Hlavní kapitola „{text}“ nezačíná na nové straně."
            )
            total_penalty += self.penalty

        if errors:
            return CheckResult(False, "\n".join(errors), total_penalty)

        return CheckResult(
            True,
            "Všechny hlavní kapitoly začínají na nové straně.",
            0,
        )