from checks.word.base_check import BaseCheck, CheckResult

class MainChapterStartsOnNewPageCheck(BaseCheck):
    name = "Hlavní kapitola nezačíná na nové straně"
    penalty = -2  # násobí se

    def run(self, document, assignment=None):
        errors = []
        total_penalty = 0

        for p in document.iter_paragraphs():
            # zjisti text (pro hlášení)
            text = document._paragraph_text(p)
            if not text:
                continue

            # zjisti styl odstavce
            style_id = document._paragraph_style_id(p)
            if not style_id:
                continue

            # zjisti úroveň nadpisu
            level = document._style_level_from_styles_xml(style_id)
            if level != 1:
                continue  # není hlavní kapitola

            # page break přímo na odstavci
            if document.paragraph_has_page_break(p):
                continue

            # page break ve stylu (dědičně)
            if document.style_has_page_break(style_id):
                continue

            # chyba
            errors.append(
                f"Hlavní kapitola „{text}“ nezačíná na nové straně."
            )
            total_penalty += self.penalty

        if errors:
            return CheckResult(
                False,
                "\n".join(errors),
                total_penalty,
            )

        return CheckResult(
            True,
            "Všechny hlavní kapitoly začínají na nové straně.",
            0,
        )