from checks.word.base_check import BaseCheck, CheckResult


class ListOfFiguresNotUpdatedCheck(BaseCheck):
    name = "Seznam obrázků není aktuální"
    penalty = -5

    def run(self, document, assignment=None):

        captions = document.iter_figure_caption_texts()
        toc_items = document.iter_list_of_figures_texts()

        missing = [
            c for c in captions
            if not any(c in t for t in toc_items)
        ]

        extra = [
            t for t in toc_items
            if not any(c in t for c in captions)
        ]

        if missing or extra:
            msg = []

            if missing:
                msg.append(f"V seznamu obrázků chybí {len(missing)} položek.")

            if extra:
                msg.append(f"Seznam obrázků obsahuje {len(extra)} neplatných položek.")

            return CheckResult(
                False,
                "Seznam obrázků není aktuální:\n" + "\n".join(msg),
                self.penalty,
            )

        return CheckResult(True, "Seznam obrázků je aktuální.", 0)