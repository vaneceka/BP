from checks.base_check import BaseCheck, CheckResult


class TOCFirstSectionContentCheck(BaseCheck):
    name = "Obsah obsahuje text z prvního oddílu"
    penalty = -10  # násobí se

    def run(self, document, assignment=None):

        # 1️⃣ najdi sekci s obsahem
        toc_section = None
        for i in range(document.section_count()):
            if document.has_toc_in_section(i):
                toc_section = i
                break

        if toc_section is None:
            return CheckResult(True, "Obsah neexistuje.", 0)

        # 2️⃣ texty položek v obsahu
        toc_items = []
        for p in document.section(toc_section):
            if not p.tag.endswith("}p"):
                continue

            if document._paragraph_is_toc_or_object_list(p):
                txt = document._paragraph_text(p)
                if txt:
                    toc_items.append(txt)

        if not toc_items:
            return CheckResult(True, "Obsah neobsahuje žádné položky.", 0)

        # 3️⃣ nadpisy z 1. oddílu
        first_section_headings = set()

        for p in document.section(0):
            if not p.tag.endswith("}p"):
                continue

            txt = document._paragraph_text(p)
            if not txt:
                continue

            style_id = document._paragraph_style_id(p)
            if not style_id:
                continue

            lvl = document._style_level_from_styles_xml(style_id)
            if lvl is not None:
                first_section_headings.add(txt)

        # 4️⃣ porovnání
        illegal = [
            item for item in toc_items
            if any(h in item for h in first_section_headings)
        ]

        if illegal:
            return CheckResult(
                False,
                "Obsah obsahuje nadpisy z prvního oddílu:\n"
                + "\n".join(f"– {t}" for t in illegal[:5]),
                self.penalty * len(illegal),
            )

        return CheckResult(
            True,
            "Obsah neobsahuje žádný text z prvního oddílu.",
            0,
        )