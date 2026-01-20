from checks.base_check import BaseCheck, CheckResult


class TOCFirstSectionContentCheck(BaseCheck):
    name = "Obsah obsahuje text z prvního oddílu"
    penalty = -10

    def run(self, document, assignment=None):
        toc_section = next(
            (i for i in range(document.section_count()) if document.has_toc_in_section(i)),
            None
        )

        if toc_section is None:
            return CheckResult(True, "Obsah neexistuje.", 0)

        toc_items = []
        for p in document.iter_section_paragraphs(toc_section):
            sid = (document._paragraph_style_id(p) or "").strip().lower()
            is_toc = (
                (sid.startswith("toc") and sid[3:].isdigit()) or
                (sid.startswith("obsah") and sid[5:].isdigit())
            )

            if not is_toc:
                break

            txt = document._paragraph_text(p)
            if txt:
                toc_items.append(txt)

        if not toc_items:
            return CheckResult(False, "Obsah byl nalezen, ale nepodařilo se načíst žádné položky.", 0)

        first_headings = set()
        for p in document.iter_section_paragraphs(0):
            txt = document._paragraph_text(p)
            if not txt:
                continue

            sid = document._paragraph_style_id(p)
            if not sid:
                continue

            if document._style_level_from_styles_xml(sid) is not None:
                first_headings.add(txt)

        illegal = [item for item in toc_items if any(h in item for h in first_headings)]

        if illegal:
            return CheckResult(
                False,
                "Obsah obsahuje nadpisy z prvního oddílu:\n" +
                "\n".join(f"– {t}" for t in illegal[:5]),
                self.penalty * len(illegal),
            )

        return CheckResult(True, "Obsah neobsahuje žádný text z prvního oddílu.", 0)