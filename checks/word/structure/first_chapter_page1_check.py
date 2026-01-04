from checks.base_check import BaseCheck, CheckResult


class FirstChapterStartsOnPageOneCheck(BaseCheck):
    name = "První kapitola ve druhém oddílu nezačíná na straně 1"
    penalty = -5

    def run(self, document, assignment=None):
        if document.section_count() < 2:
            return CheckResult(
                False,
                "Dokument nemá druhý oddíl – nelze ověřit začátek kapitol.",
                self.penalty,
            )

        sect_pr = document.section_properties(1)
        if sect_pr is None:
            return CheckResult(
                False,
                "Ve druhém oddílu chybí nastavení oddílu (sectPr).",
                self.penalty,
            )

        pg_num = sect_pr.find("w:pgNumType", document.NS)
        if pg_num is None:
            return CheckResult(
                False,
                "Ve druhém oddílu není nastaveno číslování stránek.",
                self.penalty,
            )

        if pg_num.attrib.get(f"{{{document.NS['w']}}}start") != "1":
            return CheckResult(
                False,
                "Číslování stránek ve druhém oddílu nezačíná od 1.",
                self.penalty,
            )

        section = document.section(1)
        first_h1 = None

        for el in section:
            if not el.tag.endswith("}p"):
                continue

            txt = document._paragraph_text(el)
            if not txt:
                continue

            style_id = document._paragraph_style_id(el)
            if not style_id:
                continue

            lvl = document._style_level_from_styles_xml(style_id)
            if lvl == 1:
                first_h1 = el
                break

        if first_h1 is None:
            return CheckResult(
                False,
                "Ve druhém oddílu nebyla nalezena žádná kapitola (Heading 1).",
                self.penalty,
            )

        for el in section:
            if el is first_h1:
                break

            if el.tag.endswith("}p"):
                if document.paragraph_is_generated_by_field(el):
                    continue
                if document._paragraph_is_toc_or_object_list(el):
                    continue

                if document._paragraph_text(el):
                    return CheckResult(
                        False,
                        "Před první kapitolou je ve druhém oddílu viditelný obsah.",
                        self.penalty,
                    )

            if el.tag.endswith("}tbl"):
                return CheckResult(
                    False,
                    "Před první kapitolou je ve druhém oddílu objekt.",
                    self.penalty,
                )

        return CheckResult(
            True,
            "První kapitola ve druhém oddílu začíná na straně 1.",
            0,
        )