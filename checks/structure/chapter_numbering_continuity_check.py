from checks.base_check import BaseCheck, CheckResult


class ChapterNumberingContinuityCheck(BaseCheck):
    name = "Číslování kapitol nepokračuje mezi oddíly"
    penalty = -5

    def run(self, document, assignment=None):

        # Musí existovat alespoň 3 oddíly
        if document.section_count() < 3:
            return CheckResult(
                True,
                "Dokument nemá třetí oddíl – nelze ověřit kontinuitu.",
                0,
            )

        # --------------------------------------------------
        # 1️⃣ Najdi první Heading 1 ve 2. a 3. oddílu
        # --------------------------------------------------

        def first_h1(section_index):
            for p in document.section(section_index):
                if not p.tag.endswith("}p"):
                    continue

                if not document._paragraph_text(p):
                    continue

                style_id = document._paragraph_style_id(p)
                if not style_id:
                    continue

                lvl = document._style_level_from_styles_xml(style_id)
                if lvl == 1:
                    return p

            return None

        h1_sec2 = first_h1(1)
        h1_sec3 = first_h1(2)

        if h1_sec2 is None:
            return CheckResult(
                False,
                "Ve druhém oddílu nebyla nalezena žádná kapitola (Heading 1).",
                self.penalty,
            )

        if h1_sec3 is None:
            return CheckResult(
                False,
                "Ve třetím oddílu nebyla nalezena žádná kapitola (Heading 1).",
                self.penalty,
            )

        # --------------------------------------------------
        # 2️⃣ Zjisti numId kapitol (přes WordDocument!)
        # --------------------------------------------------

        num2 = document.get_heading_num_id(h1_sec2)
        num3 = document.get_heading_num_id(h1_sec3)

        if num2 is None or num3 is None:
            return CheckResult(
                False,
                "Kapitoly nejsou číslované – nelze ověřit kontinuitu.",
                self.penalty,
            )

        # --------------------------------------------------
        # 3️⃣ Porovnání kontinuity číslování
        # --------------------------------------------------

        if num2 != num3:
            return CheckResult(
                False,
                "Číslování kapitol ve třetím oddílu nepokračuje z předchozího oddílu.",
                self.penalty,
            )

        # --------------------------------------------------
        # ✔ Vše v pořádku
        # --------------------------------------------------

        return CheckResult(
            True,
            "Číslování kapitol mezi oddíly správně pokračuje.",
            0,
        )