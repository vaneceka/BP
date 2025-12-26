from checks.base_check import BaseCheck, CheckResult


class ListLevel2UsedCheck(BaseCheck):
    name = "Použití seznamu 2. úrovně"
    penalty = -1  # násobí se
    # NOTE mozna, kdyz bude word v anglictine, bude problem. -> Zjistit!
    # sem si dej všechny varianty, které v praxi v docx vídáš
    LEVEL2_STYLE_IDS = {
        "slovanseznam2",        # číslovaný seznam 2
        "seznamsodrkami2",      # odrážky 2
        "cislovanyseznam2",     # fallback
        "numberlist2",
        "bulletlist2",
    }

    def run(self, document, assignment=None):
        for p in document.iter_paragraphs():
            ppr = p.find("w:pPr", document.NS)
            if ppr is None:
                continue

            ps = ppr.find("w:pStyle", document.NS)
            if ps is None:
                continue

            style_id = ps.attrib.get(f"{{{document.NS['w']}}}val")
            if not style_id:
                continue

            if style_id.strip().lower() in self.LEVEL2_STYLE_IDS:
                return CheckResult(True, "Seznam druhé úrovně je použit.", 0)

        return CheckResult(False, "V dokumentu není použit seznam druhé úrovně.", self.penalty)