from checks.base_check import BaseCheck, CheckResult

#NOTE predelat pro ODT
class ListLevel2UsedCheck(BaseCheck):
    name = "Použití seznamu 2. úrovně"
    penalty = -1 


    STYLE_IDS = {
        "slovanseznam2",        
        "seznamsodrkami2",      
        "cislovanyseznam2", 
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

            if style_id.strip().lower() in self.STYLE_IDS:
                return CheckResult(True, "Seznam druhé úrovně je použit.", 0)

        return CheckResult(False, "V dokumentu není použit seznam druhé úrovně.", self.penalty)