from checks.base_check import BaseCheck, CheckResult


class HeaderFooterMissingCheck(BaseCheck):
    name = "Záhlaví nebo zápatí v dokumentu není řešeno"
    penalty = -30

    def run(self, document, assignment=None):

        has_content = False

        for i in range(document.section_count()):
            sect_pr = document.section_properties(i)
            if sect_pr is None:
                continue

            refs = (
                sect_pr.findall("w:headerReference", document.NS)
                + sect_pr.findall("w:footerReference", document.NS)
            )

            for ref in refs:
                # 🔑 POZOR – relationship namespace
                r_id = ref.attrib.get(
                    "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
                )
                if not r_id:
                    continue

                xml = document.load_part_by_rid(r_id)
                if xml is None:
                    continue

                # 🔍 text
                for t in xml.findall(".//w:t", document.NS):
                    if t.text and t.text.strip():
                        has_content = True
                        break

                # 🔍 pole (PAGE, DATE, atd.)
                for instr in xml.findall(".//w:instrText", document.NS):
                    if instr.text and instr.text.strip():
                        has_content = True
                        break

                if has_content:
                    break

            if has_content:
                break

        if not has_content:
            return CheckResult(
                False,
                "Dokument neobsahuje žádné záhlaví ani zápatí.",
                self.penalty,
            )

        return CheckResult(
            True,
            "Záhlaví nebo zápatí je v dokumentu řešeno.",
            0,
        )