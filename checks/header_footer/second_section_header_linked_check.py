from checks.base_check import BaseCheck, CheckResult


class SecondSectionHeaderLinkedCheck(BaseCheck):
    name = "Druhý oddíl je propojen s předchozím oddílem v záhlaví"
    penalty = -2

    R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

    def run(self, document, assignment=None):

        if document.section_count() < 2:
            return CheckResult(True, "Dokument má méně než dva oddíly.", 0)

        sect1 = document.section_properties(0)
        sect2 = document.section_properties(1)

        if sect1 is None or sect2 is None:
            return CheckResult(True, "Oddíly nemají nastavení.", 0)

        headers1 = sect1.findall("w:headerReference", document.NS)
        headers2 = sect2.findall("w:headerReference", document.NS)

        # ❗️ KLÍČOVÉ PRAVIDLO:
        # žádné headerReference ve 2. oddílu = propojeno
        if not headers2:
            return CheckResult(
                False,
                "Záhlaví druhého oddílu je propojeno s předchozím oddílem.",
                self.penalty,
            )

        def header_map(headers):
            result = {}
            for h in headers:
                h_type = h.attrib.get(f"{{{document.NS['w']}}}type", "default")
                r_id = h.attrib.get(f"{{{self.R_NS}}}id")
                if r_id:
                    result[h_type] = r_id
            return result

        map1 = header_map(headers1)
        map2 = header_map(headers2)

        # porovnej stejné typy (default / first / even)
        for h_type in map1:
            if h_type in map2 and map1[h_type] == map2[h_type]:
                return CheckResult(
                    False,
                    "Záhlaví druhého oddílu je propojeno s předchozím oddílem.",
                    self.penalty,
                )

        return CheckResult(
            True,
            "Záhlaví druhého oddílu není propojeno s předchozím.",
            0,
        )