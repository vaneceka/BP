from checks.word.base_check import BaseCheck, CheckResult
import re

# NOTE nefunguje. Je treab se na to poradne jeste podivat
class BibliographyISO690Check(BaseCheck):
    name = "Seznam literatury není dle ISO 690"
    penalty = -10

    def run(self, document, assignment=None):

        # 1️⃣ Bibliografie musí existovat (řeší jiný check)
        if not document.has_word_bibliography():
            return CheckResult(True, "Bibliografie chybí – řeší jiný check.", 0)

        items: list[str] = []

        # 2️⃣ Vytáhni položky z Word bibliografie (SDT)
        for sdt in document._xml.findall(".//w:sdt", document.NS):
            sdt_pr = sdt.find("w:sdtPr", document.NS)
            if sdt_pr is None or sdt_pr.find("w:bibliography", document.NS) is None:
                continue

            for p in sdt.findall(".//w:p", document.NS):
                ppr = p.find("w:pPr", document.NS)
                if ppr is None:
                    continue

                ps = ppr.find("w:pStyle", document.NS)
                if ps is None:
                    continue

                style = ps.attrib.get(
                    f"{{{document.NS['w']}}}val", ""
                ).lower()

                if style not in ("bibliografie", "bibliography"):
                    continue

                text = document._paragraph_text(p)
                if text:
                    items.append(text)

        if not items:
            return CheckResult(True, "Bibliografie neobsahuje položky.", 0)

        # 3️⃣ Heuristická ISO 690 validace
        iso_ok = 0

        for raw in items:
            # odeber číslování „1. “
            t = re.sub(r"^\s*\d+\.\s*", "", raw).strip()
            t_lower = t.lower()

            # ❌ zakázané dle ISO 690:2022
            if "[online]" in t_lower:
                continue

            # základní prvky
            has_author = re.search(
                r"^[A-ZÁČĎÉĚÍŇÓŘŠŤÚŮÝŽ][A-Za-zÁČĎÉĚÍŇÓŘŠŤÚŮÝŽ\-]+,\s*[A-Z]",
                t
            )
            has_year = re.search(r"(19|20)\d{2}", t)
            has_url = "http://" in t_lower or "https://" in t_lower
            has_cit = re.search(r"\[cit\.\s*\d{4}-\d{2}-\d{2}\]", t_lower)

            # 🔍 rozhodnutí o typu zdroje
            is_online = has_url or has_cit

            # 🌐 ONLINE ZDROJ
            if is_online:
                if has_author and has_year and has_url and has_cit:
                    iso_ok += 1
                continue

            # 📘 TIŠTĚNÝ ZDROJ
            has_place_publisher = ":" in t and "," in t
            if has_author and has_year and has_place_publisher:
                iso_ok += 1

        ratio = iso_ok / len(items)

        # méně než 70 % položek OK → FAIL
        if ratio < 0.7:
            return CheckResult(
                False,
                "Seznam literatury neodpovídá normě ISO 690 (heuristicky).",
                self.penalty,
            )

        return CheckResult(
            True,
            "Seznam literatury odpovídá normě ISO 690 (heuristicky).",
            0,
        )