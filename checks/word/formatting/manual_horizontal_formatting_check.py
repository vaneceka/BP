import re
from checks.base_check import BaseCheck, CheckResult

BAD_PATTERNS = [
    r" {4,}",   # 4+ mezer
    r"\.{4,}",  # 4+ teček
    r"-{4,}",   # 4+ pomlček
    r"_{4,}",   # 4+ podtržítek
]

class ManualHorizontalSpacingCheck(BaseCheck):
    name = "Ruční horizontální zarovnání pomocí mezer/znaků"
    penalty = -5 

    def run(self, document, assignment=None):
        found = 0

        for p in document.iter_paragraphs():
            raw = document.paragraph_text_raw(p)
            if not raw:
                continue

            style_id = document._paragraph_style_id(p)
            if style_id and "toc" in style_id.lower():
                continue

            for pat in BAD_PATTERNS:
                if re.search(pat, raw):
                    found += 1
                    break

        if found > 0:
            return CheckResult(
                False,
                "Nalezeno ruční horizontální formátování.",
                found * self.penalty,
            )

        return CheckResult(True, "Nenalezeno ruční horizontální formátování.", 0)