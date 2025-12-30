import re
from checks.base_check import BaseCheck, CheckResult

BAD_PATTERNS = [
    r" {4,}",   # 4+ mezer
    r"\.{4,}",  # 4+ teček
    r"-{4,}",   # 4+ pomlček
    r"_{4,}",   # 4+ podtržítek
]

# NOTE mozna opravit aby to nevypisovalo v jakych odstavcich
class ManualHorizontalSpacingCheck(BaseCheck):
    name = "Ruční horizontální zarovnání pomocí mezer/znaků"
    penalty = -5  # násobí se

    def run(self, document, assignment=None):
        errors = []

        for p in document.iter_paragraphs():
            raw = document.paragraph_text_raw(p)
            if not raw:
                continue

            # ignoruj TOC
            style_id = document._paragraph_style_id(p)
            if style_id and "toc" in style_id.lower():
                continue

            for pat in BAD_PATTERNS:
                if re.search(pat, raw):
                    sample = raw.replace("\t", "\\t")
                    errors.append(f"„{sample[:120]}“")
                    break

        if errors:
            return CheckResult(
                False,
                "Nalezeno ruční horizontální formátování v odstavcích:\n- " + "\n- ".join(errors),
                len(errors) * self.penalty,
            )

        return CheckResult(True, "Nenalezeno ruční horizontální formátování.", 0)