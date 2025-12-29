from checks.base_check import BaseCheck, CheckResult
import re

# NOTE poradne to jeste otestovat
class ObjectCaptionDescriptionCheck(BaseCheck):
    name = "V titulku chybí stručný popis objektu"
    penalty = -5

    def run(self, document, assignment=None):
        errors = []

        for obj in document.iter_objects():
            obj_type = obj["type"]
            element = obj["element"]

            if obj_type not in ("image", "chart", "table"):
                continue

            # 🔎 hledej titulek POUZE jako SEQ
            caption_p = None

            for p in (
                document.paragraph_before(element),
                document.paragraph_after(element),
            ):
                if p is not None and document.paragraph_has_seq_caption(p):
                    caption_p = p
                    break

            # chybějící titulek řeší jiný check
            if caption_p is None:
                continue

            text = document._paragraph_text(caption_p)
            if not text:
                errors.append("Titulek objektu je prázdný.")
                continue

            # odeber "Obrázek 1:", "Tabulka 2 –", "Graf 3"
            description = re.sub(
                r"^(Obrázek|Tabulka|Graf)\s+\d+\s*[:\-]?\s*",
                "",
                text,
                flags=re.IGNORECASE,
            ).strip()

            # ❌ žádný popis
            if not description:
                errors.append("V titulku chybí stručný popis objektu.")
                continue

            # ❌ příliš krátký popis (heuristika)
            if len(description) < 5 or len(description.split()) < 1:
                errors.append(
                    f"Popis v titulku je příliš stručný („{description}“)."
                )

        if errors:
            return CheckResult(
                False,
                "V titulcích objektů chybí popis:\n"
                + "\n".join(f"– {e}" for e in errors),
                self.penalty * len(errors),
            )

        return CheckResult(
            True,
            "Všechny titulky obsahují stručný popis objektu.",
            0,
        )