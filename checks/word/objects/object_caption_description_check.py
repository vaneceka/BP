from checks.base_check import BaseCheck, CheckResult
import re


class ObjectCaptionDescriptionCheck(BaseCheck):
    name = "Titulek objektu neodpovídá zadání"
    penalty = -5

    def _norm(self, text: str) -> str:
        return re.sub(r"\s+", " ", text.strip()).lower()

    def run(self, document, assignment=None):

        if assignment is None or not hasattr(assignment, "objects"):
            return CheckResult(True, "Chybí assignment", 0)

        errors = []

        doc_images = []
        for obj in document.iter_objects():
            if obj["type"] == "image":
                doc_images.append(obj)

        expected_images = []
        for obj in assignment.objects:
            if obj["type"] == "image":
                expected_images.append(obj)

        if len(doc_images) < len(expected_images):
            errors.append(
                f"V dokumentu je méně obrázků ({len(doc_images)}) než v zadání ({len(expected_images)})."
            )

        for i, expected in enumerate(expected_images):

            if i >= len(doc_images):
                break

            element = doc_images[i]["element"]
            expected_caption = expected["caption"]

            # najdi titulek
            caption_p = None

            before = document.paragraph_before(element)
            if before and document.paragraph_has_seq_caption(before):
                caption_p = before

            after = document.paragraph_after(element)
            if caption_p is None and after and document.paragraph_has_seq_caption(after):
                caption_p = after

            if caption_p is None:
                errors.append(
                    f"Obrázek č. {i+1}: chybí titulek."
                )
                continue

            actual_text = document._paragraph_text(caption_p)

            if not actual_text:
                errors.append(
                    f"Obrázek č. {i+1}: titulek je prázdný."
                )
                continue

            actual_text = re.sub(
                r"^(Obrázek|Tabulka|Graf)\s+\d+\s*[:\-]?\s*",
                "",
                actual_text,
                flags=re.IGNORECASE,
            ).strip()

            if self._norm(actual_text) != self._norm(expected_caption):
                errors.append(
                    f"Obrázek č. {i+1}:\n"
                    f"  očekáváno: „{expected_caption}“\n"
                    f"  nalezeno:  „{actual_text}“"
                )

        if errors:
            return CheckResult(
                False,
                "Chyby v titulcích obrázků:\n"
                + "\n".join("– " + e for e in errors),
                self.penalty * len(errors),
            )

        return CheckResult(
            True,
            "Všechny titulky obrázků odpovídají zadání.",
            0,
        )