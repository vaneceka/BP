from checks.base_check import BaseCheck, CheckResult


class ObjectCaptionCheck(BaseCheck):
    name = "Chybějící nebo ručně vytvořený titulek objektu"
    penalty = -5

    def run(self, document, assignment=None):
        errors = []

        expected_labels = {
            "image": "Obrázek",
            "chart": "Graf",
            "table": "Tabulka",
        }

        for obj in document.iter_objects():
            obj_type = obj["type"]
            expected = expected_labels.get(obj_type)
            element = obj["element"]

            if not expected:
                continue

            before_p = document.paragraph_before(element)
            after_p = document.paragraph_after(element)

            before_label = (
                document.paragraph_has_seq_caption(before_p)
                if before_p is not None else None
            )
            after_label = (
                document.paragraph_has_seq_caption(after_p)
                if after_p is not None else None
            )

            if obj_type == "table":
                label = before_label or after_label

                if label is None:
                    # žádný SEQ
                    has_text = False
                    if before_p and document._paragraph_text(before_p):
                        has_text = True
                    if after_p and document._paragraph_text(after_p):
                        has_text = True

                    if has_text:
                        errors.append("Tabulka má titulek vytvořený ručně.")
                    else:
                        errors.append("Tabulka nemá žádný titulek.")

                elif label != expected:
                    errors.append(
                        f"Tabulka má špatný typ návěští („{label}“)."
                    )

            else:
                if after_p is None:
                    errors.append(f"{expected} nemá žádný titulek.")
                elif after_label is None:
                    errors.append(f"{expected} má titulek vytvořený ručně.")
                elif after_label != expected:
                    errors.append(
                        f"{expected} má špatný typ návěští („{after_label}“)."
                    )

        if errors:
            return CheckResult(
                False,
                "Vložené objekty mají chybné titulky:\n"
                + "\n".join(f"– {e}" for e in errors),
                self.penalty * len(errors),
            )

        return CheckResult(True, "Všechny objekty mají správné titulky.", 0)