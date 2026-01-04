from checks.base_check import BaseCheck, CheckResult

class ObjectCaptionBindingCheck(BaseCheck):
    name = "Vložený objekt není s titulkem spojen"
    penalty = -1

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

            if obj_type == "table":
                caption_p = document.paragraph_before(element)
            else:
                caption_p = document.paragraph_after(element) or document.paragraph_before(element)

            if caption_p is None:
                continue 

            style_id = document._paragraph_style_id(caption_p)
            if not style_id:
                errors.append(f"{expected} není s titulkem spojen.")
                continue

            style = document._find_style_by_id(style_id)
            name_el = style.find("w:name", document.NS) if style is not None else None
            style_name = name_el.attrib.get(f"{{{document.NS['w']}}}val", "").lower() if name_el is not None else ""

            if "titulek" not in style_name and "caption" not in style_name:
                errors.append(f"{expected} není s titulkem spojen.")
                continue

            label = document.paragraph_has_seq_caption(caption_p)
            if label != expected:
                errors.append(f"{expected} není s titulkem spojen.")

        if errors:
            return CheckResult(
                False,
                "Některé objekty nejsou s titulky správně spojeny:\n"
                + "\n".join(f"– {e}" for e in errors),
                self.penalty * len(errors),
            )

        return CheckResult(
            True,
            "Všechny objekty jsou správně spojeny s titulky.",
            0,
        )