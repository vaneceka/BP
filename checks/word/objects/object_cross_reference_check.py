from checks.base_check import BaseCheck, CheckResult


class ObjectCrossReferenceCheck(BaseCheck):
    name = "Chybí křížový odkaz na objekt"
    penalty = -1

    def run(self, document, assignment=None):
        errors = []

        anchors_in_text = document.iter_crossref_anchors_in_body_text()
        for obj in document.iter_objects():
            if obj["type"] not in ("image", "chart", "table"):
                continue

            element = obj["element"]

            # titulek může být před nebo za
            caption_p = (
                document.paragraph_after(element)
                or document.paragraph_before(element)
            )

            # pokud to není titulek pokračuj
            if caption_p is None or not document.paragraph_is_caption(caption_p):
                continue

            # najdi bookmarky v titulku 
            caption_bookmarks = []
            for bm in caption_p.findall(".//w:bookmarkStart", document.NS):
                name = bm.attrib.get(f"{{{document.NS['w']}}}name")
                if name:
                    caption_bookmarks.append(name)

            # pokud titulky nemají bookmark -> chyba
            if not caption_bookmarks:
                errors.append(f"{obj['type']} není v textu zmíněn křížovým odkazem.")
                continue

            # žádný křížový odkaz v textu vůbec neexistuje
            if not anchors_in_text:
                errors.append(f"{obj['type']} není v textu zmíněn křížovým odkazem.")
                continue

            # pokud žádný bookmark z titulku není v odkazech v textu -> chyba
            if not any(bm in anchors_in_text for bm in caption_bookmarks):
                errors.append(f"{obj['type']} není v textu zmíněn křížovým odkazem.")

        if errors:
            return CheckResult(
                False,
                "V textu není prostřednictvím křížového odkazu zmíněn vložený objekt:\n"
                + "\n".join(f"– {e}" for e in errors),
                self.penalty * len(errors),
            )

        return CheckResult(True, "Všechny objekty jsou v textu zmíněny pomocí křížových odkazů.", 0)