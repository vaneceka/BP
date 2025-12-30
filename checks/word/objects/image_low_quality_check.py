from checks.word.base_check import BaseCheck, CheckResult
from PIL import Image
import io


class ImageLowQualityCheck(BaseCheck):
    name = "Chybí obrázek nebo má nízké rozlišení"
    penalty = -100

    MIN_WIDTH = 300
    MIN_HEIGHT = 300

    def run(self, document, assignment=None):
        image_objects = [o for o in document.iter_objects() if o["type"] == "image"]

        if not image_objects:
            return CheckResult(False, "Není vložen žádný obrázek.", self.penalty)

        for obj in image_objects:
            element = obj["element"]

            # v jednom odstavci může být i víc obrázků
            for rid in document.object_image_rids(element):
                img_bytes = document.get_image_bytes(rid)
                if not img_bytes:
                    continue

                try:
                    image = Image.open(io.BytesIO(img_bytes))
                    width, height = image.size
                except Exception:
                    continue

                if width < self.MIN_WIDTH or height < self.MIN_HEIGHT:
                    return CheckResult(
                        False,
                        f"Obrázek má nízké rozlišení ({width}×{height} px).",
                        self.penalty,
                    )

        return CheckResult(True, "Všechny obrázky mají dostatečné rozlišení.", 0)