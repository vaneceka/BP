from checks.base_check import BaseCheck, CheckResult
from PIL import Image
import io


class ImageLowQualityCheck(BaseCheck):
    name = "Chybí obrázek nebo má nízké rozlišení"
    penalty = -100

    MIN_WIDTH = 300
    MIN_HEIGHT = 300

    def run(self, document, assignment=None):
        images = list(document.iter_images())

        if not images:
            return CheckResult(
                False,
                "Není vložen žádný obrázek.",
                self.penalty,
            )

        for img in images:
            img_bytes = document.get_image_bytes(img["rId"])
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

        return CheckResult(
            True,
            "Všechny obrázky mají dostatečné rozlišení.",
            0,
        )