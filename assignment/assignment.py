from dataclasses import dataclass
from typing import Dict, List, Any
from .assignment_model import StyleSpec


@dataclass
class Assignment:
    # styly
    styles: Dict[str, StyleSpec]

    # očekávané nadpisy (ze zadání)
    headlines: List[dict]

    # obrázky / tabulky / objekty
    objects: List[dict]

    # literatura
    bibliography: List[dict]