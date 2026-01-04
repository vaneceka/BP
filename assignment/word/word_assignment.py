from dataclasses import dataclass
from typing import Dict, List
from .word_assignment_model import StyleSpec


@dataclass
class Assignment:
    styles: Dict[str, StyleSpec]
    headlines: List[dict]
    objects: List[dict]
    bibliography: List[dict]