import json

from .word_assignment_model import StyleSpec, WordAssignment


def load_assignment(path: str) -> WordAssignment:
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    styles = {}

    for name, spec in data.get("styles", {}).items():
        styles[name] = StyleSpec(
            name=name,
            font=spec.get("type"),
            size=spec.get("size"),
            bold=spec.get("bold"),
            italic=spec.get("italic"),
            allCaps=spec.get("allCaps"),
            alignment=spec.get("alignment"),
            color=spec.get("color"),
            lineHeight=spec.get("lineHeight"),
            pageBreakBefore=spec.get("pageBreakBefore"),
            numLevel=spec.get("numLevel"),
            basedOn=spec.get("basedOn"),
            spaceBefore=spec.get("spaceBefore"),
            tabs=[(t[0], int(t[1])) for t in spec.get("tabs", [])] or None
        )
    
    headlines = data.get("headlines", [])
    objects = data.get("objects", [])
    bibliography = data.get("bibliography", [])

    return WordAssignment(
        styles=styles,
        headlines=headlines,
        objects=objects,
        bibliography=bibliography,
    )