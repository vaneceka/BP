class TextDocument:
    BUILTIN_STYLE_NAMES = {
        "normal",
        "heading 1",
        "heading 2",
        "heading 3",
        "heading 4",
        "caption",
        "bibliography",
        "toc heading",
        "table of contents",
        "content heading",
    }

    def _norm(self, name: str) -> str:
        return name.strip().lower()

    def split_assignment_styles(self, assignment):
        custom = {}
        builtin = {}

        for name, spec in assignment.styles.items():
            if self._norm(name) in self.BUILTIN_STYLE_NAMES:
                builtin[name] = spec
            else:
                custom[name] = spec

        return custom, builtin