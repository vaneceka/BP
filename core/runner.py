class Runner:
    def __init__(self, checks):
        self.checks = checks

    def run(self, document, assignment):
        results = []
        for check in self.checks:
            results.append((check, check.run(document, assignment)))
        return results