class Report:
    def __init__(self):
        self.total_penalty = 0
        self.entries = []

    def add(self, check_name, result):
        self.entries.append((check_name, result))
        self.total_penalty += result.points

    def print(self):
        print("\n=== VÝSLEDKY HODNOCENÍ ===\n")

        for name, result in self.entries:
            status = "OK" if result.passed else "CHYBA"
            print(f"{status}: {name}")
            print(f"  {result.message}")
            print(f"  Body: {result.points}\n")

        print(f"CELKOVÁ BODOVÁ SANKCE: {self.total_penalty}")