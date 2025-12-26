from abc import ABC, abstractmethod

class CheckResult:
    def __init__(self, passed: bool, message: str, points: int, fatal: bool = False):
        self.passed = passed
        self.message = message
        self.points = points
        self.fatal = fatal


class BaseCheck(ABC):
    name: str
    penalty: int

    @abstractmethod
    def run(self, document, assignment=None) -> CheckResult:
        pass