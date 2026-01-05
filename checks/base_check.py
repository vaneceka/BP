from abc import ABC, abstractmethod

from assignment.word.word_assignment_model import WordAssignment


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
    def run(self,
            document,
            assignment: WordAssignment | None = None) -> CheckResult:
        pass