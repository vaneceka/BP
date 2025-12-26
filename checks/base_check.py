from abc import ABC, abstractmethod
from document.word_document import WordDocument

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
    def run(self, document: WordDocument, assignment=None) -> CheckResult:
        pass