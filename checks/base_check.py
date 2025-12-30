from abc import ABC, abstractmethod
from document.excel_document import ExcelDocument
from document.word_document import WordDocument
from assignment.word.word_assignment import Assignment

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
            #  document: WordDocument | ExcelDocument,
            document,
            assignment: Assignment | None = None) -> CheckResult:
        pass