from pathlib import Path
import pytest

from main2 import run_checks

DOCS_DIR = Path(__file__).resolve().parents[1] / "docs"


def docx_files():
    files = sorted(DOCS_DIR.glob("*.docx"))
    if not files:
        pytest.skip(
            "Ve složce docs/ nejsou žádné .docx soubory",
            allow_module_level=True
        )
    return files

@pytest.mark.parametrize("word_path", docx_files(), ids=lambda p: p.name)
def test_word_docs_smoke(word_path: Path):
    """
    Smoke test pro Word dokumenty
    """
    report = run_checks(
        path=word_path,
        run_excel=False,
    )
    assert report is not None


def spreadsheet_files():
    files = sorted(
        list(DOCS_DIR.glob("*.xlsx")) +
        list(DOCS_DIR.glob("*.ods"))
    )
    if not files:
        pytest.skip(
            "Ve složce docs/ nejsou žádné .xlsx ani .ods soubory",
            allow_module_level=True
        )
    return files


@pytest.mark.parametrize("sheet_path", spreadsheet_files(), ids=lambda p: p.name)
def test_spreadsheet_docs_smoke(sheet_path: Path):
    """
    Smoke test pro Excel i LibreOffice Calc dokumenty
    """
    report = run_checks(
        path=sheet_path,
        run_excel=True,
        excel_assignment_path="assignment/excel/assignment.json",
    )
    assert report is not None