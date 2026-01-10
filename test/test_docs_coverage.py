from pathlib import Path
import pytest

from main2 import run_checks

DOCS_DIR = Path(__file__).resolve().parents[1] / "docs"


def docx_files():
    files = sorted(DOCS_DIR.glob("*.docx"))
    if not files:
        pytest.skip("Ve složce docs/ nejsou žádné .docx soubory")
    return files


def xlsx_files():
    files = sorted(DOCS_DIR.glob("*.xlsx"))
    if not files:
        pytest.skip("Ve složce docs/ nejsou žádné .xlsx soubory")
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


@pytest.mark.parametrize("excel_path", xlsx_files(), ids=lambda p: p.name)
def test_excel_docs_smoke(excel_path: Path):
    """
    Smoke test pro Excel dokumenty
    """
    report = run_checks(
        path=excel_path,
        run_excel=True,
        excel_assignment_path="assignment/excel/assignment.json",
    )
    assert report is not None