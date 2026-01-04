from pathlib import Path
import pytest

from main2 import run_checks  # uprav, pokud se soubor nejmenuje main.py

DOCS_DIR = Path(__file__).resolve().parents[1] / "docs"

def docx_files():
    files = sorted(DOCS_DIR.glob("*.docx"))
    if not files:
        pytest.skip("Ve složce docs/ nejsou žádné .docx soubory")
    return files

@pytest.mark.parametrize("word_path", docx_files(), ids=lambda p: p.name)
def test_docs_smoke(word_path: Path):
    # smoke: nesmí to spadnout, a coverage se nasbírá
    report = run_checks(word_path, run_excel=False)
    assert report is not None