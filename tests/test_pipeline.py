"""Integration test: full LaTeX -> DOCX run on a small fixture.
Skipped automatically when the external engines (make4ht/soffice/pdftotext)
are not installed."""
import importlib.util
import shutil
import sys
import zipfile
from pathlib import Path

import pytest

SRC = Path(__file__).resolve().parents[1] / 'src'
spec = importlib.util.spec_from_file_location(
    'tex2docx_engine', SRC / 'tex2docx_engine' / '__init__.py',
    submodule_search_locations=[SRC / 'tex2docx_engine'])
pkg = importlib.util.module_from_spec(spec)
sys.modules.setdefault('tex2docx_engine', pkg)
spec.loader.exec_module(pkg)

needs_engines = pytest.mark.skipif(
    any(shutil.which(c) is None for c in ('make4ht', 'soffice', 'pdftotext', 'pdflatex')),
    reason='butuh make4ht + LibreOffice + poppler + pdflatex')


@pytest.fixture()
def project(tmp_path):
    fixture = Path(__file__).parent / 'fixtures' / 'minimal.tex'
    (tmp_path / 'minimal.tex').write_text(fixture.read_text(), encoding='utf-8')
    return tmp_path


@needs_engines
def test_convert_produces_docx_with_pdf_layout(project):
    out = pkg.convert(str(project), main='minimal')
    assert Path(out).is_file()

    doc = zipfile.ZipFile(out).read('word/document.xml').decode('utf-8')
    # oracle jarak vertikal: minimal ada satu w:before terukur (gap >= 6pt)
    assert 'w:before="' in doc
    # oracle perataan: blok center dari \begin{center} terbawa / ditemukan
    assert 'w:jc w:val="center"' in doc
    # konten utuh
    assert 'SMOKE TEST DOCUMENT' in doc
