"""Integration test: full LaTeX -> DOCX run on a small fixture.
Skipped automatically when the external engines (make4ht/soffice/pdftotext)
are not installed."""
import shutil
import zipfile
from pathlib import Path

import pytest

import tex2docx_engine as pkg  # via tests/conftest.py

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


needs_core_engines = pytest.mark.skipif(
    any(shutil.which(c) is None for c in ('make4ht', 'soffice')),
    reason='butuh make4ht + LibreOffice (tanpa toolchain PDF)')


@needs_core_engines
def test_convert_no_pdf_native_only(project):
    """Mode --no-pdf: tanpa pdflatex/oracle — properti native dari marker
    tetap hadir (vskip 6pt -> w:before=120 twips), PDF tidak dibuat."""
    out = pkg.convert(str(project), main='minimal', pdf='never')
    assert Path(out).is_file()
    assert not (project / 'minimal.pdf').exists()

    doc = zipfile.ZipFile(out).read('word/document.xml').decode('utf-8')
    # \vskip 6pt (fixture) -> 6pt * 20 twips = native marker path utuh
    assert 'w:before="120"' in doc
    # scope {\centering ...} -> rata tengah native
    assert 'w:jc w:val="center"' in doc
    assert 'SMOKE TEST DOCUMENT' in doc
