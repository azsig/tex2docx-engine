"""Unit test gate penerimaan make4ht (_make4ht_odt) — stdlib murni, tanpa
engine eksternal. Gate menilai ODT HASIL, bukan teks log: 'Unbalanced Tag'
dimaafkan hanya bila content.xml lolos parse strict (domfilter make4ht
menyembuhkannya) DAN kontennya tak pendek vs oracle PDF."""
import sys
import zipfile

sys.path.insert(0, str(__import__('pathlib').Path(__file__).resolve().parents[1] / 'src'))
from tex2docx_engine import pipeline  # noqa: E402

HEAD = ('<?xml version="1.0" encoding="utf-8"?>'
        '<office:document-content '
        'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" '
        'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">')
TAIL = '</office:text></office:body></office:document-content>'


def _odt(path, content):
    with zipfile.ZipFile(path, 'w') as z:
        z.writestr('content.xml', content)


def _body(txt):
    return (HEAD + '<office:body><office:text>'
            f'<text:p>{txt}</text:p>'
            + TAIL)


def test_content_parses_valid(tmp_path):
    p = tmp_path / 'ok.odt'
    _odt(str(p), _body('isi'))
    assert pipeline._odt_content_parses(str(p)) is True


def test_content_parses_tolak_xml_rusak(tmp_path):
    p = tmp_path / 'bad.odt'
    _odt(str(p), _body('isi') + '</text:p>')      # tag tutup liar
    assert pipeline._odt_content_parses(str(p)) is False


def test_content_parses_tolak_zip_rusak(tmp_path):
    p = tmp_path / 'nope.odt'
    p.write_bytes(b'PHP')
    assert pipeline._odt_content_parses(str(p)) is False


def test_complete_tanpa_oracle_selalu_true(tmp_path):
    p = tmp_path / 'a.odt'
    _odt(str(p), _body('seadanya'))
    assert pipeline._odt_complete(str(p), None) is True
    assert pipeline._odt_complete(str(p), 0) is True      # PDF tanpa teks
    assert pipeline._odt_complete(str(p), 150) is True    # < 200 chars: skip


def test_complete_ratio_threshold(tmp_path):
    p = tmp_path / 'a.odt'
    _odt(str(p), _body('x' * 300))                        # 300 chars konten
    assert pipeline._odt_complete(str(p), 300) is True    # 300 >= 180
    assert pipeline._odt_complete(str(p), 1000) is False  # 300 < 600 (0.3)


def test_odt_tail_paragraf_non_tabel_terakhir(tmp_path):
    """Fingerprint ekor = paragraf NON-TABEL terakhir: teks tabel tak stabil
    urutan konkatenasinya antar format; paragraf flow dijamin kontigu."""
    p = tmp_path / 'a.odt'
    _odt(str(p), (HEAD + '<office:body><office:text>'
                  '<text:p>satu</text:p>'
                  '<table:table xmlns:table="urn:oasis:names:tc:opendocument'
                  ' xmlns:table:1.0"><table:table-row><table:table-cell>'
                  '<text:p>sel tabel</text:p>'
                  '</table:table-cell></table:table-row></table:table>'
                  '<text:p>paragraf akhir dokumen</text:p>'
                  + TAIL))
    assert pipeline._odt_tail(str(p)) == 'paragrafakhirdokumen'


def test_bytes_text_buang_tag_ws_dan_entitas():
    assert pipeline._bytes_text(b'<p>It&#160;works.</p>\n  <p>A B</p>') \
        == 'Itworks.AB'
