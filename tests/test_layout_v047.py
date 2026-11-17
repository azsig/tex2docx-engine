r"""Unit test perbaikan v0.4.7 — stdlib murni, tanpa engine.

Empat laporan lanjutan dari Naskah_LaTeX.docx v0.4.6 + permintaan font
murni dari LaTeX:
1. drawing_above lama memakai prev_nonempty: MELOMPATI paragraf gambar
   murni (teksnya kosong) dan bisa mendarat di logo halaman judul
   (byline teks bukan tembok) — halaman judul terpotong setelah SKRIPSI;
2. fo:font-size junk ooffice.4ht (17.3pt) diaktifkan v0.4.6 -> bab
   17.5pt; solusi umum: probe NFSS \f@size/\f@series/\f@shape dari
   \@makechapterhead & \@startsection (t2dcfg) + fallback fix_odt;
3+4. persamaan 2 kolom 50/50 dan persamaan dalam \item bukan tabel ->
     semua dinormalisasi jadi tabel 3 kolom di polish;
5. 900+ <text:span> TFM (bold/italic/ukuran LaTeX) diabaikan LibreOffice
   karena style-nya tak pernah didefinisikan ooffice.4ht.
"""
import re
import zipfile
from xml.etree import ElementTree as ET

import tex2docx_engine as pkg  # via tests/conftest.py

from test_layout_v046 import (EMU_CM, FO_NS, ODT_NS, STYLE_NS, _docx_drawing,
                              _odt, _read_odt)

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
M_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'


# ---------- 1. drawing_above versi baru ----------

def _fake_segs(raws, texts):
    segs = []
    for raw in raws:
        segs.append(['text', ''])
        segs.append(['para', raw])
    idxs = [i for i, (k, _) in enumerate(segs) if k == 'para']
    return segs, idxs, {i: texts[i // 2] for i in idxs}


def test_drawing_above_gambar_murni_langsung_di_atas_caption():
    r"""Kasus inti Gambar 4.4: paragraf gambar tak ber-TEKS — prev_nonempty
    lama melompatinya sehingga break tak pernah naik ke gambar."""
    segs, idxs, texts = _fake_segs(
        ['<w:p><w:drawing><wp:x/></w:p>',
         '<w:p><w:t>Gambar 4.4 Kurva</w:t></w:p>'],
        ['', 'gambar44kurva'])
    assert pkg.graft._drawing_above(1, segs, idxs, texts) == 0


def test_drawing_above_byline_jadi_tembok_cover():
    r"""Halaman judul: break PERNYATAAN tak boleh naik ke logo cover —
    paragraf teks byline di antara menghentikan pencarian."""
    segs, idxs, texts = _fake_segs(
        ['<w:p><w:drawing><wp:x/></w:p>',      # logo cover
         '<w:p><w:t>diajukan oleh YASIR ABDULAZIZ</w:t></w:p>',
         '<w:p/>', '<w:p/>', '<w:p/>',
         '<w:p><w:t>PERNYATAAN BEBAS PLAGIASI</w:t></w:p>'],
        ['', 'diajukanolehyasirabdulaziz', '', '', '', 'pernyataanbebasplagiasi'])
    assert pkg.graft._drawing_above(5, segs, idxs, texts) is None


def test_drawing_above_paragraf_kosong_dilompati_dan_teks_panjang_menolak():
    segs, idxs, texts = _fake_segs(
        ['<w:p><w:drawing><wp:x/></w:p>', '<w:p/>', '<w:p/>',
         '<w:p><w:t>Gambar 5.1</w:t></w:p>'],
        ['', '', '', 'gambar51'])
    assert pkg.graft._drawing_above(3, segs, idxs, texts) == 0
    # gambar bawa teks panjang -> bukan paragraf gambar murni, tolak
    segs2, idxs2, texts2 = _fake_segs(
        ['<w:p><w:drawing><wp:x/></w:p>', '<w:p><w:t>teks</w:t></w:p>'],
        ['x' * 50, 'tekssetelah'])
    assert pkg.graft._drawing_above(1, segs2, idxs2, texts2) is None


# ---------- 2. fallback font heading + span TFM ----------

def test_tfm_span_styles_dibuatkan(tmp_path):
    r"""900+ span TFM ooffice.4ht (bold/italic/ukuran LaTeX) diabaikan
    LibreOffice karena style-nya tak terdefinisi — dibuatkan otomatis."""
    content = (f'<office:document-content xmlns:office="{ODT_NS["office"]}" '
               f'xmlns:text="{ODT_NS["text"]}" xmlns:style="{STYLE_NS}">'
               '<office:body><office:text>'
               '<text:p><text:span text:style-name="ptmb8t-x-x-144">'
               'JUDUL SKRIPSI</text:span></text:p>'
               '<text:p><text:span text:style-name="ptmbi8t-x-x-144">'
               'DYNAMIC CHANNEL ATTENTION</text:span></text:p>'
               '<text:p><text:span text:style-name="ptmr8t-x-x-80">'
               'catatan kecil</text:span></text:p>'
               '<text:p><text:span text:style-name="BukanTFM">x</text:span>'
               '</text:p>'
               '</office:text></office:body></office:document-content>')
    dst = str(tmp_path / 'out.odt')
    pkg.fix_odt.run(_odt(tmp_path, content), dst)
    root = ET.fromstring(_read_odt(dst))
    made = {st.get(f'{{{STYLE_NS}}}name'): st
            for st in root.iter(f'{{{STYLE_NS}}}style')}
    b = made.get('ptmb8t-x-x-144')
    assert b is not None
    tp = b.find(f'{{{STYLE_NS}}}text-properties')
    assert tp.get(f'{{{FO_NS}}}font-weight') == 'bold'
    assert tp.get(f'{{{FO_NS}}}font-size') == '14.4pt'
    bi = made['ptmbi8t-x-x-144'].find(f'{{{STYLE_NS}}}text-properties')
    assert bi.get(f'{{{FO_NS}}}font-weight') == 'bold'
    assert bi.get(f'{{{FO_NS}}}font-style') == 'italic'
    sm = made['ptmr8t-x-x-80'].find(f'{{{STYLE_NS}}}text-properties')
    assert sm.get(f'{{{FO_NS}}}font-size') == '8pt'
    assert 'BukanTFM' not in made           # nama non-TFM dibiarkan


def test_tfm_ec_dan_cm_size():
    for nm, sz in (('ecrm1200', '12pt'), ('ecti1000', '10pt'),
                   ('cmbx12', '12pt'), ('lmr12', '12pt')):
        m = pkg.fix_odt.RE_TFM_SPAN.match(nm)
        assert m, nm
        fam, variant, d1, d2 = m.groups()
        assert (d2 or d1) is not None
        val = int(d2) / 10 if d2 else (
            int(d1) / 100 if int(d1) >= 100 else float(d1))
        assert f'{val:g}pt' == sz, nm


# ---------- 3+4. normalisasi persamaan 3 kolom di polish ----------

EQ_2COL = (
    '<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/>'
    '<w:jc w:val="start"/><w:tblLayout w:type="fixed"/></w:tblPr>'
    '<w:tblGrid><w:gridCol w:w="4534"/><w:gridCol w:w="4536"/></w:tblGrid>'
    '<w:tr>'
    '<w:tc><w:p><w:pPr><w:pStyle w:val="Equp"/></w:pPr>'
    f'<w:r><m:oMath xmlns:m="{M_NS}"><m:r><m:t>a+b=c</m:t></m:r>'
    '</m:oMath></w:r></w:p></w:tc>'
    '<w:tc><w:p><w:pPr><w:pStyle w:val="Equnump"/></w:pPr>'
    '<w:r><w:t>(3.1)</w:t></w:r></w:p></w:tc>'
    '</w:tr></w:tbl>')


def test_persamaan_2kolom_jadi_3kolom(tmp_path):
    xml = (f'<?xml version="1.0"?><w:document xmlns:w="{W_NS}" '
           f'xmlns:wp="x" xmlns:m="{M_NS}"><w:body>'
           f'<w:p><w:r><w:t>tanda satu-satunya paragraf biasa.</w:t></w:r></w:p>'
           + EQ_2COL + '</w:body></w:document>')
    src = tmp_path / 'in.docx'
    with zipfile.ZipFile(src, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('word/document.xml', xml)
        z.writestr('word/_rels/document.xml.rels',
                   '<?xml version="1.0"?><Relationships/>')
        z.writestr('word/styles.xml', '<?xml version="1.0"?><w:styles/>')
    dst = str(tmp_path / 'out.docx')
    pkg.polish.run(str(src), dst)
    doc = zipfile.ZipFile(dst).read('word/document.xml').decode('utf-8')
    assert '<w:gridCol w:w="1088"/><w:gridCol w:w="6894"/>' in doc
    assert '<w:jc w:val="end"/>' in doc          # sel nomor rata kanan
    assert '(3.1)' in doc and 'a+b=c' in doc     # isi selamat
    assert doc.count('<w:tbl>') == 1


def test_persamaan_ratakan_list_dirangkum_tabel(tmp_path):
    xml = (f'<?xml version="1.0"?><w:document xmlns:w="{W_NS}" '
           f'xmlns:m="{M_NS}"><w:body>'
           # hasil ratakan fix_odt yang diekspor LibreOffice
           '<w:p><w:pPr><w:spacing w:before="100"/></w:pPr>'
           f'<m:oMath xmlns:m="{M_NS}"><m:r><m:t>x=y</m:t></m:r></m:oMath>'
           '<w:r><w:tab/></w:r><w:r><w:t>(3.45)</w:t></w:r></w:p>'
           # matematika INLINE ber-teks sesudahnya: TIDAK boleh dirangkum
           '<w:p><w:r><m:oMath xmlns:m="' + M_NS + '"><m:r><m:t>k</m:t></m:r>'
           '</m:oMath></w:r><w:r><w:t> adalah konstanta dan sebagainya.</w:t>'
           '</w:r></w:p>'
           '</w:body></w:document>')
    src = tmp_path / 'in.docx'
    with zipfile.ZipFile(src, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('word/document.xml', xml)
        z.writestr('word/_rels/document.xml.rels',
                   '<?xml version="1.0"?><Relationships/>')
        z.writestr('word/styles.xml', '<?xml version="1.0"?><w:styles/>')
    dst = str(tmp_path / 'out.docx')
    pkg.polish.run(str(src), dst)
    doc = zipfile.ZipFile(dst).read('word/document.xml').decode('utf-8')
    assert '<w:gridCol w:w="6894"/>' in doc      # persamaan display dirangkum
    assert '(3.45)' in doc
    # paragraf inline tetap paragraf — ekornya bukan tab+nomor persamaan
    assert doc.count('<w:tbl>') == 1
    assert 'adalah konstanta' in doc


# ---------- typography: OMML bukan prosa ----------

def test_paragraf_omml_tak_ditata(tmp_path):
    long_math = 'satu dua tiga empat lima enam tujuh delapan sembilan sepuluh'
    para = (f'<w:p><w:r><m:oMath xmlns:m="{M_NS}"><m:r>'
            f'<m:t>{long_math}</m:t></m:r></m:oMath></w:r></w:p>')
    src = tmp_path / 'in.docx'
    with zipfile.ZipFile(src, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('word/document.xml',
                   f'<?xml version="1.0"?><w:document xmlns:w="{W_NS}" '
                   f'xmlns:m="{M_NS}"><w:body>'
                   '<w:p><w:r><w:t>BAB I PENDAHULUAN</w:t></w:r></w:p>'
                   + para + '</w:body></w:document>')
    dst = str(tmp_path / 'out.docx')
    pkg.typography.run(str(src), dst, str(tmp_path), 'x')
    paras = re.findall(r'<w:p\b[^>]*?(?:/>|>.*?</w:p>)',
                       zipfile.ZipFile(dst).read('word/document.xml')
                       .decode('utf-8'), re.S)
    eqp = next(p for p in paras if 'oMath' in p)
    assert 'w:line="360"' not in eqp and 'w:firstLine' not in eqp


# ---------- polish: drawing besar tetap clamp 13.5cm (regresi) ----------

def test_drawing_besar_non_logo_tetap_clamp(tmp_path):
    cx, cy = int(30 * EMU_CM), int(22.5 * EMU_CM)
    src = _docx_drawing(tmp_path, 'alat_annotated.jpg', cx, cy)
    dst = str(tmp_path / 'out.docx')
    pkg.polish.run(src, dst)
    doc = zipfile.ZipFile(dst).read('word/document.xml').decode('utf-8')
    assert f'cx="{int(13.5 * EMU_CM)}"' in doc
