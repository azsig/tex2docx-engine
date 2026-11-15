r"""Unit test perbaikan v0.4.6 — stdlib murni, tanpa engine.

Enam laporan dari Naskah_LaTeX.docx hasil v0.4.5:
1. persamaan bernomor di dalam \item hilang (table:table di list-item —
   ODF melarang, LibreOffice membuang);
2a. gambar isi resolusi tinggi menyusut jadi ukuran logo (syarat
    `or cx > 25cm` di polish);
2b. page-break menempel di CAPTION di bawah gambar, bukan paragraf gambar;
3. vcur membeku di DAFTAR ISI (63 entri tanpa celah >= 6pt) lalu entri
   daftar tertimpa w:before judul bab — regresi jendela v0.4.5;
4. $\\times$ dsb. jadi eqArr 3 kotak kosong (mrow kosong + mo tunggal);
5. Heading2/3 ikut list_para (judul dibuat numPr oleh make4ht);
6. Heading1..4 tak bold (atribut font di paragraph-properties, bukan
   text-properties — LibreOffice mengabaikannya).
"""
import re
import zipfile

import tex2docx_engine as pkg  # via tests/conftest.py

EMU_CM = 360000
FO_NS = 'http://www.w3.org/1999/XSL/Format'
STYLE_NS = 'urn:oasis:names:tc:opendocument:xmlns:style:1.0'


# ---------- 1. persamaan di dalam list-item ----------

ODT_NS = {
    'office': 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
    'text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
    'table': 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
    'draw': 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0',
    'style': STYLE_NS,
}


def _odt(tmp_path, content):
    src = tmp_path / 'in.odt'
    with zipfile.ZipFile(src, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('content.xml', content)
        z.writestr('styles.xml',
                   f"<office:document-styles xmlns:office=\"{ODT_NS['office']}\" "
                   f"xmlns:style=\"{STYLE_NS}\"/>")
    return str(src)


def _read_odt(path, name='content.xml'):
    return zipfile.ZipFile(path).read(name).decode('utf-8')


def test_persamaan_table_dalam_list_item_diratakan(tmp_path):
    content = f"""<office:document-content { " ".join(f'xmlns:{k}="{v}"' for k, v in ODT_NS.items()) }>
<office:body><office:text>
<text:list><text:list-item><text:p>Item enumerate</text:p>
<table:table><table:table-row>
<table:table-cell><text:p><draw:frame draw:name="rumus"/></text:p></table:table-cell>
<table:table-cell><text:p>(3.45)</text:p></table:table-cell>
</table:table-row></table:table>
</text:list-item></text:list>
</office:text></office:body></office:document-content>"""
    dst = str(tmp_path / 'out.odt')
    pkg.fix_odt.run(_odt(tmp_path, content), dst)
    out = _read_odt(dst)
    assert '<table:table' not in out          # tabel di list-item lenyap
    assert '<draw:frame' in out               # frame rumus selamat
    assert '(3.45)' in out                    # nomor persamaan selamat
    assert '<text:tab' in out                 # pemisah rumus | nomor
    # hasilnya paragraf, bukan lagi table
    assert re.search(r'<text:list-item>\s*<text:p>Item enumerate</text:p>\s*'
                     r'<text:p>', out)


# ---------- 2a. polish: logo hanya lewat nama ----------

def _docx_drawing(tmp_path, name, cx, cy, fname='in.docx'):
    xml = (f'<?xml version="1.0"?><w:document xmlns:w="http://schemas.'
           f'openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="x">'
           f'<w:body><w:p><w:r><w:drawing><wp:extent cx="{cx}" cy="{cy}"/>'
           f'<wp:docPr id="1" name="{name}"/></w:drawing></w:r></w:p>'
           f'</w:body></w:document>')
    path = tmp_path / fname
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('word/document.xml', xml)
        z.writestr('word/_rels/document.xml.rels',
                   '<?xml version="1.0"?><Relationships/>')
        z.writestr('word/styles.xml', '<?xml version="1.0"?><w:styles/>')
    return str(path)


def test_gambar_besar_non_logo_tak_menyusut_jadi_logo(tmp_path):
    r"""Gambar isi cx>25cm rasio 0.5-2.0 dulu dikira logo (syarat
    `or cx > 25 * EMU_CM`) dan menyusut ke 5.5cm — kini clamp ke batas
    halaman (13.5cm) saja."""
    cx, cy = int(30 * EMU_CM), int(22.5 * EMU_CM)      # rasio 0.75
    src = _docx_drawing(tmp_path, 'Gambar_4_4.png', cx, cy)
    dst = str(tmp_path / 'out.docx')
    pkg.polish.run(src, dst)
    doc = zipfile.ZipFile(dst).read('word/document.xml').decode('utf-8')
    want_w = int(13.5 * EMU_CM)
    assert f'cx="{want_w}"' in doc                     # clamp 13.5cm, bukan logo
    assert f'cx="{int(5.5 * EMU_CM)}"' not in doc


def test_logo_tetap_menyusut_dari_namanya(tmp_path):
    cx, cy = int(14.6 * EMU_CM), int(15.2 * EMU_CM)    # 'persegi raksasa'
    src = _docx_drawing(tmp_path, 'logoUGM.png', cx, cy)
    dst = str(tmp_path / 'out.docx')
    pkg.polish.run(src, dst)
    doc = zipfile.ZipFile(dst).read('word/document.xml').decode('utf-8')
    assert f'cx="{int(5.5 * EMU_CM)}"' in doc


# ---------- 2b/3. graft: unit jendela & penanda gambar ----------

def test_jendela_default_120():
    idxs = list(range(300))
    texts = {i: f'para{i}' for i in idxs}
    texts[299] = 'targetunik'                          # jauh di luar jendela
    assert pkg.graft._find_line_target('targetunik', texts, idxs, 0) is None


# ---------- 4. MathML mrow kosong / mo tunggal ----------

def test_mathml_mrow_kosong_dan_mo_tunggal():
    xml = ('<math xmlns="http://www.w3.org/1998/Math/MathML">'
           '<mrow></mrow><mo stretchy="false">×</mo><mrow></mrow></math>')
    out, n_mrow, n_mo = pkg.fix_odt._clean_mathml(xml)
    assert n_mrow == 2 and n_mo == 1
    assert '<mrow' not in out
    assert '<mi>×</mi>' in out


def test_mathml_rumus_sehat_tak_tersentuh():
    xml = ('<math xmlns="http://www.w3.org/1998/Math/MathML"><mrow><mi>a</mi>'
           '</mrow><mo>+</mo><mrow><mi>b</mi></mrow></math>')
    out, n_mrow, n_mo = pkg.fix_odt._clean_mathml(xml)
    assert (n_mrow, n_mo) == (0, 0) and out == xml


def test_mathml_dibersihkan_di_formula_embed(tmp_path):
    r"""Objek rumus ODT ter-embed di *-m*/content.xml — semua member
    berakhiran content.xml ikut dibersihkan."""
    content = ('<office:document-content '
               'xmlns:office="%s" xmlns:text="%s">'
               % (ODT_NS['office'], ODT_NS['text'])
               + '<office:body><office:text><text:p>x</text:p>'
                 '</office:text></office:body></office:document-content>')
    mathml = ('<math xmlns="http://www.w3.org/1998/Math/MathML">'
              '<mrow></mrow><mo>⊙</mo><mrow></mrow></math>')
    src = tmp_path / 'in.odt'
    with zipfile.ZipFile(src, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('content.xml', content)
        z.writestr('styles.xml',
                   '<office:document-styles xmlns:office="%s" '
                   'xmlns:style="%s"/>' % (ODT_NS['office'], STYLE_NS))
        z.writestr('skripsi-m1/content.xml', mathml)
    dst = str(tmp_path / 'out.odt')
    pkg.fix_odt.run(str(src), dst)
    assert '<mi>⊙</mi>' in _read_odt(dst, 'skripsi-m1/content.xml')


# ---------- ukuran gambar dari LaTeX (marker t2dimg) ----------

def test_t2dimg_width_diterapkan_dengan_rasio_frame():
    cm = ('<!--t2dimg:hasw=1,w=156.49014pt,hash=0,h=0.0pt,scale=1-->'
          '<draw:frame svg:width="414pt" svg:height="430pt">')
    out, n = pkg.fix_odt._apply_latex_image_sizes('x' + cm + 'y')
    assert n == 1
    assert 'svg:width="156.49pt"' in out
    assert f'svg:height="{156.49014 * 430 / 414:.2f}pt"' in out
    assert 't2dimg' not in out     # marker terkonsumsi bersama frame-nya


def test_t2dimg_scale_dan_tanpa_marker():
    sc = ('<!--t2dimg:hasw=0,w=0.0pt,hash=0,h=0.0pt,scale=2-->'
          '<draw:frame svg:width="100pt" svg:height="50pt">')
    out, n = pkg.fix_odt._apply_latex_image_sizes(sc)
    assert n == 1 and 'svg:width="150.00pt"' in out and \
        'svg:height="75.00pt"' in out
    plain = '<draw:frame svg:width="100pt" svg:height="50pt">'
    assert pkg.fix_odt._apply_latex_image_sizes(plain) == (plain, 0)


# ---------- 6. font Heading pindah ke text-properties ----------

def test_heading_font_attrs_pindah_ke_text_properties():
    styles = (f'<?xml version="1.0"?><office:document-styles '
              f'xmlns:office="{ODT_NS["office"]}" xmlns:style="{STYLE_NS}" '
              f'xmlns:fo="{FO_NS}">'
              '<style:style style:name="Heading-2" style:family="paragraph">'
              '<style:paragraph-properties fo:font-weight="bold" '
              'fo:font-size="12pt" fo:margin-bottom="25pt"/>'
              '</style:style>'
              '<style:style style:name="Standard" style:family="paragraph">'
              '<style:paragraph-properties fo:font-weight="bold"/>'
              '</style:style>'
              '</office:document-styles>').encode('utf-8')
    out, n = pkg.fix_odt._fix_heading_fonts(styles)
    assert n == 2                                    # hanya Heading-2
    root = __import__('xml.etree.ElementTree', fromlist=['Element']).fromstring(out)
    st = next(s for s in root.iter(f'{{{STYLE_NS}}}style')
              if s.get(f'{{{STYLE_NS}}}name') == 'Heading-2')
    pp = st.find(f'{{{STYLE_NS}}}paragraph-properties')
    tp = st.find(f'{{{STYLE_NS}}}text-properties')
    assert tp.get(f'{{{FO_NS}}}font-weight') == 'bold'
    assert tp.get(f'{{{FO_NS}}}font-size') == '12pt'
    assert f'{{{FO_NS}}}margin-bottom' in pp.attrib
    assert all(k.rsplit('}', 1)[-1] not in ('font-weight', 'font-size')
               for k in pp.attrib)
    # style non-heading tak tersentuh
    std = next(s for s in root.iter(f'{{{STYLE_NS}}}style')
               if s.get(f'{{{STYLE_NS}}}name') == 'Standard')
    assert std.find(f'{{{STYLE_NS}}}text-properties') is None


# ---------- 5. Heading tak ikut list_para di typography ----------

def test_judul_heading_tak_tertimpa_line15_justify(tmp_path):
    heading = ('<w:p><w:pPr><w:pStyle w:val="Heading2"/>'
               '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'
               '</w:pPr><w:r><w:t>Judul sub bab yang cukup panjang untuk '
               'melewati ambang enam puluh karakter pemeriksa format</w:t>'
               '</w:r></w:p>')
    body = ('<w:p><w:r><w:t>Paragraf badan panjang lebih dari enam puluh '
            'karakter supaya ikut ditata sebagai prosa isi bab.</w:t></w:r>'
            '</w:p>')
    src = tmp_path / 'in.docx'
    with zipfile.ZipFile(src, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('word/document.xml',
                   '<?xml version="1.0"?><w:document xmlns:w="http://schemas.'
                   'openxmlformats.org/wordprocessingml/2006/main"><w:body>'
                   '<w:p><w:r><w:t>BAB I PENDAHULUAN</w:t></w:r></w:p>'
                   + heading + body + '</w:body></w:document>')
    dst = str(tmp_path / 'out.docx')
    pkg.typography.run(str(src), dst, str(tmp_path), 'x')
    paras = re.findall(r'<w:p\b[^>]*?(?:/>|>.*?</w:p>)',
                       zipfile.ZipFile(dst).read('word/document.xml')
                       .decode('utf-8'), re.S)
    head = next(p for p in paras if 'Heading2' in p)
    assert 'w:line="360"' not in head and '<w:jc' not in head
    bodyp = next(p for p in paras if 'ikut ditata' in p)
    assert 'w:line="360"' in bodyp
