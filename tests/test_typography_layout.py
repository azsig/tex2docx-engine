r"""Unit test tipografi & oracle jarak vertikal — stdlib murni, tanpa engine.

Tiga regresi laporan v0.4.4 pada Naskah_LaTeX.docx:
1. find_boundary mundur <= 6 paragraf NON-entri apa pun — 2 gumpalan span
   yatim nomencl (829 & 2.631 karakter, tanpa <w:tab/> dan tanpa nomor
   halaman sehingga entry_like==False) + 4 byline INTISARI = pas 6, jadi
   gumpalan ikut dianggap byline dan TIDAK terganti blok barunya.
2. Pencarian maju kursor vcur oracle jarak vertikal tanpa jendela: teks
   label gambar TikZ ("Filter Kalman 1D", jadi SVG di DOCX) nyangkut di sel
   Tabel 5.6 ratusan paragraf di depan — kursor melompat dan 12 halaman
   kelewat tanpa w:before.
3. Paragraf list di-skip seluruhnya dari tata letak sehingga Insideitemize
   (basedOn Normal) jatuh ke spasi tunggal rata kiri, bukan 1.5 justify.
"""
import zipfile

import tex2docx_engine as pkg  # via tests/conftest.py

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
PARA_RE = r'<w:p\b[^>]*?(?:/>|>.*?</w:p>)'


def _docx(tmp_path, paras, name='in.docx'):
    body = ''.join(paras)
    xml = (f'<?xml version="1.0" encoding="UTF-8"?>'
           f'<w:document xmlns:w="{W_NS}"><w:body>{body}</w:body></w:document>')
    path = tmp_path / name
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('word/document.xml', xml)
    return str(path)


def _read_docx(path):
    return zipfile.ZipFile(path).read('word/document.xml').decode('utf-8')


def _paras(xml):
    import re
    return re.findall(PARA_RE, xml, re.S)


def _p(text):
    return f'<w:p><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'


# ---------- 1. gumpalan span yatim nomencl ikut terganti ----------

def test_gumpalan_nomencl_terganti_byline_utuh(tmp_path):
    r"""find_boundary tak boleh mundur melewati paragraf PANJANG: byline
    abstrak selalu baris pendek; gumpalan nomencl 800-2600+ karakter di
    antara daftar dan heading berikutnya adalah konten lama yang wajib
    terganti, bukan dianggap bagian byline."""
    blob1 = _p('Lambang RomawiLambang Kuantitas Satuan ' * 22)          # ~858 kr
    blob2 = _p('SingkatanArtificial Intelligence (Kecerdasan Buatan)'
               'Access Point ' * 40)                                    # ~2640 kr
    bylines = [_p('PENINGKATAN AKURASI LOKALISASI SUMBER RADIOAKTIF '
                  'PENDETEKAN LOGAM BERAT DALAM SAMPEL LINGKUNGAN'),
               _p('oleh'),
               _p('Yasir Abdulaziz 22/497079/TK/54468'),
               _p('Diajukan kepada Departemen Teknik Nuklir dan Teknik '
                  'Fisika Fakultas Teknik')]
    paras = ([_p('DAFTAR LAMBANG DAN SINGKATAN'), blob1, blob2]
             + bylines + [_p('INTISARI')])
    src = _docx(tmp_path, paras)

    (tmp_path / 'skripsi.toc').write_text(
        '\\contentsline{chapter}{INTISARI}{v}\n', encoding='utf-8')
    (tmp_path / 'skripsi.nls').write_text(
        '\\nomgroup{A}\n'
        '\\item [\\textbf {Lambang Romawi}] \\begingroup \n'
        '\\item [{$A$}] \\begingroup matriks transisi status'
        '\\nomunit\\{$m^2/s$\\}\n', encoding='utf-8')

    dst = str(tmp_path / 'out.docx')
    pkg.typography.run(src, dst, str(tmp_path), 'skripsi')

    doc = _read_docx(dst)
    assert 'matriks transisi status' in doc       # entri baru dari .nls
    assert 'Lambang Kuantitas' not in doc         # gumpalan lama terganti
    assert 'Access Point' not in doc
    assert 'Yasir Abdulaziz' in doc               # 4 byline INTISARI utuh
    assert '>INTISARI<' in doc                    # heading berikutnya utuh


# ---------- 2. jendela pencarian maju kursor vcur ----------

def _fake_doc(n, special):
    idxs = list(range(n))
    texts = {i: f'para{i}' for i in idxs}         # filler < 8 kr: bebas noise fuzzy
    texts.update(special)
    return idxs, texts


def test_vcur_label_tikz_tak_melompat_ke_sel_jauh():
    r"""Baris PDF dari teks gambar TikZ tidak ada sebagai paragraf DOCX di
    sekitar kursor — pencarian harus None (skip), BUKAN nyangkut di sel
    tabel ber-teks sama ratusan paragraf di depan."""
    idxs, texts = _fake_doc(200, {180: 'filterkalman1d'})   # sel Tabel 5.6
    assert pkg.graft._find_line_target('filterkalman1d', texts, idxs, 40) is None
    texts[45] = 'filterkalman1d'                            # kecocokan wajar
    assert pkg.graft._find_line_target('filterkalman1d', texts, idxs, 40) == 45


def test_vcur_jendela_default_120_paragraf():
    idxs, texts = _fake_doc(205, {159: 'targetunik', 160: 'targetunik'})
    assert pkg.graft._find_line_target('targetunik', texts, idxs, 40) == 159
    texts[159] = 'para159'          # filler pulih — kini hanya idx 160 yang cocok
    assert pkg.graft._find_line_target('targetunik', texts, idxs, 40) is None


def test_vcur_dua_arah_sel_tabel():
    """Baris PDF melebihi satu paragraf DOCX (sel 'Pembimbing Utama : ...'
    = 3 sel) — cukup sel pertamanya awalan baris PDF."""
    idxs, texts = _fake_doc(40, {20: 'pembimbingutama'})
    assert pkg.graft._find_line_target('pembimbingutamadrxyz', texts, idxs, 0) == 20


# ---------- 3. paragraf list: 1.5 spasi + justify, indent list utuh ----------

def test_list_dapat_spasi_justify_tanpa_firstline(tmp_path):
    body = _p('Ini paragraf badan berusia panjang lebih dari enam puluh '
              'karakter agar ikut ditata sebagai prosa isi bab.')
    long_body = ('<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr>'
                 '<w:r><w:t>' + body + '</w:t></w:r></w:p>')
    num_list = ('<w:p><w:pPr>'
                '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr>'
                '<w:spacing w:before="100" w:after="0"/>'
                '<w:jc w:val="start"/></w:pPr>'
                '<w:r><w:t>Item pendek satu</w:t></w:r></w:p>')
    style_list = ('<w:p><w:pPr><w:pStyle w:val="Insideitemize"/>'
                  '<w:jc w:val="start"/></w:pPr>'
                  '<w:r><w:t>Item gaya itemize</w:t></w:r></w:p>')
    src = _docx(tmp_path, [_p('BAB I PENDAHULUAN'), long_body, num_list, style_list])
    dst = str(tmp_path / 'out.docx')
    pkg.typography.run(src, dst, str(tmp_path), 'skripsi')

    paras = _paras(_read_docx(dst))
    num_p = next(p for p in paras if '<w:numPr>' in p)
    style_p = next(p for p in paras if 'Insideitemize' in p)
    body_p = next(p for p in paras if 'enam puluh' in p)

    # list: 1.5 spasi + justify; spacing milik list (w:before="100") lestari
    assert 'w:line="360"' in num_p and 'w:lineRule="auto"' in num_p
    assert '<w:jc w:val="both"/>' in num_p
    assert 'w:before="100"' in num_p
    assert 'w:firstLine' not in num_p            # hanging indent tak tersentuh
    assert '<w:jc w:val="start"/>' not in num_p  # rata kiri LO diganti

    assert 'w:line="360"' in style_p and '<w:jc w:val="both"/>' in style_p
    assert 'w:firstLine' not in style_p

    # prosa badan tetap: firstLine 3em + 1.5 spasi + justify
    assert 'w:firstLine="720"' in body_p
    assert 'w:line="360"' in body_p and '<w:jc w:val="both"/>' in body_p


def test_upsert_spacing_merge_pertahankan_before_after(tmp_path):
    r"""w:spacing di-merge atributnya, bukan ditimpa: w:before/w:after yang
    sudah ada pada paragraf (jarak terukur graft pada konversi ulang) tak
    boleh hilang saat LINE_15 hanya menambah w:line."""
    para = ('<w:p><w:pPr><w:spacing w:before="249" w:after="120"/></w:pPr>'
            '<w:r><w:t>Paragraf prosa panjang dengan spacing terukur yang '
            'wajib dipertahankan setelah penataan ulang.</w:t></w:r></w:p>')
    src = _docx(tmp_path, [_p('BAB I PENDAHULUAN'), para])
    dst = str(tmp_path / 'out.docx')
    pkg.typography.run(src, dst, str(tmp_path), 'skripsi')

    body_p = next(p for p in _paras(_read_docx(dst)) if 'terukur' in p)
    assert '<w:spacing w:before="249" w:after="120" w:line="360" ' \
           'w:lineRule="auto"/>' in body_p
