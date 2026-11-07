"""Unit test mesin marker fix_odt — stdlib murni, tanpa engine eksternal.

Membangun ODT kecil secara in-memory (content.xml + styles.xml), menjalankan
fix_odt.run, lalu menegaskan marker -> automatic style ODT."""
import io
import sys
import zipfile

sys.path.insert(0, str(__import__('pathlib').Path(__file__).resolve().parents[1] / 'src'))
from tex2docx_engine import fix_odt  # noqa: E402

NS_OFFICE = 'urn:oasis:names:tc:opendocument:xmlns:office:1.0'
NS_TEXT = 'urn:oasis:names:tc:opendocument:xmlns:text:1.0'
NS_STYLE = 'urn:oasis:names:tc:opendocument:xmlns:style:1.0'

CONTENT_TMPL = """<?xml version="1.0" encoding="utf-8"?>
<office:document-content xmlns:office="{o}" xmlns:text="{t}" xmlns:style="{s}"
 xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
<office:body><office:text>
{body}
</office:text></office:body></office:document-content>"""


def build_odt(path, body, heading_align=None):
    styles = ('<?xml version="1.0" encoding="utf-8"?>'
              f'<office:document-styles xmlns:office="{NS_OFFICE}" '
              f'xmlns:style="{NS_STYLE}" xmlns:fo="http://www.w3.org/1999/XSL/Format">'
              f'<office:styles>'
              + (f'<style:style style:name="Heading-1" style:family="paragraph">'
                 f'<style:paragraph-properties fo:text-align="{heading_align}"/>'
                 f'</style:style>' if heading_align else '')
              + '</office:styles></office:document-styles>')
    content = CONTENT_TMPL.format(o=NS_OFFICE, t=NS_TEXT, s=NS_STYLE, body=body)
    with zipfile.ZipFile(path, 'w') as z:
        z.writestr('content.xml', content)
        z.writestr('styles.xml', styles)


def content_of(path):
    with zipfile.ZipFile(path) as z:
        return z.read('content.xml').decode('utf-8')


def P(txt, style='Text-body'):
    return f"<text:p text:style-name='{style}'>{txt}</text:p>"


def test_vs_marker_jadi_margin_top(tmp_path):
    src, dst = tmp_path / 'a.odt', tmp_path / 'b.odt'
    build_odt(str(src), f"{P('Satu.')}<!--t2dvs:28.45pt-->{P('Dua.')}")
    fix_odt.run(str(src), str(dst))
    c = content_of(dst)
    assert 'fo:margin-top="28.45pt"' in c
    assert 't2dvs' not in c and '<!--' not in c    # marker tak boleh bocor
    assert '>Dua.</text:p>' in c               # konten utuh


def test_glue_beruntun_dijumlah(tmp_path):
    src, dst = tmp_path / 'a.odt', tmp_path / 'b.odt'
    build_odt(str(src),
              f"{P('Satu.')}<!--t2dvs:10pt--><!--t2dvs:5.5pt-->{P('Dua.')}")
    fix_odt.run(str(src), str(dst))
    assert 'fo:margin-top="15.5pt"' in content_of(dst)


def test_vs_kecil_dan_negatif_dibuang(tmp_path):
    src, dst = tmp_path / 'a.odt', tmp_path / 'b.odt'
    build_odt(str(src),
              f"{P('Satu.')}<!--t2dvs:0.5pt-->{P('Dua.')}"
              f"<!--t2dvs:6pt--><!--t2dvs:-6pt-->{P('Tiga.')}")
    fix_odt.run(str(src), str(dst))
    c = content_of(dst)
    assert 'fo:margin-top' not in c


def test_align_scope_sampai_end(tmp_path):
    src, dst = tmp_path / 'a.odt', tmp_path / 'b.odt'
    build_odt(str(src),
              f"{P('Atas.')}<!--t2dalign:center-->{P('Tengah satu.')}"
              f"{P('Tengah dua.')}<!--t2dalign:end-->{P('Bawah.')}")
    fix_odt.run(str(src), str(dst))
    c = content_of(dst)
    # dua paragraf dalam scope -> keduanya memakai automatic style t2dP
    # (dedup: SATU definisi style, DUA referensi) yang membawa text-align
    import re as _re
    refs = _re.findall(r'text:style-name="(t2dP\d+)">Tengah', c)
    assert len(refs) == 2 and len(set(refs)) == 1
    assert 'fo:text-align="center"' in c
    # di luar scope tidak kena
    i_atas = c.index('>Atas.<')
    i_bawah = c.index('>Bawah.<')
    seg_bawah = c[i_bawah:c.index('</text:p>', i_bawah)]
    assert 'text-align' not in seg_bawah
    _ = i_atas


def test_brk_jadi_break_before(tmp_path):
    src, dst = tmp_path / 'a.odt', tmp_path / 'b.odt'
    build_odt(str(src), f"{P('Satu.')}<!--t2dbrk-->{P('Dua.')}")
    fix_odt.run(str(src), str(dst))
    assert 'fo:break-before="page"' in content_of(dst)


def test_brk_trailing_di_dalam_paragraf(tmp_path):
    """\newpage tanpa blank line: marker di ujung paragraf -> break blok
    BERIKUTNYA, bukan dibuang."""
    src, dst = tmp_path / 'a.odt', tmp_path / 'b.odt'
    build_odt(str(src),
              f"{P('Satu.')}<text:p text:style-name='Text-body'>Dua. <!--t2dbrk--></text:p>{P('Tiga.')}")
    fix_odt.run(str(src), str(dst))
    assert 'fo:break-before="page"' in content_of(dst)


def test_marker_intra_paragraf_dibuang(tmp_path):
    src, dst = tmp_path / 'a.odt', tmp_path / 'b.odt'
    build_odt(str(src),
              "<text:p text:style-name='Text-body'>Awal <!--t2dvs:30pt--> lanjut</text:p>")
    fix_odt.run(str(src), str(dst))
    c = content_of(dst)
    assert 'fo:margin-top' not in c and 't2d' not in c


def test_marker_sebelum_blok_pertama_dibuang(tmp_path):
    src, dst = tmp_path / 'a.odt', tmp_path / 'b.odt'
    build_odt(str(src), f"<!--t2dvs:20pt--><!--t2dbrk-->{P('Satu.')}")
    fix_odt.run(str(src), str(dst))
    c = content_of(dst)
    assert 'fo:margin-top' not in c and 'fo:break-before' not in c


def test_harden_heading_align_dari_styles(tmp_path):
    src, dst = tmp_path / 'a.odt', tmp_path / 'b.odt'
    build_odt(str(src),
              "<text:h text:style-name='Heading-1'>Judul</text:h>",
              heading_align='center')
    fix_odt.run(str(src), str(dst))
    c = content_of(dst)
    assert 'fo:text-align="center"' in c
    assert 'style:parent-style-name="Heading-1"' in c


def test_tanpa_marker_tak_berubah(tmp_path):
    src, dst = tmp_path / 'a.odt', tmp_path / 'b.odt'
    body = f"{P('Satu.')}{P('Dua.')}"
    build_odt(str(src), body)
    fix_odt.run(str(src), str(dst))
    c = content_of(dst)
    assert 't2dP' not in c and 'office:automatic-styles' not in c


def test_span_yatim_dalam_section_dibungkus(tmp_path):
    """env center = text:section berisi span yatim (ODF invalid tanpa wrap);
    LO membuang span itu saat export DOCX."""
    src, dst = tmp_path / 'a.odt', tmp_path / 'b.odt'
    body = ("<text:section text:name='begin-end-env-1' text:style-name='begin-end-env'>"
            "<text:p text:style-name='begin-env-p'/>"
            "<text:span text:style-name='cmbx-12'>JUDUL </text:span>"
            "<text:p text:style-name='end-env-p'/></text:section>")
    build_odt(str(src), body)
    fix_odt.run(str(src), str(dst))
    c = content_of(dst)
    # span harus langsung berada di dalam <text:p> (bukan anak section)
    assert '<text:p><text:span text:style-name="cmbx-12">JUDUL </text:span></text:p>' in c


def test_align_scope_section_dipropagasikan(tmp_path):
    src, dst = tmp_path / 'a.odt', tmp_path / 'b.odt'
    body = ("<text:section text:name='e1' text:style-name='begin-end-env'>"
            f"{P('Dalam.')}</text:section>")
    build_odt(str(src), f"<!--t2dalign:center-->{body}<!--t2dalign:end-->{P('Luar.')}")
    fix_odt.run(str(src), str(dst))
    c = content_of(dst)
    i = c.index('>Dalam.<')
    seg = c[c.rfind('text:style-name', 0, i):i]
    assert 't2dP' in seg                    # paragraf dalam section diberi style
    assert 'fo:text-align="center"' in c


def test_paragraf_marker_only_dihapus(tmp_path):
    """HCode marker di vmode dilahirkan t4ht sebagai text:p HAMPA berisi
    komentar — paragraf itu harus dibuang total, markernya tetap bekerja
    untuk blok berikutnya (skripsi: +300 paragraf hampa tanpa ini)."""
    src, dst = tmp_path / 'a.odt', tmp_path / 'b.odt'
    body = (f"{P('Satu.')}"
            "<text:p text:style-name='Text-body'><!--t2dvs:12pt--></text:p>"
            f"{P('Dua.')}")
    build_odt(str(src), body)
    fix_odt.run(str(src), str(dst))
    c = content_of(dst)
    assert 'fo:margin-top="12.0pt"' in c     # marker tetap sampai blok Dua
    assert '<text:p text:style-name="Text-body"><!--' not in c  # paragraf hampa hilang
    # jumlah paragraf menyusut: Satu & Dua saja
    import re as _re
    assert len(_re.findall(r'<text:p[ >]', c)) == 2


def test_paragraf_marker_only_align_buka_scope(tmp_path):
    src, dst = tmp_path / 'a.odt', tmp_path / 'b.odt'
    body = (f"{P('Atas.')}"
            "<text:p text:style-name='Text-body'><!--t2dalign:center--></text:p>"
            f"{P('Tengah.')}"
            "<text:p text:style-name='Text-body'><!--t2dalign:end--></text:p>"
            f"{P('Bawah.')}")
    build_odt(str(src), body)
    fix_odt.run(str(src), str(dst))
    c = content_of(dst)
    assert 'fo:text-align="center"' in c
    import re as _re
    assert len(_re.findall(r'<text:p[ >]', c)) == 3  # dua paragraf hampa hilang
    i_bawah = c.index('>Bawah.<')
    seg = c[c.rfind('text:style-name', 0, i_bawah):i_bawah]
    assert 't2dP' not in seg                 # scope tak bocor ke Bawah


def test_align_end_ekor_paragraf_menutup_scope_walk(tmp_path):
    """align:end di EKOR paragraf (marker tail) harus menutup scope yang
    dibuka marker ANTAR-BLOK — bukan hanya scope intra-paragraf. Bug nyata
    skripsi: \\centering kepala bab membuka scope antar-blok, end aftergroup
    jatuh di ekor paragraf heading; scope bocor ke seluruh isi bab
    (Heading-2 jadi center padahal kiri di PDF)."""
    src, dst = tmp_path / 'a.odt', tmp_path / 'b.odt'
    body = (f"{P('Atas.')}"
            "<!--t2dalign:center-->"
            "<text:p text:style-name='Text-body'>Tengah. <!--t2dalign:end--></text:p>"
            f"{P('Bawah.')}")
    build_odt(str(src), body)
    fix_odt.run(str(src), str(dst))
    c = content_of(dst)
    assert 'fo:text-align="center"' in c       # 'Tengah.' kena scope
    i = c.index('>Tengah.')
    seg = c[c.rfind('text:style-name', 0, i):i]
    assert 't2dP' in seg
    i = c.index('>Bawah.<')                     # 'Bawah.' bebas scope
    seg = c[c.rfind('text:style-name', 0, i):i]
    assert 't2dP' not in seg


def test_repair_luka_ujung_bibliografi(tmp_path):
    """Dua luka struktural ujung dokumen tex4ht yang terverifikasi di skripsi
    (XML tak valid apa adanya): (1) </text:p> LIAR setelah penutup kontainer
    pustaka — bekas pemulihan "Missing }" \\end{thebibliography}; (2) marker
    <!--t2dbrk--> terpancar SETELAH </office:document-content> tertutup.
    fix_odt harus tetap bisa memakannya (repair pra-parse), markernya tetap
    bekerja, dan bibliografi tak dibungkus jadi inline yatim."""
    raw = ('<?xml version="1.0" encoding="utf-8"?>'
           f'<office:document-content xmlns:office="{NS_OFFICE}" '
           f'xmlns:text="{NS_TEXT}" xmlns:style="{NS_STYLE}">'
           '<office:body><office:text>'
           + P('Satu.')
           + '<!--t2dvs:12pt-->'
           + '<text:bibliography text:name="bib-1">'
           + '<text:index-body>'
           + P('Entri pustaka.')
           + '</text:index-body></text:bibliography>'
           + '      </text:p> '
           + P('Dua.')
           + '</office:text></office:body></office:document-content>'
           + '<!--t2dbrk--><!--t2dbrk-->')
    src, dst = tmp_path / 'a.odt', tmp_path / 'b.odt'
    with zipfile.ZipFile(src, 'w') as z:
        z.writestr('content.xml', raw)
        z.writestr('styles.xml',
                   '<?xml version="1.0" encoding="utf-8"?>'
                   f'<office:document-styles xmlns:office="{NS_OFFICE}" '
                   f'xmlns:style="{NS_STYLE}"><office:styles/>'
                   f'</office:document-styles>')
    fix_odt.run(str(src), str(dst))          # tanpa repair: ParseError
    c = content_of(dst)
    assert 'Entri pustaka.' in c             # konten pustaka selamat
    assert 'fo:margin-top="12.0pt"' in c     # marker menembus kontainer -> Dua
    assert c.rstrip().endswith('</office:document-content>')  # junk pasca-root hilang
    import re as _re
    assert not _re.search(r'</text:bibliography>\s*</text:p>', c)  # liar hilang
    # bibliografi = blok, tak dibungkus <text:p> oleh wrap_orphans
    assert '<text:p><text:bibliography' not in c


def test_marker_dalam_section_diproses(tmp_path):
    """Marker pada paragraf DI DALAM text:section (env center) harus tetap
    bekerja: paragraf hampa pembawa marker dibuang, align+vs menempel ke
    blok konten di dalam section. (Bug nyata skripsi: klasifikasi memakai
    ancestor top-level -> section -> 794 marker align/vs tersesat,
    align=0.)"""
    src, dst = tmp_path / 'a.odt', tmp_path / 'b.odt'
    body = ("<text:section text:name='e1' text:style-name='begin-end-env'>"
            "<text:p text:style-name='begin-env-p'/>"
            "<text:p text:style-name='Text-body'>"
            "<!--t2dalign:center--><!--t2dvs:12pt--></text:p>"
            f"{P('Dalam satu.')}"
            f"{P('Dalam dua.')}"
            "<text:p text:style-name='Text-body'><!--t2dalign:end--></text:p>"
            "<text:p text:style-name='end-env-p'/>"
            "</text:section>")
    build_odt(str(src), f"{P('Luar.')}{body}")
    fix_odt.run(str(src), str(dst))
    c = content_of(dst)
    assert 'fo:text-align="center"' in c
    assert 'fo:margin-top="12.0pt"' in c
    for txt in ('Dalam satu.', 'Dalam dua.'):
        i = c.index('>' + txt)
        seg = c[c.rfind('text:style-name', 0, i):i]
        assert 't2dP' in seg, txt          # keduanya kena scope center
    # paragraf hampa pembawa marker + paragraf tepi env kosong dibuang:
    # tinggal Luar + Dua paragraf dalam section
    import re as _re
    assert len(_re.findall(r'<text:p[ >]', c)) == 3
    i = c.index('>Luar.<')
    seg = c[c.rfind('text:style-name', 0, i):i]
    assert 't2dP' not in seg              # scope tak bocor keluar section


def test_paragraf_whitespace_dibuang_kecuali_sel_tabel(tmp_path):
    """Paragraf whitespace-only = artefak rantai tex4ht/xtpipes (skripsi:
    ~450 buah di dalam section env center) — dibuang; paragraf di dalam
    table:table-cell DIPERTAHANKAN (ODF mewajibkan >= 1 paragraf per sel)."""
    src, dst = tmp_path / 'a.odt', tmp_path / 'b.odt'
    body = ("<text:section text:name='e1'>"
            "<text:p text:style-name='Text-body'>       </text:p>"
            f"{P('Isi.')}"
            "</text:section>"
            "<table:table table:name='t1'><table:table-row><table:table-cell>"
            "<text:p text:style-name='Text-body'> </text:p>"
            "</table:table-cell></table:table-row></table:table>")
    build_odt(str(src), body)
    fix_odt.run(str(src), str(dst))
    c = content_of(dst)
    import re as _re
    n = _re.findall(r'<text:p[ >]', c)
    assert len(n) == 2                     # 'Isi.' + paragraf wajib sel tabel
    assert 'Isi.' in c
    assert "<table:table-cell><text:p" in c  # sel tetap punya paragrafnya
