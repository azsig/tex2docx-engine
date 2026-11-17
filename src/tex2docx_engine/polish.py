"""
polish — Bersih-bersih artefak konversi (layer 1 dari 3 post-processing).
1. Artefak hyperlink lintas-file ODT dihapus (w:hyperlink via rels + field HYPERLINK)
2. Font diseragamkan (default: Times New Roman) di document.xml + styles.xml
3. Ukuran gambar di-clamp ke lebar halaman; logo persegi raksasa -> ukuran wajar
"""
import re
import sys
import zipfile

EMU_CM = 360000
MAX_W_CM, MAX_H_CM, LOGO_CM = 13.5, 18.0, 5.5


def run(SRC, DST, FONT='Times New Roman'):
    zin = zipfile.ZipFile(SRC)
    data = {n: zin.read(n) for n in zin.namelist()}
    zin.close()

    doc = data['word/document.xml'].decode('utf-8')
    rels = data['word/_rels/document.xml.rels'].decode('utf-8')
    styles = data['word/styles.xml'].decode('utf-8')
    log = []

    # ---------- 1a. Unwrap w:hyperlink yang menunjuk file .odt ----------
    rid_target = {}
    for tag in re.findall(r'<Relationship\b[^>]*/>', rels):
        mid = re.search(r'\bId="([^"]+)"', tag)
        mtg = re.search(r'\bTarget="([^"]+)"', tag)
        if mid and mtg:
            rid_target[mid.group(1)] = mtg.group(1)
    odt_rids = {rid for rid, t in rid_target.items() if '.odt' in t}

    def _unwrap(m):
        rid = re.search(r'r:id="([^"]+)"', m.group(1))
        if rid and rid.group(1) in odt_rids:
            return m.group(2)
        return m.group(0)

    doc = re.sub(r'<w:hyperlink([^>]*)>(.*?)</w:hyperlink>', _unwrap, doc, flags=re.S)

    # ---------- 1b. Bersihkan field HYPERLINK kompleks ber-target .odt ----------
    def clean_odt_fields(doc):
        out, pos = [], 0
        for m in re.finditer(r'<w:fldChar w:fldCharType="begin"\s*/?>', doc):
            if m.start() < pos:
                continue
            e = re.search(r'<w:fldChar w:fldCharType="end"\s*/?>', doc[m.start():])
            if not e:
                continue
            region = doc[m.start():m.start() + e.end()]
            if '.odt' not in region:
                continue
            out.append(doc[pos:m.start()])
            kept = re.sub(r'<w:r\b[^>]*>(?:(?!</w:r>).)*?(?:<w:fldChar|<w:instrText)(?:(?!</w:r>).)*?</w:r>',
                          '', region, flags=re.S)
            out.append(kept)
            pos = m.start() + e.end()
        out.append(doc[pos:])
        return ''.join(out)

    n_before = doc.count('.odt')
    doc = clean_odt_fields(doc)
    log.append(f'artefak .odt dibersihkan : {n_before} -> {doc.count(".odt")}')

    # ---------- 2. Font seragam ----------
    def unify_font(s):
        for attr in ('w:ascii', 'w:hAnsi', 'w:eastAsia', 'w:cs'):
            s = re.sub(attr + r'="[^"]*"', attr + '="' + FONT + '"', s)
        for attr in ('w:asciiTheme', 'w:hAnsiTheme', 'w:eastAsiaTheme', 'w:cstheme'):
            s = re.sub(r'\s+' + attr + r'="[^"]*"', '', s)
        return s

    doc = unify_font(doc)
    styles = unify_font(styles)
    log.append(f'font diseragamkan       : {FONT} (document.xml + styles.xml)')

    # ---------- 3. Ukuran gambar ----------
    def fix_extents(d):
        m = re.search(r'<wp:extent cx="(\d+)" cy="(\d+)"/>', d)
        if not m:
            return d
        cx, cy = int(m.group(1)), int(m.group(2))
        ncx, ncy, ratio = cx, cy, cy / cx
        # Logo: HANYA lewat NAMA gambar (docPr name, mis. 'logoUGM.png').
        # Dulu `or cx > 25cm` ikut menjadikan gambar isi resolusi tinggi
        # "logo": 6 gambar skripsi uji (cx>25cm, rasio 0.5-2.0) menyusut ke
        # 5.5cm padahal LaTeX memesan ~10-13.5cm — sisa halaman menganga.
        # Gambar raksasa non-logo cukup di-clamp ke batas halaman (else).
        # Ukuran persis dari LaTeX sendiri ditanam fix_odt dari marker
        # t2dimg (width/height/scale hasil evaluasi engine) — jalur nama di
        # sini tinggal fallback bila marker tak ada.
        nm = re.search(r'<wp:docPr[^>]*\bname="([^"]*)"', d)
        is_logo = bool(nm and 'logo' in nm.group(1).lower())
        if is_logo and 0.5 < ratio < 2.0 and cx > 2 * LOGO_CM * EMU_CM:
            ncx, ncy = int(LOGO_CM * EMU_CM), int(LOGO_CM * EMU_CM * ratio)
        else:
            if cx > MAX_W_CM * EMU_CM:
                ncx, ncy = int(MAX_W_CM * EMU_CM), int(MAX_W_CM * EMU_CM * ratio)
            if ncy > MAX_H_CM * EMU_CM:
                ncy, ncx = int(MAX_H_CM * EMU_CM), int(MAX_H_CM * EMU_CM / ratio)
        if (ncx, ncy) == (cx, cy):
            return d
        return d.replace(f'cx="{cx}" cy="{cy}"', f'cx="{ncx}" cy="{ncy}"')

    n_img = len(re.findall(r'<w:drawing>', doc))
    doc = re.sub(r'<w:drawing>.*?</w:drawing>', lambda m: fix_extents(m.group(0)), doc, flags=re.S)
    log.append(f'gambar diproses         : {n_img} drawing (clamp {MAX_W_CM}cm, logo {LOGO_CM}cm)')

    # ---------- 4. normalisasi persamaan -> tabel 3 kolom ----------
    # Tabel persamaan ooffice.4ht hanya 2 kolom TANPA lebar grid — LibreOffice
    # membagi 50/50 (4534|4536) dan rumus terdorong ke paruh kiri. Persamaan
    # di dalam \item (list-item ODF melarang table — diratakan fix_odt jadi
    # text:p ber-<m:oMath> + tab + nomor) kini bisa dirangkum kembali jadi
    # tabel tanpa menyentuh penomoran w:numId OOXML. Semua jadi SATU bentuk:
    # [kosong 12% | rumus center 76% | nomor kanan 12%], tanpa border,
    # vAlign center — kolom kiri kosong penyeimbang agar rumus tepat di
    # tengah halaman seperti \begin{equation} LaTeX.
    W_SIDE, W_MID = 1088, 6894

    def _build_3col_eq_table(formula_inner, num_inner):
        return (
            '<w:tbl>'
            '<w:tblPr>'
            '<w:tblW w:w="5000" w:type="pct"/>'
            '<w:jc w:val="center"/>'
            '<w:tblLayout w:type="fixed"/>'
            '<w:tblCellMar>'
            '<w:top w:w="0" w:type="dxa"/><w:start w:w="0" w:type="dxa"/>'
            '<w:bottom w:w="0" w:type="dxa"/><w:end w:w="0" w:type="dxa"/>'
            '</w:tblCellMar>'
            '</w:tblPr>'
            f'<w:tblGrid><w:gridCol w:w="{W_SIDE}"/>'
            f'<w:gridCol w:w="{W_MID}"/><w:gridCol w:w="{W_SIDE}"/></w:tblGrid>'
            '<w:tr>'
            f'<w:tc><w:tcPr><w:tcW w:w="{W_SIDE}" w:type="dxa"/>'
            '<w:tcBorders/><w:vAlign w:val="center"/></w:tcPr>'
            '<w:p><w:pPr><w:spacing w:before="60" w:after="60"/></w:pPr></w:p></w:tc>'
            f'<w:tc><w:tcPr><w:tcW w:w="{W_MID}" w:type="dxa"/>'
            '<w:tcBorders/><w:vAlign w:val="center"/></w:tcPr>'
            '<w:p><w:pPr><w:pStyle w:val="Equp"/>'
            '<w:spacing w:before="60" w:after="60"/><w:jc w:val="center"/></w:pPr>'
            f'{formula_inner}</w:p></w:tc>'
            f'<w:tc><w:tcPr><w:tcW w:w="{W_SIDE}" w:type="dxa"/>'
            '<w:tcBorders/><w:vAlign w:val="center"/></w:tcPr>'
            '<w:p><w:pPr><w:pStyle w:val="Equnump"/>'
            '<w:spacing w:before="60" w:after="60"/><w:jc w:val="end"/></w:pPr>'
            f'{num_inner}</w:p></w:tc>'
            '</w:tr>'
            '</w:tbl>'
        )

    n_eqtbl_up = [0]

    def _upgrade_2col_eq_tbl(m):
        tbl = m.group(0)
        if 'w:val="Equp"' not in tbl and 'w:val="Equnump"' not in tbl:
            return tbl
        tcs = re.findall(r'<w:tc\b[^>]*>.*?</w:tc>', tbl, flags=re.S)
        if len(tcs) != 2:
            return tbl
        p0 = re.search(r'<w:p\b[^>]*>(?:<w:pPr>.*?</w:pPr>)?(.*?)</w:p>',
                       tcs[0], flags=re.S)
        p1 = re.search(r'<w:p\b[^>]*>(?:<w:pPr>.*?</w:pPr>)?(.*?)</w:p>',
                       tcs[1], flags=re.S)
        if not (p0 and p1):
            return tbl
        n_eqtbl_up[0] += 1
        return _build_3col_eq_table(p0.group(1), p1.group(1))

    doc = re.sub(r'<w:tbl\b[^>]*>.*?</w:tbl>', _upgrade_2col_eq_tbl, doc, flags=re.S)

    # ekor paragraf persamaan ratakan list-item: ... <w:tab/> run + nomor run
    # + </w:p> — nomor persamaan (x.y) selalu run terakhir.
    RE_EQNUM_TAIL = re.compile(
        r'(<w:r\b[^>]*>(?:(?!</w:r>).)*?<w:tab\s*/>(?:(?!</w:r>).)*?</w:r>)\s*'
        r'(<w:r\b[^>]*>(?:(?!</w:r>).)*?<w:t[^>]*>\((\d+(?:\.\d+)*)\)</w:t>'
        r'\s*</w:r>)\s*</w:p>$', re.S)

    n_eqwrap = [0]

    def _wrap_flat_eq_para(m):
        p = m.group(0)
        if '<m:oMath' not in p:
            return p
        sp = RE_EQNUM_TAIL.search(p)
        if not sp:
            return p
        head = p[:sp.start()]
        head = re.sub(r'^<w:p\b[^>]*>(?:<w:pPr>.*?</w:pPr>)?', '', head,
                      count=1, flags=re.S)
        num_run = re.sub(r'<w:tab\s*/>', '', sp.group(2))
        n_eqwrap[0] += 1
        return _build_3col_eq_table(head, num_run)

    doc = re.sub(r'<w:p\b[^>]*>.*?</w:p>', _wrap_flat_eq_para, doc, flags=re.S)
    if n_eqtbl_up[0] or n_eqwrap[0]:
        log.append(f'persamaan dinormalisasi  : {n_eqtbl_up[0]} tabel 2->3 '
                   f'kolom, {n_eqwrap[0]} paragraf dirangkum tabel')

    data['word/document.xml'] = doc.encode('utf-8')
    data['word/styles.xml'] = styles.encode('utf-8')
    with zipfile.ZipFile(DST, 'w', zipfile.ZIP_DEFLATED) as z:
        for n, b in data.items():
            z.writestr(n, b)

    print(f'[polish] {SRC} -> {DST}')
    for line in log:
        print('  -', line)


def main(argv=None):
    a = sys.argv[1:] if argv is None else argv
    if len(a) < 2:
        sys.exit('pemakaian: polish <in.docx> <out.docx> [font]')
    run(a[0], a[1], a[2] if len(a) > 2 else 'Times New Roman')


if __name__ == '__main__':
    main()
