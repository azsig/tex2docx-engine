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
        if cx > 25 * EMU_CM and 0.8 < ratio < 1.25:      # logo persegi raksasa -> 5.5 cm
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
