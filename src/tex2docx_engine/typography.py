r"""
typography_docx.py — Lapisan tipografi untuk pipeline engine-mode.
Sumber kebenaran: .cls (parindent 3em, onehalfspace), file .toc/.lof/.lot/.nls
(entri daftar + nomor halaman dari run pdflatex), struktur dokumen (front matter vs isi).
Aturan:
  A. Bersihkan w:br page sisa konversi LO + collapse paragraf kosong beruntun
  B. Ganti blok DAFTAR ISI/TABEL/GAMBAR dengan entri dari .toc/.lot/.lof
     (indent per level, tab kanan titik-titik + nomor halaman).
     Parser brace-balanced: menangani \numberline {X}Judul MAUPUN varian
     bersarang \numberline {X}{\ignorespaces Judul} yang dihasilkan caption/hyperref.
  C. Ganti blok DAFTAR LAMBANG DAN SINGKATAN dengan entri dari .nls (nomencl),
     dikelompokkan per \nomgroup; nama grup dari \renewcommand{\nomgroup} di .cls.
  D. Front matter (sebelum BAB I) dibedakan per section:
     - prosa (KATA PENGANTAR, PERSEMBAHAN, MOTTO, INTISARI, ABSTRACT):
       indent baris pertama 3em + spasi 1.5, TANPA spacing antar paragraf (perilaku PDF)
     - surat/daftar (judul, pernyataan, pengesahan): spacing antar paragraf
     Isi BAB: justify + first-line indent 3em (tanpa spacing)
  E. (dihapus) Perataan byline abstrak & indent baris kata kunci kini murni
     GEOMETRI PDF (graft_pdf_layout.py) — bukan detektor kata.
Pemakaian: python3 typography_docx.py <in.docx> <out.docx> <dir_toc> <prefix_toc>
           prefix_toc = nama dasar file toc tanpa ekstensi (mis. skripsi)
"""
import glob
import os
import re
import sys
import zipfile
def run(SRC, DST, BASE, TOCBASE):

    PARINDENT = 720                                  # 3em @ 12pt = 36pt
    LINE_15 = '<w:spacing w:line="360" w:lineRule="auto"/>'
    SP_AFTER = '<w:spacing w:after="240" w:line="360" w:lineRule="auto"/>'
    JC_BOTH = '<w:jc w:val="both"/>'
    TOC_HEADS = ['DAFTAR ISI', 'DAFTAR TABEL', 'DAFTAR GAMBAR', 'DAFTAR LAMBANG DAN SINGKATAN']
    PPR_ORDER = ['w:pStyle', 'w:pageBreakBefore', 'w:tabs', 'w:spacing', 'w:ind', 'w:jc']

    zin = zipfile.ZipFile(SRC)
    data = {n: zin.read(n) for n in zin.namelist()}
    zin.close()
    doc = data['word/document.xml'].decode('utf-8')
    log = []

    # ---------- A. bersih-bersih ----------
    # (run kosong sisa penghapusan <w:br> dibiarkan — valid & tak terlihat di Word)
    n_br = len(re.findall(r'<w:br w:type="page"\s*/>', doc))
    doc = re.sub(r'<w:br w:type="page"\s*/>', '', doc)
    log.append(f'w:br page sisa LO dibuang: {n_br}')

    para_re = re.compile(r'<w:p\b[^>]*?(?:/>|>.*?</w:p>)', re.S)

    def ptext(raw):
        t = raw.replace('</w:t>', ' </w:t>')
        t = re.sub(r'<[^>]+>', '', t)
        return re.sub(r'\s+', ' ', t).strip()

    segs, last = [], 0
    for m in para_re.finditer(doc):
        segs.append(['text', doc[last:m.start()]])
        segs.append(['para', m.group(0)])
        last = m.end()
    segs.append(['text', doc[last:]])

    # collapse paragraf kosong beruntun (sumber "halaman kosong")
    n_collapsed = 0
    i = 0
    while i < len(segs):
        if segs[i][0] == 'para' and not ptext(segs[i][1]):
            j = i
            while j < len(segs) and segs[j][0] == 'para' and not ptext(segs[j][1]):
                j += 1
            for k in range(i + 1, j):
                if segs[k][0] == 'para':
                    segs[k] = ['text', '']
                    n_collapsed += 1
            i = j
        else:
            i += 1
    log.append(f'paragraf kosong dicollapse: {n_collapsed}')

    def pidx():
        return [k for k, (t, _) in enumerate(segs) if t == 'para']

    def find_para(pred):
        for k in pidx():
            if pred(segs[k][1]):
                return k
        return None

    # ---------- helpers LaTeX ----------
    def clean_latex(t):
        """Buang perintah format (simpan isinya), sisa perintah, kurung kurawal, dan $ math."""
        prev = None
        while prev != t:                               # \textit {Mean Absolute \textit {Error}}
            prev = t
            t = re.sub(r'\\(?:textit|textbf|emph|texttt|textsc|textrm)\s*\{([^{}]*)\}', r'\1', t)
        t = re.sub(r'\\ignorespaces|\\relax|\\large|\\bfseries|\\itshape', '', t)
        t = re.sub(r'\\[a-zA-Z@]+\*?', '', t)
        return re.sub(r'\s+', ' ', t.replace('{', '').replace('}', '').replace('$', '')).strip()

    def scan_group(s, i):
        """s[i] == '{' -> (isi_group, indeks_setelah_kurung_penutup); brace-balanced."""
        depth = 0
        for j in range(i, len(s)):
            if s[j] == '{':
                depth += 1
            elif s[j] == '}':
                depth -= 1
                if depth == 0:
                    return s[i + 1:j], j + 1
        return s[i + 1:], len(s)

    def parse_toc_file(path):
        r"""Parser \contentsline yang brace-balanced — menangani format klasik
        \\numberline {X}Judul dan bersarang \\numberline {X}{\\ignorespaces Judul}."""
        entries = []
        if not os.path.isfile(path):
            return entries
        txt = open(path, encoding='utf-8', errors='ignore').read()
        for m in re.finditer(r'\\contentsline\s*\{', txt):
            i = m.end() - 1
            level_raw, i = scan_group(txt, i)
            while i < len(txt) and txt[i].isspace():
                i += 1
            if i >= len(txt) or txt[i] != '{':
                continue
            title, i = scan_group(txt, i)
            while i < len(txt) and txt[i].isspace():
                i += 1
            page = ''
            if i < len(txt) and txt[i] == '{':
                page, _ = scan_group(txt, i)
            level = {'chapter': 0, 'section': 1, 'subsection': 2, 'subsubsection': 3}.get(level_raw.strip(), 1)
            t = title.strip()
            mnl = re.match(r'\\numberline\s*', t)
            if mnl:                                     # pisahkan nomor dari judul
                rest = t[mnl.end():]
                k = rest.find('{')
                if k != -1:
                    num, k2 = scan_group(rest, k)
                    after = rest[k2:].strip()
                    if after.startswith('{'):
                        inner, _ = scan_group(after, after.find('{'))
                        t = num + ' ' + inner
                    else:
                        t = num + ' ' + after
                else:
                    t = rest
            t = clean_latex(t)
            if t:
                entries.append((level, t, page.strip()))
        return entries

    def esc(s):
        return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

    def toc_para(level, title, page):
        ppr = ['<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="8500"/></w:tabs>', LINE_15]
        if level:
            ppr.append(f'<w:ind w:left="{240 * level}"/>')
        if level == 0:
            ppr.append('<w:rPr><w:b/></w:rPr>')
            rpr = '<w:rPr><w:b/></w:rPr>'
        else:
            rpr = ''
        return (f'<w:p><w:pPr>{"".join(ppr)}</w:pPr>'
                f'<w:r>{rpr}<w:t xml:space="preserve">{esc(title)}</w:t></w:r>'
                f'<w:r><w:tab/></w:r>'
                f'<w:r>{rpr}<w:t>{esc(page)}</w:t></w:r></w:p>')

    # ---------- B+C. ganti blok daftar ----------
    toc_entries = parse_toc_file(os.path.join(BASE, f'{TOCBASE}.toc'))
    front0 = {re.sub(r'[\s\W]+', '', t).lower() for lvl, t, _ in toc_entries if lvl == 0}

    # Posisi tabel: paragraf di dalam <w:tc> (untuk mengenali entri daftar lama
    # yang dirender tex4ht sebagai tabel, mis. nomencl).
    INTBL = set()
    _depth = 0
    for _i, (_t, _v) in enumerate(segs):
        if _t == 'text':
            _depth += _v.count('<w:tbl>') - _v.count('</w:tbl>')
        elif _depth > 0:
            INTBL.add(_i)

    def entry_like(k):
        """Paragraf mirip ENTRI DAFTAR lama (hasil tex4ht): di dalam tabel,
        bertab, atau diakhiri nomor halaman (1-3 digit). Paragraf byline abstrak
        (judul / oleh / nama+NIM / diajukan) TIDAK seperti itu — NIM berakhiran
        run digit >= 4 sehingga tidak dianggap nomor halaman. Struktural, bebas
        bahasa."""
        raw = segs[k][1]
        if k in INTBL or '<w:tab/>' in raw:
            return True
        return bool(re.search(r'(?<!\d)\d{1,3}\s*$', ptext(raw)))

    generated = set()          # indeks seg paragraf buatan kita → dikecualikan dari tata letak
    protected = []             # rentang (lo, hi) blok daftar → dikecualikan dari tata letak

    def find_boundary(lo):
        """Batas blok daftar: heading section berikutnya (nama daftar dari konstanta
        TOC_HEADS, atau entri bab level-0 dari .toc — keduanya struktural).
        Bila tepat sebelum heading itu ada rangkaian <= 6 paragraf NON-entri
        (byline abstrak berada di antara daftar terakhir dan heading abstrak),
        batas naik ke awal rangkaian supaya byline tidak ikut terhapus."""
        P = pidx()
        ordn = {k: n for n, k in enumerate(P)}
        head = next((k for k in P if k > lo and
                     (ptext(segs[k][1]) in TOC_HEADS or
                      re.sub(r'[\s\W]+', '', ptext(segs[k][1])).lower() in front0)), None)
        if head is None:
            return None
        start = head
        hi = ordn[head]
        for q in (P[hi - 1:hi - 7:-1] if hi else []):
            if q <= lo or not ptext(segs[q][1]) or entry_like(q):
                break
            start = q
        return start

    def replace_block(head, new_paras):
        hi = find_para(lambda r: ptext(r) == head)
        if hi is None:
            log.append(f'{head:<28s}: SKIP (heading tidak ketemu)')
            return
        end = find_boundary(hi)
        if end is None:
            log.append(f'{head:<28s}: SKIP (batas blok tidak ketemu — tidak jadi dihapus)')
            return
        n = 0
        for k in range(hi + 1, end):
            if segs[k][0] == 'para':
                segs[k] = ['text', '']
                n += 1
        for off, p in enumerate(new_paras):
            segs.insert(hi + 1 + off, ['para', p])
            generated.add(hi + 1 + off)
        protected.append((hi, end + len(new_paras)))
        log.append(f'{head:<28s}: {len(new_paras)} entri ({n} paragraf lama diganti)')

    for head, fname in (('DAFTAR ISI', f'{TOCBASE}.toc'),
                        ('DAFTAR TABEL', f'{TOCBASE}.lot'),
                        ('DAFTAR GAMBAR', f'{TOCBASE}.lof')):
        entries = parse_toc_file(os.path.join(BASE, fname))
        if entries:
            replace_block(head, [toc_para(l, t, p) for l, t, p in entries])

    # ---------- C. DAFTAR LAMBANG dari .nls (nomencl) ----------
    GREEK = {'alpha': 'α', 'beta': 'β', 'gamma': 'γ', 'delta': 'δ', 'epsilon': 'ε',
             'varepsilon': 'ε', 'zeta': 'ζ', 'eta': 'η', 'theta': 'θ', 'iota': 'ι',
             'kappa': 'κ', 'lambda': 'λ', 'mu': 'μ', 'nu': 'ν', 'xi': 'ξ', 'pi': 'π',
             'rho': 'ρ', 'sigma': 'σ', 'tau': 'τ', 'phi': 'φ', 'varphi': 'φ',
             'chi': 'χ', 'psi': 'ψ', 'omega': 'ω', 'Gamma': 'Γ', 'Delta': 'Δ',
             'Theta': 'Θ', 'Lambda': 'Λ', 'Xi': 'Ξ', 'Pi': 'Π', 'Sigma': 'Σ',
             'Phi': 'Φ', 'Psi': 'Ψ', 'Omega': 'Ω'}

    def clean_symbol(s):
        s = s.strip()
        while s.startswith('{') and s.endswith('}'):   # [{$A$}] → $A$
            s = s[1:-1].strip()
        m = re.fullmatch(r'\$(.+)\$', s)
        italic = bool(m)
        if m:
            s = m.group(1)
        # ganti hanya perintah Yunani yang dikenal; perintah lain dibiarkan utk clean_latex
        s = re.sub(r'\\([a-zA-Z]+)', lambda mm: GREEK.get(mm.group(1), '\\' + mm.group(1)), s)
        return clean_latex(s), italic

    def parse_nls(path):
        groups = []
        if not os.path.isfile(path):
            return groups
        cur = None
        for line in open(path, encoding='utf-8', errors='ignore'):
            mg = re.match(r'\s*\\nomgroup\{(\w+)\}', line)
            if mg:
                cur = {'key': mg.group(1), 'items': []}
                groups.append(cur)
                continue
            mi = re.match(r'\s*\\item\s*\[(.*)\]\s*\\begingroup\s*(.*)$', line)
            if mi and cur is not None:
                sym_raw, rest = mi.group(1), mi.group(2)
                mu = re.search(r'\\nomunit\s*\{(.*?)\}', rest)
                desc = rest[:mu.start()] if mu else rest
                desc = re.split(r'\\nomeqref', desc)[0]
                unit = clean_latex(mu.group(1)).replace('--', '–') if mu else ''
                sym, italic = clean_symbol(sym_raw)
                cur['items'].append({
                    'sym': sym, 'italic': italic, 'desc': clean_latex(desc),
                    'unit': unit, 'header': ('textbf' in sym_raw and 'Lambang' in sym_raw)})
        return groups

    def load_group_names():
        """Nama grup dari \renewcommand{\nomgroup} di .cls/.sty; fallback map standar."""
        names = {'A': 'Lambang Romawi', 'B': 'Lambang Yunani', 'C': 'Subskrip',
                 'D': 'Superskrip', 'E': 'Singkatan'}
        for f in sorted(glob.glob(os.path.join(BASE, '*.cls'))) + \
                 sorted(glob.glob(os.path.join(BASE, '*.sty'))):
            try:
                txt = open(f, encoding='utf-8', errors='ignore').read()
            except OSError:
                continue
            m = re.search(r'\\renewcommand\{\\nomgroup\}', txt)
            if not m:
                continue
            block = txt[m.start():m.start() + 2500]
            for mm in re.finditer(r'\\equal\{#1\}\{(\w+)\}\}(?:(?!\\equal).)*?\\textbf\s*\{([^}]+)\}',
                                  block, re.S):
                names[mm.group(1)] = clean_latex(mm.group(2))
            break
        return names

    TAB2 = '<w:tabs><w:tab w:val="left" w:pos="1080"/></w:tabs>'
    TAB3 = '<w:tabs><w:tab w:val="left" w:pos="1080"/><w:tab w:val="right" w:pos="8504"/></w:tabs>'
    NOM_SP = '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>'

    def nom_row(sym, desc, unit, has_unit, bold=False, italic=False, before=None, brk=False):
        ppr = ['<w:pageBreakBefore/>'] if brk else []
        ppr.append(TAB3 if has_unit else TAB2)
        ppr.append(f'<w:spacing w:before="{before}" w:after="0" w:line="240" w:lineRule="auto"/>'
                   if before else NOM_SP)
        rpr = '<w:rPr>' + ('<w:b/>' if bold else '') + ('<w:i/>' if italic else '') + '</w:rPr>'
        runs = [f'<w:r>{rpr}<w:t xml:space="preserve">{esc(sym)}</w:t></w:r>',
                '<w:r><w:tab/></w:r>',
                f'<w:r>{rpr}<w:t xml:space="preserve">{esc(desc)}</w:t></w:r>']
        if has_unit:
            runs += ['<w:r><w:tab/></w:r>',
                     f'<w:r>{rpr}<w:t xml:space="preserve">{esc(unit)}</w:t></w:r>']
        return f'<w:p><w:pPr>{"".join(ppr)}</w:pPr>{"".join(runs)}</w:p>'

    groups = parse_nls(os.path.join(BASE, f'{TOCBASE}.nls'))
    if groups:
        gnames = load_group_names()
        paras = []
        for g in groups:
            items = [it for it in g['items'] if not it['header']]
            header = next((it for it in g['items'] if it['header']), None)
            has_unit = bool(header and header['unit']) or any(it['unit'] for it in items)
            is_e = g['key'].upper() == 'E'          # \newpage sebelum grup E di .cls
            paras.append(nom_row(gnames.get(g['key'], g['key']), '', '', False,
                                 bold=True, before=None if is_e else 708, brk=is_e))
            if header:
                paras.append(nom_row(header['sym'], header['desc'], header['unit'],
                                     has_unit, bold=True))
            for it in items:
                paras.append(nom_row(it['sym'], it['desc'], it['unit'], has_unit,
                                     italic=it['italic']))
        replace_block('DAFTAR LAMBANG DAN SINGKATAN', paras)

    # ---------- D. indent / spacing / justify ----------
    def in_protected(k):
        return any(lo < k < hi for lo, hi in protected)

    i_bab1 = next((k for k in pidx()
                   if ptext(segs[k][1]).startswith('BAB I') and not in_protected(k)), None)

    def upsert_ppr(raw, additions):
        """Sisipkan elemen pPr dengan urutan schema valid; ganti bila sudah ada."""
        if '<w:pPr' not in raw:
            m = re.search(r'<w:p\b[^>]*>', raw)
            return raw[:m.end()] + '<w:pPr>' + ''.join(additions) + '</w:pPr>' + raw[m.end():]
        for add in additions:
            tag = re.match(r'<(w:\w+)', add).group(1)
            raw = re.sub(f'<{tag}[^>]*/>|<{tag}[^>]*>.*?</{tag}>', '', raw, flags=re.S)
            later = PPR_ORDER[PPR_ORDER.index(tag) + 1:]
            m = re.search(r'<w:pPr[^>]*>', raw)
            ppr_start = m.end()
            insert_at = None
            for t in later:
                mm = re.search(f'<{t}[ />]', raw[ppr_start:])
                if mm:
                    insert_at = ppr_start + mm.start()
                    break
            if insert_at is None:
                mm = re.search(r'<w:rPr[ />]|</w:pPr>', raw[ppr_start:])
                insert_at = ppr_start + (mm.start() if mm else len(raw[ppr_start:]))
            raw = raw[:insert_at] + add + raw[insert_at:]
        return raw

    PROSE_RE = re.compile(r'PENGANTAR|PRAKATA|FOREWORD|PERSEMBAHAN|DEDICAT|MOTTO|'
                          r'INTISARI|ABSTRAK|ABSTRACT|RINGKASAN|R[ÉE]SUM[ÉE]', re.I)

    n_prose = n_letter = n_body = 0
    sec_prose = False
    for k in pidx():
        t = ptext(segs[k][1])
        raw = segs[k][1]
        if i_bab1 is not None and k >= i_bab1:
            # ----- isi BAB: indent 3em + 1.5 spasi -----
            if in_protected(k) or k in generated:
                continue
            if len(t) < 60 or '<w:numPr>' in raw or re.search(r'<w:pStyle w:val="[^"]*(enum|list)', raw):
                continue
            adds = [LINE_15, f'<w:ind w:firstLine="{PARINDENT}"/>']
            if not re.search(r'<w:jc\b', raw):
                adds.append(JC_BOTH)
            segs[k][1] = upsert_ppr(raw, adds)
            n_body += 1
        else:
            # ----- front matter: per section -----
            tn2 = re.sub(r'[\s\W]+', '', t).lower()
            if tn2 in front0 or t.startswith('BAB '):
                sec_prose = bool(PROSE_RE.search(t))
                continue
            if in_protected(k) or k in generated:
                continue
            if len(t) < 60 or '<w:numPr>' in raw or re.search(r'<w:pStyle w:val="[^"]*(enum|list)', raw):
                continue
            if sec_prose:
                if t == t.upper() and any(c.isalpha() for c in t):
                    continue                          # baris judul kapital (blok judul intisari) — biarkan
                adds = [LINE_15, f'<w:ind w:firstLine="{PARINDENT}"/>']
                n_prose += 1
            else:
                adds = [SP_AFTER]
                n_letter += 1
            if not re.search(r'<w:jc\b', raw):
                adds.append(JC_BOTH)
            segs[k][1] = upsert_ppr(raw, adds)

    # ---------- E. (dihapus — kini murni geometri PDF di graft_pdf_layout.py) ----------
    # Dulu: byline abstrak dirata-tengahkan lewat detektor kata (oleh/by, pola NIM,
    # awalan Diajukan/Submitted) dan baris kata kunci/keywords di-strip indentnya.
    # Keduanya digantikan oracle geometri: rata-tengah dibaca dari tepi kiri/titik
    # tengah baris PDF; paragraf yang baris pertamanya MULAI DI MARGIN dibuang
    # firstLine-nya oleh oracle jarak vertikal. Tidak ada detektor kata tersisa.

    log.append(f'front matter prosa    : {n_prose} (indent 3em + 1.5 spasi, tanpa spacing)')
    log.append(f'front matter surat   : {n_letter} (spacing antar paragraf)')
    log.append(f'isi BAB ditata       : {n_body} (indent 3em + 1.5 spasi + justify)')

    doc = ''.join(v for _, v in segs)
    data['word/document.xml'] = doc.encode('utf-8')
    with zipfile.ZipFile(DST, 'w', zipfile.ZIP_DEFLATED) as z:
        for n, b in data.items():
            z.writestr(n, b)

    print(f'[typography] {SRC} -> {DST}')
    for line in log:
        print('  -', line)


def main(argv=None):
    a = sys.argv[1:] if argv is None else argv
    if len(a) < 4:
        sys.exit('pemakaian: typography <in.docx> <out.docx> <dir_toc> <prefix_toc>')
    run(a[0], a[1], a[2], a[3])

if __name__ == '__main__':
    main()
