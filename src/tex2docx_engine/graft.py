r"""
graft_pdf_layout.py — Panen keputusan layout dari PDF hasil pdflatex (page-break oracle)
lalu tanam ke DOCX. PDF = ground truth dari page builder TeX yang asli; DOCX dari make4ht
berisi konten + struktur. Layer ini menyatukan keduanya: setiap batas halaman PDF
dicocokkan ke paragraf DOCX dan diberi page-break.
- Header/footer halaman dideteksi statistik (baris muncul di >=30% halaman) — template-agnostic
- Baris PDF di-strip SEBELUM dipotong (judul rata-tengah punya indentasi spasi di
  mode -layout; memotong dulu menghasilkan anchor sampah seperti 'ha')
- Anchor terlalu pendek tidak dipakai untuk pencocokan 'contains' (mencegah salah
  cocok ke sembarang paragraf)
- Break ditempel ke paragraf target; bila sebelumnya ada rangkaian JUDUL SECTION
  (dari .toc — mis. INTISARI yang dipindah polish ke depan blok judul), break naik
  ke rangkaian itu. Baris nama/NIM tanda tangan bukan judul section → tidak dilewati.
- ALIGNMENT oracle: tex4ht hanya meneruskan \\begin{center} (jadi style 'center');
  deklarasi \\centering di dalam env (intisari/abstract) HILANG di konversi. Maka
  perataan dibaca dari GEOMETRI baris PDF dalam SATUAN TITIK (pdftotext -bbox) —
  grid sel -layout skalanya bergantung ukuran font sehingga judul ber-font besar
  'salah pusat' versi sel padahal fisiknya rata tengah persis. Blok rata tengah =
  titik tengah baris mengelompok & tepi kiri menjauh dari margin; baris ber-celah
  antar-kata lebar (kolom tabel / nomor entri daftar di kanan) memisah blok, jadi
  judul di atas tabel tetap jadi blok sendiri. Bebas bahasa, bebas .cls, bebas kata.
- WEIGHT oracle: bold judul dibaca dari NAMA FONT pdftohtml -xml (weight ada di
  nama PostScript: NimbusRomNo9L-Medi = bold-nya Times URW, *-Bold, CMBX*) —
  mapping style make4ht→docx kerap menjatuhkan bold judul. Bold ditanam pada
  paragraf yang terbukti rata tengah + cocok teksnya (>=60% lebar teks ber-font
  bold).
- VERTICAL-SPACE oracle: \\vskip/\\vspace/\\parskip juga HILANG (LibreOffice menulis
  w:before="0" literal di semua paragraf). Jarak vertikal diukur dari koordinat-y
  pdftotext -bbox: rata baris normal halaman = median jarak antar-baris; selisih
  di atas ambang = jarak yang disengaja -> ditanam sebagai w:spacing/@w:before.
  Baris PERTAMA tiap halaman di-skip (jaraknya ke header/top, bukan antar-paragraf;
  paginasinya sudah ditangani page-break oracle). Jarak sangat besar (>= 40pt)
  dilewati: biasanya area gambar/tabel yang tak punya teks di bbox.
  Bonus dari koordinat-x: paragraf yang baris pertamanya mulai DI MARGIN
  (delta < 15pt) dipastikan tanpa w:firstLine — parindent nyata ~35pt —
  menggantikan detektor kata "kata kunci/keywords" yang lama.
Pemakaian: python3 graft_pdf_layout.py <oracle.pdf> <in.docx> <out.docx> [file.toc]
"""
import os
import re
import subprocess
import sys
import unicodedata
import zipfile
from collections import Counter
def run(PDF, SRC, DST, TOC=""):

    def norm(s):
        s = unicodedata.normalize('NFKC', s)
        return re.sub(r'[\s­\-‐‑…·•.,;:!?"\'“”‘’«»()\[\]]+', '', s).lower()

    # ---------- 1. Anchor batas halaman dari PDF (pdftotext, halaman dipisah \f) ----------
    raw = subprocess.run(['pdftotext', '-layout', PDF, '-'],
                         capture_output=True, text=True, check=True).stdout
    pages = [p for p in raw.split('\f') if p.strip()]
    total = len(pages)

    head_foot = Counter()
    page_lines = []
    for p in pages:
        lines = [l for l in p.split('\n') if l.strip()]
        page_lines.append(lines)
        for l in lines[:2] + lines[-2:]:
            head_foot[norm(l)] += 1
    REPEATED = {t for t, c in head_foot.items() if t and c >= max(3, total * 0.3)}

    def is_chrome(l):
        t = norm(l)
        return (not t) or (t in REPEATED) or re.fullmatch(r'[0-9ivxlcm]+', t)

    anchors = []
    for pno in range(1, total):
        content = [l.strip() for l in page_lines[pno] if not is_chrome(l)]
        first = content[0] if content else ''
        second = content[1][:36] if len(content) > 1 else ''
        anchors.append((pno + 1, norm(first[:48]), norm(first[:24] + second)))

    # ---------- 2. Paragraf DOCX ----------
    zin = zipfile.ZipFile(SRC)
    data = {n: zin.read(n) for n in zin.namelist()}
    zin.close()
    doc = data['word/document.xml'].decode('utf-8')

    para_re = re.compile(r'<w:p\b[^>]*?(?:/>|>.*?</w:p>)', re.S)
    segs, last = [], 0
    for m in para_re.finditer(doc):
        segs.append(['text', doc[last:m.start()]])
        segs.append(['para', m.group(0)])
        last = m.end()
    segs.append(['text', doc[last:]])

    def ptext_raw(raw):
        t = raw.replace('</w:t>', ' </w:t>')
        t = re.sub(r'<[^>]+>', '', t)
        return t

    idxs = [i for i, (k, _) in enumerate(segs) if k == 'para']
    para_texts = {i: norm(ptext_raw(segs[i][1])) for i in idxs}
    para_raw = {i: ptext_raw(segs[i][1]).strip() for i in idxs}

    # judul section dari .toc → satu-satunya paragraf yang boleh dilewati walk-back
    def clean_title(t):
        t = re.sub(r'\\numberline\s*\{[^}]*\}', '', t)
        t = re.sub(r'\\(?:textit|textbf|emph|texttt)\s*\{([^{}]*)\}', r'\1', t)
        t = re.sub(r'\\[a-zA-Z@]+\*?', '', t)
        return re.sub(r'\s+', ' ', t.replace('{', '').replace('}', '')).strip()

    toc_titles = set()
    if TOC and os.path.isfile(TOC):
        txt = open(TOC, encoding='utf-8', errors='ignore').read()
        for m in re.finditer(r'\\contentsline\s*\{(\w+)\}\s*\{(.*?)\}\s*\{([^}]*)\}', txt, re.S):
            t = clean_title(m.group(2))
            if t:
                toc_titles.add(norm(t))

    # ---------- 3. Cocokkan anchor -> paragraf (monoton, maju) ----------
    def add_break(raw):
        if '<w:pageBreakBefore/>' in raw:
            return raw
        br = '<w:pageBreakBefore/>'
        if '<w:pPr' not in raw:
            m = re.search(r'<w:p\b[^>]*>', raw)
            return raw[:m.end()] + '<w:pPr>' + br + '</w:pPr>' + raw[m.end():]
        if '<w:pStyle' in raw:
            m = re.search(r'<w:pStyle[^>]*/>', raw)
            return raw[:m.end()] + br + raw[m.end():]
        m = re.search(r'<w:pPr[^>]*>', raw)
        return raw[:m.end()] + br + raw[m.end():]

    def prev_nonempty(j):
        for q in range(j - 1, -1, -1):
            if para_texts[idxs[q]]:
                return q
        return None

    def section_label(i):
        """Judul section (dari .toc bila tersedia) — BUKAN baris nama/NIM tanda tangan."""
        if toc_titles:
            t = para_texts[i]
            return t in toc_titles or re.sub(r'^bab[ivxlcm]+', '', t) in toc_titles
        t = para_raw[i]                       # fallback tanpa .toc: pendek + kapiral semua
        return len(para_texts[i]) < 40 and any(c.isalpha() for c in t) and t.upper() == t

    def attach(target_j):
        """Naikkan break melewati rangkaian judul section (maks 3 lompatan) tanpa break."""
        j = target_j
        for _ in range(3):
            q = prev_nonempty(j)
            if (q is not None and section_label(idxs[q])
                    and 'pageBreakBefore' not in segs[idxs[q]][1]):
                j = q
            else:
                break
        segs[idxs[j]][1] = add_break(segs[idxs[j]][1])

    n_break = n_mid = n_miss = 0
    missed = []
    cursor = 0
    for page, anchor, anchor2 in anchors:
        if len(anchor) < 8:
            anchor = ''                        # judul 2 baris ("BAB II" saja) → andalkan gabungan
        use2 = (len(anchor2) >= max(16, len(anchor) + 4)) if anchor else (len(anchor2) >= 12)
        if not anchor and not use2:
            n_miss += 1
            missed.append(page)
            continue
        target = None
        for j in range(cursor, len(idxs)):
            t = para_texts[idxs[j]]
            if not t:
                continue
            if ((anchor and (t.startswith(anchor) or (len(anchor) >= 12 and anchor in t)))
                    or (use2 and (t.startswith(anchor2) or anchor2 in t))):
                target = j
                break
        if target is None:
            n_miss += 1
            missed.append(page)
            continue
        tn = para_texts[idxs[target]]
        if (anchor and tn.startswith(anchor)) or (use2 and tn.startswith(anchor2)):
            attach(target)                     # batas = awal paragraf → persis
            n_break += 1
        else:
            q = prev_nonempty(target)          # batas di tengah: hanya selamat bila ada judul section
            if q is not None and section_label(idxs[q]):
                attach(q)
                n_break += 1
            else:
                n_mid += 1
        cursor = target

    # ---------- 4. ALIGNMENT + WEIGHT oracle (geometri baris PDF, satuan TITIK) ----------
    # Grid sel pdftotext -layout skalanya per-ukuran-font (judul ber-font besar =
    # sel lebih lebar) sehingga 'pusat' versi sel menyesatkan: judul DAFTAR TABEL
    # fisiknya rata tengah persis tapi mid-nya jauh dari pusat versi sel. Maka
    # perataan & pemisah tabel diukur di koordinat TITIK xMin/xMax (-bbox) yang
    # bebas skala. Dipakai juga oracle jarak vertikal (bagian 5).
    bbox = subprocess.run(['pdftotext', '-bbox', PDF, '-'],
                          capture_output=True, text=True, check=True).stdout
    bpages, page_w = [], []
    for chunk in re.split(r'<page ', bbox)[1:]:
        page_w.append(float(re.match(r'width="([\d.]+)"', chunk).group(1)))
        ws = re.findall(r'<word xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" '
                        r'yMax="([\d.]+)">([^<]*)</word>', chunk)
        # Baris = kata dengan overlap vertikal >= 40% tinggi yang lebih pendek.
        # (Ambang selisih-y tetap menghancurkan baris ber-font campur: kata italic
        # di tengah teks regular punya yMin beda ~3.5pt dan terlempar jadi 'baris'
        # sendiri — menimbulkan lubang gap palsu di tengah paragraf.)
        words = sorted(((float(a), float(b), float(c), float(d), e)
                        for a, b, c, d, e in ws), key=lambda t: (t[1], t[0]))
        lines = []
        for x0, y0, x1, y1, w in words:
            if lines:
                L = lines[-1]
                ov = min(L['y1'], y1) - max(L['y0'], y0)
                if ov > 0.4 * min(L['y1'] - L['y0'], y1 - y0):
                    L['words'].append((x0, x1, w))
                    L['y0'] = min(L['y0'], y0)
                    L['y1'] = max(L['y1'], y1)
                    continue
            lines.append({'y0': y0, 'y1': y1, 'words': [(x0, x1, w)]})
        for L in lines:                       # urut baca + gap antar-kata nyata
            wsx = sorted(L['words'])
            L['x0'], L['x1'] = wsx[0][0], wsx[-1][1]
            L['t'] = ' '.join(w for _, _, w in wsx)
            L['gap'] = max((b[0] - a[1] for a, b in zip(wsx, wsx[1:])), default=0.0)
        bpages.append(lines)

    # FONT oracle (pdftohtml -xml): weight bold ada di NAMA PostScript font yang
    # dipakai pdflatex menggambar baris — NimbusRomNo9L-Medi (Times-Bold URW),
    # LMRoman*-Bold, CMBX*. Mapping style make4ht->docx kerap menjatuhkan bold
    # judul; bold ditanam ulang dari bukti font PDF. Bebas kata.
    def font_is_bold(family):
        fam = re.sub(r'^[A-Z]{6}\+', '', family)    # buang prefix subset 'HRVNZZ+'
        return bool(re.search(r'(?:^|-)(?:medi(?:um)?|bold|semibold|demibold|'
                              r'extrabold|black|heavy)(?:$|-)|^cm(?:bx|bs)', fam, re.I))

    fpages = []                                     # per halaman: [(y_pt, lebar_pt, bold)]
    fonts = {}
    try:
        hxml = subprocess.run(['pdftohtml', '-xml', '-i', '-stdout', PDF],
                              capture_output=True, text=True, check=True).stdout
        for pno, chunk in enumerate(re.split(r'<page ', hxml)[1:]):
            if pno >= len(page_w):
                break
            scale = float(re.search(r'width="([\d.]+)"', chunk).group(1)) / page_w[pno]
            # id font GLOBAL: <fontspec> hanya dicetak pada halaman pertama
            # kemunculannya — akumulasi lintas halaman, jangan reset per chunk
            for i, f in re.findall(r'<fontspec id="(\d+)"[^>]*family="([^"]*)"',
                                   chunk):
                fonts[int(i)] = font_is_bold(f)
            fpages.append([(float(m.group(1)) / scale, float(m.group(2)) / scale,
                            fonts.get(int(m.group(3)), False))
                           for m in re.finditer(r'<text top="([\d.]+)" left="[\d.]+" '
                                                r'width="([\d.]+)" height="[\d.]+" '
                                                r'font="(\d+)"', chunk)])
    except Exception:
        fpages = []

    # \centering/\begin{center}/\hfill/macro apa pun dieksekusi engine -> baris hasilnya
    # rata tengah SECARA GEOMETRIS (titik). Blok = baris konten berurutan; baris
    # 'tabel/tab' (celah antar-kata >= 40pt: kolom tabel, nomor entri daftar di
    # kanan) MEMISAH blok — judul tepat di atas tabel tetap jadi blok sendiri.
    def set_center(raw):
        if '<w:numPr>' in raw or 'w:val="center"' in raw:
            return raw, False
        raw = re.sub(r'<w:jc [^>]*/>', '', raw)
        raw = re.sub(r'<w:ind [^>]*/>', '', raw)     # center: parindent tidak berlaku
        jc = '<w:jc w:val="center"/>'
        if '<w:pPr' not in raw:
            m = re.search(r'<w:p\b[^>]*>', raw)
            return raw[:m.end()] + '<w:pPr>' + jc + '</w:pPr>' + raw[m.end():], True
        m = re.search(r'<w:rPr[ />]|</w:pPr>', raw)
        pos = m.start() if m else raw.find('</w:pPr>')
        return raw[:pos] + jc + raw[pos:], True

    def set_bold(raw):
        """<w:b/> ke semua run paragraf (weight font PDF terbukti bold)."""
        cnt = [0]

        def repl(m):
            r = m.group(0)
            if '<w:b/>' in r:
                return r
            g = re.search(r'<w:rPr>(<w:rFonts[^>]*/>)?', r)
            if g:
                cnt[0] += 1
                return r[:g.end()] + '<w:b/>' + r[g.end():]
            g = re.match(r'<w:r\b[^>]*>', r)
            cnt[0] += 1
            return r[:g.end()] + '<w:rPr><w:b/></w:rPr>' + r[g.end():]

        new = re.sub(r'<w:r\b[^>]*>(?:(?!</w:r>).)*</w:r>', repl, raw, flags=re.S)
        return new, cnt[0] > 0

    def block_bold(pno, sb):
        """Rasio lebar teks ber-font bold pada blok (koordinat pdftohtml -> bbox)."""
        if pno >= len(fpages) or not fpages[pno]:
            return 0.0
        lo = min(L['y0'] for L in sb) - 6
        hi = max(L['y0'] for L in sb) + 6
        tot = bold = 0.0
        for top, width, isb in fpages[pno]:
            if lo <= top <= hi:
                tot += width
                bold += width if isb else 0.0
        return bold / tot if tot else 0.0

    n_center = n_bold = 0
    centered_log = []
    acur = 0
    # kolom teks GLOBAL (median antar halaman): margin per-halaman dari min/max
    # baris konten mudah bias — halaman daftar/tabel punya entri sampai ke kanan
    # sehingga pusatnya bergeser dan judulnya sendiri tampak 'tidak di tengah'.
    # Halaman badan dominan, mediannya stabil.
    page_cols = [((lambda cl, cr: ((cl + cr) / 2, cr - cl))(
        min(L['x0'] for L in c), max(L['x1'] for L in c)))
        for c in ([L for L in ln if not is_chrome(L['t'])] for ln in bpages)
        if len(c) >= 3]
    page_cols.sort()
    if page_cols:
        col_c, col_w = page_cols[len(page_cols) // 2]
    else:
        col_c, col_w = 0.0, 1.0
    col_l = col_c - col_w / 2
    for pno, lines in enumerate(bpages):
        content = [L for L in lines if not is_chrome(L['t'])]
        if not content:
            continue
        blocks, cur = [], []
        for L in content:
            if L['gap'] >= 40:                      # tabel/tab: akhiri blok, buang baris
                if cur:
                    blocks.append(cur)
                cur = []
                continue
            cur.append(L)
        if cur:
            blocks.append(cur)
        for blk in blocks:
            # pecah sub-blok pada lonjakan mid (ganti kelompok perataan); baris
            # pendek yang diikuti jauh lebih lebar (judul DAFTAR GAMBAR <- entri
            # ber-leader); atau baris lebar yang diikuti baris pendek (blok surat
            # -> judul INTISARI). Kelompok perataan berganti di kedua arah.
            subs, sub, pm, pw = [], [], None, None
            for L in blk:
                w = L['x1'] - L['x0']
                mid = (L['x0'] + L['x1']) / 2
                if sub and ((pm is not None and abs(mid - pm) > 12) or
                            (pw is not None and pw < 0.45 * col_w and w > 1.8 * pw) or
                            (pw is not None and pw > 0.4 * col_w and w < 0.4 * col_w) or
                            (pw is not None and w < 0.45 * pw and w < 0.15 * col_w)):
                    subs.append(sub)
                    sub = []
                sub.append(L)
                pm, pw = mid, w
            if sub:
                subs.append(sub)
            for si, sb in enumerate(subs):
                lefts = [L['x0'] - col_l for L in sb]
                mids = [(L['x0'] + L['x1']) / 2 for L in sb]
                med = sorted(mids)[len(mids) // 2]
                dev = max(abs(m - med) for m in mids)
                var = max(lefts) - min(lefts)
                maxw = max(L['x1'] - L['x0'] for L in sb)
                if dev > 4:
                    continue               # titik tengah tidak mengelompok (justify dsb.)
                if len(sb) == 1:
                    # baris tunggal harus SEMPIT & di pusat kolom. Baris rata
                    # tengah mustahil MULAI di dekat margin (x0 >= pusat - w/2),
                    # maka baris terakhir pendek paragraf rata-kiri yang kebetulan
                    # berakhir di pusat ditolak oleh syarat tepi kirinya.
                    if (maxw > 0.7 * col_w or abs(mids[0] - col_c) > 6
                            or lefts[0] < 0.12 * col_w):
                        continue
                else:
                    # blok ~ selebar kolom (>= 0.85 kolom) biasanya paragraf
                    # justify yang kebetulan mid-nya seimbang — tapi judul panjang
                    # yang benar-benar center presisi (mid == pusat kolom ±3pt di
                    # SEMUA barisnya) tetap boleh
                    tight = max(abs(m - col_c) for m in mids) <= 3
                    if (maxw > 0.85 * col_w and not tight) or abs(med - col_c) > 12:
                        continue
                    # baris di margin (indent kecil, varian sempit) = paragraf biasa
                    if min(lefts) <= 2 and var < 5:
                        continue
                    # rata tengah = tepi kiri bervariasi; blok sempit indent-stabil
                    # (mis. blok dedikasi di halaman judul) boleh bila jauh dari
                    # margin & mid-nya di sekitar pusat kolom
                    ok_narrow = (min(lefts) >= 25 and maxw <= 0.8 * col_w and
                                 abs(med - col_c) <= 12)
                    # judul lebar multi-baris (halaman judul): tepi kirinya nyaris
                    # seragam persis seperti justify, tapi mid-nya presisi di pusat
                    if not (var >= 5 or ok_narrow or (tight and min(lefts) > 2)):
                        continue
                # cocokkan: teks sub-blok, lalu GABUNGAN dengan sub-blok berikutnya
                # ("BAB II" + "TINJAUAN PUSTAKA" bisa terpecah oleh aturan lebar)
                tgt, blktext = None, ''
                nsub1 = len(norm(''.join(L['t'] for L in subs[si])))
                for take in (1, 2, 3):
                    if si + take > len(subs):
                        break
                    if take > 1 and nsub1 > 24:
                        break   # gabungan hanya untuk sub sangat pendek ('BAB II')
                    blktext = norm(''.join(L['t'] for s2 in subs[si:si + take]
                                           for L in s2))
                    if len(blktext) < 2:
                        continue           # 'by' (byline EN) tetap tertangkap via geometri
                    # jendela maju terbatas: urutan docx bisa beda dari PDF
                    # (front matter dipindah polish), tapi tidak ratusan paragraf —
                    # scan tanpa batas membuat satu salah-match melompatkan kursor
                    # jauh & meracuni semua halaman berikutnya
                    # pass 1 persis boleh jauh (judul asli bisa terdorong ratusan
                    # paragraf oleh entri daftar isi buatan polish); pass 2 prefix
                    # harus dekat — fuzzy, mudah salah sasaran. Pass 0 mundur:
                    # polish memindah INTISARI/ABSTRAK ke depan sehingga targetnya
                    # di belakang kursor — persis-sama aman dicari ke belakang.
                    for j in range(acur, min(acur + 600, len(idxs))):
                        if para_texts[idxs[j]] == blktext:
                            tgt = j
                            break
                    if tgt is None:
                        for j in range(acur - 1, -1, -1):
                            if para_texts[idxs[j]] == blktext:
                                tgt = j
                                break
                    if tgt is None:
                        for j in range(acur, min(acur + 80, len(idxs))):
                            t = para_texts[idxs[j]]
                            if t and len(blktext) >= 12 and (
                                    t.startswith(blktext[:16]) or
                                    blktext.startswith(t[:16])):
                                tgt = j
                                break
                    if tgt is not None:
                        break
                if tgt is None:
                    continue
                # make4ht memetakan \chapter ke paragraf bernomor (numPr) — pada
                # judul itu murni artefak struktur; buang agar bisa ditanam jc
                # (dan nomor tak di-drop numPr asli Word tak ikut tampil salah)
                segs[idxs[tgt]][1] = re.sub(
                    r'<w:numPr>.*?</w:numPr>|<w:numPr\s*/>', '',
                    segs[idxs[tgt]][1], flags=re.S)
                new, changed = set_center(segs[idxs[tgt]][1])
                if changed:
                    segs[idxs[tgt]][1] = new
                    n_center += 1
                    if len(centered_log) < 12:
                        centered_log.append(blktext[:34])
                if block_bold(pno, sb) >= 0.6:
                    new, ch = set_bold(segs[idxs[tgt]][1])
                    if ch:
                        segs[idxs[tgt]][1] = new
                        n_bold += 1
                acur = max(acur, tgt)   # match mundur tak menarik kursor mundur

    # ---------- 5. VERTICAL-SPACE oracle (jarak antar-baris dari koordinat-y PDF) ----------
    VGAP_MIN_PT, VGAP_MAX_PT = 6.0, 40.0

    def upsert_attr(raw, attr, tw):
        m = re.search(r'<w:spacing\b[^>]*/>', raw)
        if m:                                   # ubah nilai yang ada (mis. "0" dari LO)
            s = m.group(0)
            if f'{attr}="' in s:
                s = re.sub(attr + r'="\d+"', f'{attr}="{tw}"', s)
            else:
                s = s.replace('/>', f' {attr}="{tw}"/>')
            return raw[:m.start()] + s + raw[m.end():]
        el = f'<w:spacing {attr}="{tw}"/>'
        if '<w:pPr' not in raw:
            m = re.search(r'<w:p\b[^>]*>', raw)
            return raw[:m.end()] + '<w:pPr>' + el + '</w:pPr>' + raw[m.end():]
        # urutan skema pPr: spacing setelah pStyle/pageBreakBefore/tabs, sebelum ind/jc
        pos = None
        for pat in (r'<w:pStyle[^>]*/>', r'<w:pageBreakBefore[^>]*/>',
                    r'<w:tabs\b.*?</w:tabs>', r'<w:pPr[^>]*>'):
            for fm in re.finditer(pat, raw, re.S):
                if pos is None or fm.end() > pos:
                    pos = fm.end()
        return raw[:pos] + el + raw[pos:]

    # paragraf pertama di baris pertama sebuah tabel: XML antar-paragraf berisi <w:tbl>
    # yang belum tertutup. w:before di paragraf dalam sel TIDAK memberi jarak di atas
    # tabel di Word — jarak itu ditanam sebagai w:after paragraf terakhir sebelum tabel.
    tbl_first = {i for i in range(1, len(segs)) if segs[i][0] == 'para'
                 and '<w:tbl>' in segs[i - 1][1] and '</w:tbl>' not in segs[i - 1][1]}

    def strip_firstline(raw):
        """Buang w:firstLine/hanging dari w:ind (baris pertama paragraf terbukti
        mulai DI MARGIN menurut PDF — mis. baris kata kunci, bukan parindent)."""
        m = re.search(r'<w:ind\b[^>]*/>', raw)
        if not m or 'firstLine' not in m.group(0):
            return raw
        keep = [a for a in re.findall(r'w:\w+="[^"]*"', m.group(0))
                if not a.startswith('w:firstLine') and not a.startswith('w:hanging')]
        el = f'<w:ind {" ".join(keep)}/>' if keep else ''
        return raw[:m.start()] + el + raw[m.end():]

    # (geometri per-baris bpages sudah diurai di bagian 4 — dipakai ulang di sini)
    n_vgap = n_vbig = 0
    vgap_log = []
    vcur = 0
    for lines in bpages:
        content = [L for L in lines if not is_chrome(L['t'])]
        if len(content) < 2:
            continue
        page_left = min(L['x0'] for L in content)   # margin kiri teks halaman ini
        gaps = [b['y0'] - a['y1'] for a, b in zip(content, content[1:])]
        normal = sorted(gaps)[len(gaps) // 2]   # rata baris normal halaman ini
        for a, b, g in zip(content, content[1:], gaps):
            extra = g - normal
            if extra < VGAP_MIN_PT:
                continue
            if extra >= VGAP_MAX_PT:
                n_vbig += 1                     # area gambar/tabel tanpa teks — tak terukur
                continue
            ln = norm(b['t'])
            if len(ln) < 2:
                continue                       # 'by' (abstract EN) tetap tertangkap via equality
            tgt = None
            for j in range(vcur, len(idxs)):
                t = para_texts[idxs[j]]
                # dua arah: baris PDF bisa melebihi satu paragraf DOCX (sel tabel:
                # 'Pembimbing Utama : Dr. ...' = 3 sel) — cukup sel pertamanya awalan
                if t and (t == ln or (len(ln) >= 12 and
                                      (t.startswith(ln[:16]) or
                                       (len(t) >= 8 and ln.startswith(t[:16]))))):
                    tgt = j
                    break
            if tgt is None:
                continue
            tw = int(round(extra * 20))
            if idxs[tgt] in tbl_first:
                q = prev_nonempty(tgt)          # jarak sebelum tabel -> w:after sebelumnya
                if q is not None:
                    segs[idxs[q]][1] = upsert_attr(segs[idxs[q]][1], 'w:after', tw)
            else:
                segs[idxs[tgt]][1] = upsert_attr(segs[idxs[tgt]][1], 'w:before', tw)
                # gap PDF = SATU jarak (semantik \vskip): milik 'before' paragraf
                # berikut — nol-kan 'after' pendahulunya agar tidak dobel (mis.
                # after=240 dari aturan surat typography + before terukur)
                q0 = prev_nonempty(tgt)
                if q0 is not None and q0 != tgt:
                    segs[idxs[q0]][1] = upsert_attr(segs[idxs[q0]][1], 'w:after', 0)
                # baris pertama paragraf mulai DI MARGIN (paragraf parindent nyata
                # mulai ~35pt dari margin) → firstLine tidak berlaku (mis. kata kunci)
                if b['x0'] - page_left < 15:
                    segs[idxs[tgt]][1] = strip_firstline(segs[idxs[tgt]][1])
            n_vgap += 1
            if len(vgap_log) < 10:
                vgap_log.append(f'{ln[:18]}(+{extra:.0f}pt)')
            vcur = tgt

    doc = ''.join(v for _, v in segs)
    data['word/document.xml'] = doc.encode('utf-8')
    with zipfile.ZipFile(DST, 'w', zipfile.ZIP_DEFLATED) as z:
        for n, b in data.items():
            z.writestr(n, b)

    print(f'[graft] PDF oracle: {total} halaman -> {total-1} batas')
    print(f'[graft] page-break ditanam : {n_break}')
    print(f'[graft] batas mid-paragraf : {n_mid} (dilewati — paragraf tidak dipotong)')
    print(f'[graft] tidak ketemu       : {n_miss} (halaman {missed})')
    print(f'[graft] rata tengah (PDF)  : {n_center} paragraf')
    print(f'[graft] judul bold (font)  : {n_bold} paragraf')
    if centered_log:
        print(f'[graft] contoh             : {", ".join(centered_log)}')
    print(f'[graft] jarak vertikal PDF : {n_vgap} paragraf (w:before)')
    if vgap_log:
        print(f'[graft] contoh jarak       : {", ".join(vgap_log)}')
    print(f'[graft] jarak >= 40pt      : {n_vbig} dilewati (area gambar/tabel)')


def main(argv=None):
    a = sys.argv[1:] if argv is None else argv
    if len(a) < 3:
        sys.exit('pemakaian: graft <oracle.pdf> <in.docx> <out.docx> [file.toc]')
    run(a[0], a[1], a[2], a[3] if len(a) > 3 else '')

if __name__ == '__main__':
    main()
