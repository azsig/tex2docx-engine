"""
pipeline — Orkestrator lengkap LaTeX -> DOCX.

    pdflatex (bila PDF belum ada)  ->  PDF oracle + file aux (.toc/.lot/.lof/.nls)
    make4ht -f odt -c t2dcfg       ->  konten + struktur + MARKER semantik native
    fix_odt + LibreOffice          ->  DOCX mentah (marker sudah jadi style)
    polish -> typography -> graft  ->  DOCX final (paginasi dipanen dari PDF)

Dua mode (default HIBRIDA):
- pdf='auto'  : pdflatex + oracle jalan bila toolchain ada; spacing/perataan
                NATIVE (marker t2dcfg) menang, oracle hanya paginasi + celah
                yang tak bertanda (\addvspace, \vfill, \\[X] intra-paragraf).
- pdf='never' (--no-pdf): tanpa pdflatex & tanpa graft sama sekali — untuk
                mesin tanpa toolchain PDF. Break halaman hanya dari marker
                \newpage + awal bab (tex4ht). Output tetap polos-valid.

Setiap tahap memakai engine asli (TeX/LibreOffice) lalu menegakkannya dengan
geometri PDF — tidak ada pengetahuan per-template di sini.
"""
import os
import shutil
import subprocess
import sys
import tempfile

from . import fix_odt, graft, polish, typography

# t2dcfg.cfg dikirim dalam paket; make4ht menemukannya lewat TEXINPUTS.
PKG_DIR = os.path.dirname(os.path.abspath(__file__))


def _need(cmd):
    path = shutil.which(cmd)
    if not path:
        raise SystemExit(f'[pipeline] {cmd} tidak ditemukan di PATH — install dulu')
    return path


def _run(cmd, cwd):
    r = subprocess.run(cmd, cwd=cwd, capture_output=True, text=True)
    if r.returncode != 0:
        tail = '\n'.join((r.stderr or r.stdout).splitlines()[-12:])
        raise SystemExit(f'[pipeline] gagal: {" ".join(cmd)}\n{tail}')
    return r


def _odt_block_count(odt):
    """Jumlah <text:p di content.xml ODT — kunci validasi kelengkapan LO."""
    import zipfile as _zf
    with _zf.ZipFile(odt) as z:
        return z.read('content.xml').count(b'<text:p')


def _docx_block_count(docx):
    import zipfile as _zf
    with _zf.ZipFile(docx) as z:
        return z.read('word/document.xml').count(b'<w:p>') \
            + z.read('word/document.xml').count(b'<w:p ')


def _odt_content_parses(odt):
    """True bila content.xml ODT lolos parse XML strict.
    make4ht memfilter .4oo lewat domfilter (LuaXML); XML tak seimbang ia
    REPARSE via DOM-HTML dan MENYEMBUHKAN — content.xml hasilnya valid walau
    log tetap menyebut 'Unbalanced Tag'. Validitas parse = bukti
    penyembuhan itu benar-benar terjadi (bukan sekadar absennya warning)."""
    import xml.etree.ElementTree as _ET
    import zipfile as _zf
    try:
        with _zf.ZipFile(odt) as z:
            _ET.fromstring(z.read('content.xml'))
        return True
    except (OSError, _zf.BadZipFile, _ET.ParseError):
        return False


def _bytes_text(data):
    """Teks bersih (tag dikupas, entitas dinormalkan, whitespace TERMASUK
    nbsp dibuang) dari bytes XML — pembanding teks lintas format ODT/DOCX.
    NBSP wajib ikut hilang: ODT menulis \xa0, DOCX kadang &#160;."""
    import html as _html
    import re as _re
    txt = _re.sub(rb'<[^>]+>', b'', data)
    return _re.sub(r'\s+', '', _html.unescape(txt.decode('utf-8', 'replace')))


def _odt_text_len(odt):
    """Jumlah karakter non-spasi konten content.xml — proxy kelengkapan
    konten (tag & marker dikupas mentah; teks dalam objek Math ikut terhitung)."""
    import zipfile as _zf
    with _zf.ZipFile(odt) as z:
        return len(_bytes_text(z.read('content.xml')))


def _odt_tail(odt, n=80):
    """Sidik jari UJUNG dokumen: teks (non-spasi) PARAGRAF non-tabel
    terakhir, dipotong n karakter ekor. LibreOffice mempertahankan urutan &
    isi paragraf flow — abort-nya memotong ekor (skripsi uji: entri pustaka
    [2..25] hilang padahal 89% blok & 97% teks global selamat). Teks tabel
    TIDAK dipakai: konkatenasi sel antar format tak stabil urutannya, dan
    teks antar-paragraf bisa diselingi tabel di DOCX."""
    import re as _re
    import xml.etree.ElementTree as _ET
    import zipfile as _zf
    T = 'urn:oasis:names:tc:opendocument:xmlns:text:1.0'
    TAB = 'urn:oasis:names:tc:opendocument:xmlns:table:1.0'
    with _zf.ZipFile(odt) as z:
        root = _ET.fromstring(z.read('content.xml'))
    parent = {c: p for p in root.iter() for c in p}

    def in_table(el):
        cur = el
        while cur is not None:
            if cur.tag.startswith(f'{{{TAB}}}table'):
                return True
            cur = parent.get(cur)
        return False

    for p in reversed(list(root.iter(f'{{{T}}}p'))):
        if in_table(p):
            continue
        t = _re.sub(r'\s+', '', ''.join(p.itertext()))
        if t:
            return t[-n:]
    return ''


def _docx_text(docx):
    import zipfile as _zf
    with _zf.ZipFile(docx) as z:
        return _bytes_text(z.read('word/document.xml'))


def _odt_complete(odt, pdf_len):
    """ODT memuat SELURUH dokumen? Proxy: rasio karakter konten ODT vs teks
    PDF oracle. Terukur di skripsi uji: hasil terpotong 27-55%, utuh ~1.0 —
    ambang 0.6 memberi margin dua arah. `pdf_len` kosong/None (mode --no-pdf,
    atau PDF tanpa lapisan teks) -> tak bisa menilai -> anggap lengkap."""
    if not pdf_len or pdf_len < 200:
        return True
    return _odt_text_len(odt) >= 0.6 * pdf_len


def _pdf_text_len(pdf, pdftotext):
    """Panjang teks non-spasi PDF oracle (None bila tak terekstrak)."""
    import re as _re
    r = subprocess.run([pdftotext, pdf, '-'], capture_output=True, text=True)
    if r.returncode != 0:
        return None
    return len(_re.sub(r'\s+', '', r.stdout or ''))


def _soffice_docx(soffice, odt_fixed, tmp, main, want_blocks):
    """LibreOffice ODT->DOCX TIDAK DETERMINISTIK untuk dokumen besar ber-objek
    Math ter-embed ratusan (skripsi uji: 912 objek): import kadang abort
    diam-diam di tengah — exit 0 tapi DOCX terpotong. Dua syarat lolos:
    jumlah blok >= 50% ODT (abort besar, ~75% hilang) DAN ekor teks ODT masih
    ada di DOCX (abort kecil di ujung — skripsi uji: 89% blok & 97% teks
    selamat tapi entri pustaka [2..25] hilang). Gagal -> profil segar, maks 3."""
    import tempfile as _tf
    want_tail = _odt_tail(odt_fixed)
    for attempt in range(1, 4):
        profile = _tf.mkdtemp(prefix='lo_tex2docx_')
        _run([soffice, '-env:UserInstallation=file://' + profile,
              '--headless', '--convert-to', 'docx', '--outdir', tmp, odt_fixed], tmp)
        raw = os.path.join(tmp, os.path.basename(odt_fixed).rsplit('.', 1)[0] + '.docx')
        if os.path.isfile(raw):
            got = _docx_block_count(raw)
            if got >= 0.5 * want_blocks and want_tail in _docx_text(raw):
                return raw
            print(f'[pipeline] LibreOffice keluaran terpotong ({got}/{want_blocks} '
                  f'blok, ekor dokumen hilang, percobaan {attempt}) — ulangi')
        else:
            print(f'[pipeline] LibreOffice tanpa keluaran (percobaan {attempt}) — ulangi')
    raise SystemExit('[pipeline] LibreOffice 3x menghasilkan DOCX terpotong')


def _clean_t4ht_debris(workdir, main):
    """Buang intermediate t4ht/xtpisks milik <main> SEBELUM percobaan ulang
    make4ht. Hanya file yang PASTI buatan engine (pola nama ketat) — file
    gambar user (mis. fig1.svg) tak tersentuh. Debu basi dari run sebelumnya
    terbukti memicu konversi gagal-sunyi pada dokumen besar."""
    import glob as _glob
    import re as _re
    base = _re.escape(main)
    pats = (f'{main}.4oo', f'{main}.4ct', f'{main}.4tc', f'{main}.4as',
            f'{main}.idv', f'{main}.xref', f'{main}.lg', f'{main}.dvi',
            f'{main}.tmp', f'{main}.odt')
    rx = _re.compile(
        f'^(?:{base}(?:-m)?\\d+x?(?:\\.svg)?\\.(?:svg|4og|tmp)|{base}\\.tmp)$')
    names = set()
    for p in pats:
        names.update(_glob.glob(os.path.join(workdir, p)))
    for f in os.listdir(workdir):
        if rx.fullmatch(f):
            names.add(os.path.join(workdir, f))
    for f in names:
        try:
            os.remove(f)
        except OSError:
            pass


def _prepare_outer_markers(workdir, mkout=''):
    """tex4ht menolak MEMBUAT file marker `.4og` di luar workdir — path `..`
    diblokir keamanannya — padahal `\\includegraphics{../dir/gambar.png}`
    sah di LaTeX (pdflatex menerimanya). Akibatnya tex4ht abort di halaman
    gambar itu ("Can't find/open file `../...'", error ini TIDAK diteruskan
    ke stdout make4ht jadi tak bisa dideteksi reaktif), sisa dokumen hilang
    atau XML-nya rusak (skripsi uji: mati di Gambar 5.7 -> BAB V-VI
    terpotong). Solusi dua jalur, sama-sama pre-create marker 0-byte:
    - proaktif: scan \\includegraphics{../...} di semua .tex workdir;
    - reaktif: parse error dari mkout (kalau kelihatan).
    Terverifikasi: tex4ht exit 1 di halaman 290 -> exit 0 sampai 329."""
    import re as _re
    rels = set(m.group(1) for m in
               _re.finditer(r"Can't find/open file `([^']+)'", mkout or ''))
    rx_arg = _re.compile(r'\\includegraphics\s*(?:\[[^\]]*\])?\s*\{([^}]+)\}')
    for root, _dirs, files in os.walk(workdir):
        if '.git' in root or os.sep + '.' in root[len(workdir):]:
            continue
        for name in files:
            if not name.endswith('.tex'):
                continue
            try:
                data = open(os.path.join(root, name), encoding='utf-8',
                            errors='replace').read()
            except OSError:
                continue
            for m in rx_arg.finditer(data):
                arg = m.group(1).strip()
                if arg.startswith('../') or '/../' in arg:
                    rels.add(os.path.relpath(
                        os.path.normpath(os.path.join(root, arg)), workdir))
    touched = []
    for rel in sorted(rels):
        target = os.path.normpath(os.path.join(workdir, rel))
        outer = (target.startswith(os.pardir + os.sep)
                 or os.sep + os.pardir + os.sep in target)
        if not outer:
            continue  # di dalam workdir — biarkan tex4ht yang mengurus
        base, ext = os.path.splitext(target)
        exts = [ext] if ext else ['.png', '.jpg', '.jpeg', '.pdf', '.eps']
        for e in exts:  # marker .4og menempel pada nama file GRAFIK, bukan
            cand = base + e + '.4og'   # meng-create grafik itu sendiri
            try:
                os.makedirs(os.path.dirname(cand), exist_ok=True)
                if not os.path.exists(cand):
                    open(cand, 'a').close()
                    touched.append(os.path.relpath(cand, workdir))
            except OSError:
                pass
    return touched


def _make4ht_odt(workdir, main, mk, env, pdf_len=None):
    r"""Jalankan make4ht sampai <main>.odt segar & terbukti UTUH.

    Keputusan lolos/tidak menilai ODT HASIL, bukan sekadar teks log:
    - ODT tak terbentuk (gagal kompilasi sesungguhnya) -> ulang.
    - "Can't find/open file `../...'": tex4ht abort di \includegraphics luar
      workdir — ODT bisa terbentuk seimbang tapi isinya berhenti di halaman
      gambar itu (skripsi uji: berhenti di Gambar 5.7, 76% dokumen).
      _prepare_outer_markers menambal ini antar percobaan -> ulang.
    - 'Unbalanced Tag' di log: bila content.xml ODT TIDAK lolos parse strict
      (luka struktural di tengah dokumen — dulu memotong 748/1625 paragraf)
      -> ulang. Bila lolos parse, domfilter make4ht sudah MENYEMBUHKAN XML
      itu via reparse HTML-DOM (kasus terverifikasi: tag </text:p> liar dari
      "Missing }" \end{thebibliography} kronis, luka di UJUNG dokumen) ->
      terima, karena kontennya utuh.
    - Pemeriksa terakhir (mode hibrida, bila `pdf_len` ada): rasio karakter
      konten ODT vs teks PDF oracle >= 0.6 — menangkap pemotongan DIAM-DIAM
      apa pun penyebabnya (ODT "seimbang" tapi pendek). Mode --no-pdf tak
      punya oracle -> pemeriksa ini dilewati.
    Antar percobaan, debu t4ht dibersihkan agar state deterministik."""
    odt = os.path.join(workdir, f'{main}.odt')
    last_tail = ''
    for attempt in range(1, 4):
        if attempt > 1:
            _clean_t4ht_debris(workdir, main)
        before = os.path.getmtime(odt) if os.path.isfile(odt) else 0
        r = subprocess.run(mk, cwd=workdir, env=env, capture_output=True, text=True)
        mkout = (r.stdout or '') + (r.stderr or '')
        last_tail = '\n'.join(mkout.splitlines()[-12:])
        fresh = os.path.isfile(odt) and os.path.getmtime(odt) > before
        if not fresh:
            bad = 'ODT tidak terbentuk'
        elif "Can't find/open file" in mkout:
            bad = 'tex4ht abort pada grafik (lihat baris marker di atas)'
        elif 'Unbalanced Tag' in mkout and not _odt_content_parses(odt):
            bad = ('XML tak seimbang & tidak tersembuhkan ('
                   + next((l.strip() for l in mkout.splitlines()
                           if 'Unbalanced Tag' in l), '') + ')')
        elif not _odt_complete(odt, pdf_len):
            bad = ('konten ODT jauh lebih pendek dari PDF oracle — '
                   'kemungkinan terpotong diam-diam')
        else:
            if r.returncode != 0:
                print(f'[pipeline] make4ht exit={r.returncode} '
                      f'(error recoverable — {main}.odt tetap dihasilkan)')
            if 'Unbalanced Tag' in mkout:
                print('[pipeline] t4ht: XML tak seimbang di titik jinak '
                      '(ujung dokumen) — make4ht menyembuhkannya, content.xml '
                      'valid & lengkap — lanjut')
            return odt
        # `\includegraphics{../dir/x.png}` -> tex4ht tak bisa menulis marker
        # .4og di luar workdir -> abort di halaman itu. Pre-create marker,
        # percobaan berikutnya lewat.
        for rel in _prepare_outer_markers(workdir, mkout):
            print(f'[pipeline] grafik di luar workdir: {rel} — marker .4og '
                  f'dibuat agar tex4ht bisa lanjut')
        print(f'[pipeline] make4ht percobaan {attempt}: {bad} — '
              + ('ulangi dengan state bersih' if attempt < 3 else 'menyerah'))
    raise SystemExit(
        f'[pipeline] make4ht 3x gagal menghasilkan ODT utuh untuk {main}.tex\n'
        f'{last_tail}')


def _ensure_pdf(workdir, main):
    """PDF = oracle; butuh >= 2 run pdflatex agar .toc/.lot/.lof terisi.
    Bila ada .bib (dan .bbl belum ada), jalankan bibtex di antaranya."""
    pdf = os.path.join(workdir, f'{main}.pdf')
    tex = os.path.join(workdir, f'{main}.tex')
    if os.path.isfile(pdf):
        return pdf
    if not os.path.isfile(tex):
        raise SystemExit(f'[pipeline] {tex} tidak ada')
    pdflatex = _need('pdflatex')
    _run([pdflatex, '-interaction=nonstopmode', f'{main}.tex'], workdir)
    if not os.path.isfile(os.path.join(workdir, f'{main}.bbl')):
        for bib in ('{m}.bib'.format(m=main), 'references.bib'):
            if os.path.isfile(os.path.join(workdir, bib)):
                _run([_need('bibtex'), main], workdir)
                break
    for _ in range(2):
        _run([pdflatex, '-interaction=nonstopmode', f'{main}.tex'], workdir)
    if not os.path.isfile(pdf):
        raise SystemExit(f'[pipeline] pdflatex tidak menghasilkan {main}.pdf')
    return pdf


def convert(workdir, main='main', out=None, font='Times New Roman',
            keep=False, quiet=False, pdf='auto'):
    """Konversi penuh. `workdir` = folder berisi <main>.tex (dan PDF bila sudah
    dikompilasi). `pdf` = 'auto' (hibrida, default) | 'never' (--no-pdf).
    Mengembalikan path DOCX final."""
    workdir = os.path.abspath(workdir)
    out = os.path.abspath(out) if out else os.path.join(workdir, f'{main}_engine.docx')
    flat = main.replace('/', '_').replace('\\', '_')      # main boleh 'bab/main'
    tmp = os.path.join(workdir, f'.{flat}_tex2docx_tmp')
    os.makedirs(tmp, exist_ok=True)
    try:
        return _convert(workdir, main, out, font, tmp, keep, pdf)
    finally:
        if not keep:
            shutil.rmtree(tmp, ignore_errors=True)


def _convert(workdir, main, out, font, tmp, keep, pdf='auto'):
    pdf_path = _ensure_pdf(workdir, main) if pdf == 'auto' else None
    if pdf_path:
        _need('pdftotext')

    # ---------- 1. make4ht: konten via engine TeX asli + marker native ----------
    # make4ht kerap exit != 0 walau ODT tetap dihasilkan lengkap (error htlatex
    # yang recoverable — mis. \insertchapterspace undefined di cls custom).
    # Kriteria sukses tahap ini = <main>.odt terbentuk, parse-valid, dan
    # (mode hibrida) kontennya tak pendek vs oracle PDF.
    for rel in _prepare_outer_markers(workdir):
        print(f'[pipeline] grafik di luar workdir: {rel} — marker .4og '
              f'dibuat agar tex4ht tidak abort di halaman gambar itu')
    pdf_len = (_pdf_text_len(pdf_path, _need('pdftotext'))
               if pdf_path else None)
    env = dict(os.environ)
    cfg = os.path.join(PKG_DIR, 't2dcfg.cfg')
    if os.path.isfile(cfg):
        env['TEXINPUTS'] = PKG_DIR + os.pathsep + env.get('TEXINPUTS', '')
        mk = [_need('make4ht'), '-f', 'odt', '-c', 't2dcfg', f'{main}.tex']
    else:  # defensif: instalasi tanpa package-data — jalankan tanpa marker
        mk = [_need('make4ht'), '-f', 'odt', f'{main}.tex']
        print('[pipeline] t2dcfg.cfg tidak ada di paket — mode tanpa marker native')
    odt = _make4ht_odt(workdir, main, mk, env, pdf_len=pdf_len)

    # ---------- 2. perbaiki ODT + konversi ke DOCX via LibreOffice ----------
    odt_fixed = os.path.join(tmp, f'{main}_fixed.odt')
    fix_odt.run(odt, odt_fixed)
    raw = _soffice_docx(_need('soffice'), odt_fixed, tmp, main,
                        _odt_block_count(odt_fixed))

    # ---------- 3. tiga layer post-processing ----------
    p1 = os.path.join(tmp, 'p1.docx')
    p2 = os.path.join(tmp, 'p2.docx')
    polish.run(raw, p1, font)
    typography.run(p1, p2, workdir, main)
    if not pdf_path:  # mode --no-pdf: akhiri di sini, tanpa graft
        shutil.copyfile(p2, out)
        print('[pipeline] mode --no-pdf: tanpa oracle PDF — paginasi hanya '
              'dari marker \\newpage + awal bab')
        return out
    toc = os.path.join(workdir, f'{main}.toc')
    graft.run(pdf_path, p2, out, toc if os.path.isfile(toc) else '')
    return out


def main(argv=None):
    a = sys.argv[1:] if argv is None else argv
    import argparse
    ap = argparse.ArgumentParser(
        prog='tex2docx',
        description='LaTeX -> DOCX via engine TeX asli + marker semantik native '
                    '+ oracle layout PDF (hibrida)')
    ap.add_argument('workdir', help='folder berisi <main>.tex')
    ap.add_argument('--main', default='main', help='nama dasar file .tex (tanpa ekstensi)')
    ap.add_argument('--out', default=None, help='path DOCX output')
    ap.add_argument('--font', default='Times New Roman', help='font keluaran')
    ap.add_argument('--keep', action='store_true', help='pertahankan file antara')
    ap.add_argument('--no-pdf', action='store_true',
                    help='tanpa pdflatex/oracle — break halaman hanya via marker '
                         '\\newpage + awal bab')
    args = ap.parse_args(a)
    path = convert(args.workdir, args.main, args.out, args.font, args.keep,
                   pdf='never' if args.no_pdf else 'auto')
    print(f'[pipeline] selesai: {path}')


if __name__ == '__main__':
    main()
