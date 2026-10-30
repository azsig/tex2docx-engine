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
    # Kriteria sukses tahap ini = <main>.odt terbentuk.
    odt = os.path.join(workdir, f'{main}.odt')
    before = os.path.getmtime(odt) if os.path.isfile(odt) else 0
    env = dict(os.environ)
    cfg = os.path.join(PKG_DIR, 't2dcfg.cfg')
    if os.path.isfile(cfg):
        env['TEXINPUTS'] = PKG_DIR + os.pathsep + env.get('TEXINPUTS', '')
        mk = [_need('make4ht'), '-f', 'odt', '-c', 't2dcfg', f'{main}.tex']
    else:  # defensif: instalasi tanpa package-data — jalankan tanpa marker
        mk = [_need('make4ht'), '-f', 'odt', f'{main}.tex']
        print('[pipeline] t2dcfg.cfg tidak ada di paket — mode tanpa marker native')
    r = subprocess.run(mk, cwd=workdir, env=env, capture_output=True, text=True)
    if not (os.path.isfile(odt) and os.path.getmtime(odt) > before):
        tail = '\n'.join((r.stderr or r.stdout).splitlines()[-12:])
        raise SystemExit(f'[pipeline] gagal: make4ht -f odt {main}.tex\n{tail}')
    if r.returncode != 0:
        print(f'[pipeline] make4ht exit={r.returncode} '
              f'(error recoverable — {main}.odt tetap dihasilkan)')

    # ---------- 2. perbaiki ODT + konversi ke DOCX via LibreOffice ----------
    odt_fixed = os.path.join(tmp, f'{main}_fixed.odt')
    fix_odt.run(odt, odt_fixed)
    # profil terpisah agar bisa jalan paralel / tanpa X
    profile = tempfile.mkdtemp(prefix='lo_tex2docx_')
    soffice = _need('soffice')
    _run([soffice, '-env:UserInstallation=file://' + profile,
          '--headless', '--convert-to', 'docx', '--outdir', tmp, odt_fixed], tmp)
    raw = os.path.join(tmp, f'{main}_fixed.docx')
    if not os.path.isfile(raw):
        raise SystemExit('[pipeline] LibreOffice tidak menghasilkan DOCX')

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
