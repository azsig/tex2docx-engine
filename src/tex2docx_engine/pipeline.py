"""
pipeline — Orkestrator lengkap LaTeX -> DOCX.

    pdflatex (bila PDF belum ada)  ->  PDF oracle + file aux (.toc/.lot/.lof/.nls)
    make4ht -f odt                 ->  konten + struktur (linearisasi teks asli)
    fix_odt + LibreOffice          ->  DOCX mentah
    polish -> typography -> graft  ->  DOCX final (layout dipanen dari PDF)

Setiap tahap memakai engine asli (TeX/LibreOffice) lalu menegakkannya dengan
geometri PDF — tidak ada pengetahuan per-template di sini.
"""
import os
import shutil
import subprocess
import sys
import tempfile

from . import fix_odt, graft, polish, typography


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
    """PDF = oracle; butuh >= 2 run pdflatex agar .toc/.lot/.lof terisi."""
    pdf = os.path.join(workdir, f'{main}.pdf')
    tex = os.path.join(workdir, f'{main}.tex')
    if os.path.isfile(pdf):
        return pdf
    if not os.path.isfile(tex):
        raise SystemExit(f'[pipeline] {tex} tidak ada')
    pdflatex = _need('pdflatex')
    for _ in range(2):
        _run([pdflatex, '-interaction=nonstopmode', f'{main}.tex'], workdir)
    if not os.path.isfile(pdf):
        raise SystemExit(f'[pipeline] pdflatex tidak menghasilkan {main}.pdf')
    return pdf


def convert(workdir, main='main', out=None, font='Times New Roman',
            keep=False, quiet=False):
    """Konversi penuh. `workdir` = folder berisi <main>.tex (dan PDF bila sudah
    dikompilasi). Mengembalikan path DOCX final."""
    workdir = os.path.abspath(workdir)
    out = os.path.abspath(out) if out else os.path.join(workdir, f'{main}_engine.docx')
    flat = main.replace('/', '_').replace('\\', '_')      # main boleh 'bab/main'
    tmp = os.path.join(workdir, f'.{flat}_tex2docx_tmp')
    os.makedirs(tmp, exist_ok=True)
    try:
        return _convert(workdir, main, out, font, tmp, keep)
    finally:
        if not keep:
            shutil.rmtree(tmp, ignore_errors=True)


def _convert(workdir, main, out, font, tmp, keep):
    pdf = _ensure_pdf(workdir, main)
    _need('pdftotext')

    # ---------- 1. make4ht: konten via engine TeX asli ----------
    _run([_need('make4ht'), '-f', 'odt', f'{main}.tex'], workdir)
    odt = os.path.join(workdir, f'{main}.odt')
    if not os.path.isfile(odt):
        raise SystemExit(f'[pipeline] make4ht tidak menghasilkan {main}.odt')

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
    toc = os.path.join(workdir, f'{main}.toc')
    graft.run(pdf, p2, out, toc if os.path.isfile(toc) else '')
    return out


def main(argv=None):
    a = sys.argv[1:] if argv is None else argv
    import argparse
    ap = argparse.ArgumentParser(
        prog='tex2docx',
        description='LaTeX -> DOCX via engine TeX asli + oracle layout PDF')
    ap.add_argument('workdir', help='folder berisi <main>.tex')
    ap.add_argument('--main', default='main', help='nama dasar file .tex (tanpa ekstensi)')
    ap.add_argument('--out', default=None, help='path DOCX output')
    ap.add_argument('--font', default='Times New Roman', help='font keluaran')
    ap.add_argument('--keep', action='store_true', help='pertahankan file antara')
    args = ap.parse_args(a)
    path = convert(args.workdir, args.main, args.out, args.font, args.keep)
    print(f'[pipeline] selesai: {path}')


if __name__ == '__main__':
    main()
