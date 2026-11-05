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

# Mode interaksi TeX (`--interaction`, setara flag asli pdflatex). Default
# `nonstopmode`: error yang bisa dipulihkan TIDAK mematikan konversi — pdflatex
# sendiri berperilaku begitu (tetap menulis PDF, exit code tetap 1), jadi
# penentu lolos-tidaknya adalah artefak akhir, bukan exit code. `errorstopmode`
# = perilaku ketat: satu error pun menghentikan segalanya.
INTERACTION_MODES = ('errorstopmode', 'batchmode', 'nonstopmode', 'scrollmode')


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


def _strict(interaction):
    """errorstopmode = berhenti pada error TeX pertama (perilaku lama)."""
    return interaction == 'errorstopmode'


def _run_tex(cmd, cwd, strict):
    """Runner perintah TeX (pdflatex, bibtex). Di luar errorstopmode, exit
    != 0 hanya PERINGATAN — makna asli nonstopmode: engine memulihkan error
    yang bisa dipulihkan dan tetap menulis artefaknya. Artefak akhir (bukan
    exit code) yang menentukan; kegagalan final didiagnosis dari .log."""
    r = subprocess.run(cmd, cwd=cwd, capture_output=True, text=True)
    if r.returncode != 0:
        tail = '\n'.join((r.stderr or r.stdout).splitlines()[-12:])
        if strict:
            raise SystemExit(f'[pipeline] gagal: {" ".join(cmd)}\n{tail}')
        print(f'[pipeline] exit={r.returncode} (error TeX recoverable — '
              f'mode nonstopmode, lanjut): {" ".join(cmd)}\n{tail}')
    return r


def _latex_log_tail(workdir, main, n=12):
    """Ekor <main>.log — sumber diagnosis terbaik bila PDF tak jadi, karena
    di luar errorstopmode exit code pdflatex tak lagi menjadi penentu."""
    log = os.path.join(workdir, f'{main}.log')
    try:
        with open(log, encoding='utf-8', errors='replace') as f:
            return '\n'.join(f.read().splitlines()[-n:])
    except OSError:
        return ''


def _find_make4ht_user_config(workdir, env=None, home=None, xdg=None):
    """Cari config make4ht milik user, dalam urutan penemuan make4ht sendiri
    (mk_config.find_config lalu find_xdg_config): `.make4ht` di workdir atau
    induk-induknya, lalu $XDG_CONFIG_HOME/make4ht/config.lua, lalu ~/.make4ht.
    Mengembalikan (path, isi); ('', '') bila tidak ada."""
    env = env if env is not None else os.environ
    cur = os.path.abspath(workdir)
    while True:
        cand = os.path.join(cur, '.make4ht')
        if os.path.isfile(cand):
            try:
                with open(cand, encoding='utf-8', errors='replace') as f:
                    return cand, f.read()
            except OSError:
                break
        parent = os.path.dirname(cur)
        if parent == cur:
            break
        cur = parent
    home = home or env.get('HOME') or os.path.expanduser('~')
    xdg = xdg or env.get('XDG_CONFIG_HOME') or os.path.join(home, '.config')
    for cand in (os.path.join(xdg, 'make4ht', 'config.lua'),
                 os.path.join(home, '.make4ht')):
        if os.path.isfile(cand):
            try:
                with open(cand, encoding='utf-8', errors='replace') as f:
                    return cand, f.read()
            except OSError:
                pass
    return '', ''


def _make4ht_env(workdir, tmp, interaction, env=None, home=None, xdg=None):
    """Kirim mode `--interaction` ke latex INTERNAL make4ht. make4ht tidak
    punya flag CLI untuk ini — default-nya batchmode senyap (mkparams.lua),
    dan cara resminya adalah config make4ht: `set_settings{interaction=...}`
    dimuat ke tabel parameter SEBELUM build file user. Injeksi = XDG_CONFIG_HOME
    subprocess menunjuk config yang dibuat di `tmp`:

        set_settings { interaction = "<mode>" }   <- punya engine, di depan
        <isi config make4ht user, bila ada>       <- belakangan, user menang

    Bila proyek (atau induknya) punya `.make4ht`, make4ht memprioritaskan file
    itu dan TAK AKAN membaca XDG — mode CLI tak bisa masuk lewat sini; env
    dikembalikan apa adanya + user diberi tahu. `batchmode` = default make4ht:
    tanpa injeksi, perilaku identik dengan sebelum ada opsi ini."""
    env = dict(env if env is not None else os.environ)
    if interaction == 'batchmode':
        return env
    user_path, user_cfg = _find_make4ht_user_config(workdir, env, home, xdg)
    if user_path and os.path.basename(user_path) == '.make4ht':
        print(f'[pipeline] {user_path} ada — make4ht memprioritaskannya di atas '
              f'config XDG; mode --interaction tak diinjeksi ke latex internal '
              f'make4ht (pdflatex tetap memakai {interaction})')
        return env
    cfgdir = os.path.join(tmp, 'xdg', 'make4ht')
    os.makedirs(cfgdir, exist_ok=True)
    cfg = os.path.join(cfgdir, 'config.lua')
    parts = ['-- tex2docx-engine: mode --interaction untuk latex internal make4ht\n'
             '-- (default make4ht = batchmode). Config make4ht milik user (bila\n'
             '-- ada) ditempel DI BAWAH penugasan ini — preferensi eksplisit\n'
             '-- user tetap menang.\n'
             f'set_settings {{ interaction = "{interaction}" }}\n']
    if user_cfg:
        parts.append('\n-- ---- config make4ht milik user (dipertahankan) ----\n'
                     + user_cfg)
    with open(cfg, 'w', encoding='utf-8') as f:
        f.write(''.join(parts))
    env['XDG_CONFIG_HOME'] = os.path.join(tmp, 'xdg')
    return env


def _ensure_pdf(workdir, main, interaction='nonstopmode'):
    """PDF = oracle; butuh >= 2 run pdflatex agar .toc/.lot/.lof terisi.
    Bila ada .bib (dan .bbl belum ada), jalankan bibtex di antaranya.
    Dijalankan dengan `-interaction=<interaction>`: di luar errorstopmode,
    error recoverable hanya peringatan — keberadaan PDF yang menentukan."""
    pdf = os.path.join(workdir, f'{main}.pdf')
    tex = os.path.join(workdir, f'{main}.tex')
    if os.path.isfile(pdf):
        return pdf
    if not os.path.isfile(tex):
        raise SystemExit(f'[pipeline] {tex} tidak ada')
    strict = _strict(interaction)
    pdflatex = _need('pdflatex')
    _run_tex([pdflatex, f'-interaction={interaction}', f'{main}.tex'], workdir, strict)
    if not os.path.isfile(os.path.join(workdir, f'{main}.bbl')):
        for bib in ('{m}.bib'.format(m=main), 'references.bib'):
            if os.path.isfile(os.path.join(workdir, bib)):
                _run_tex([_need('bibtex'), main], workdir, strict)
                break
    for _ in range(2):
        _run_tex([pdflatex, f'-interaction={interaction}', f'{main}.tex'], workdir, strict)
    if not os.path.isfile(pdf):
        raise SystemExit(f'[pipeline] pdflatex tidak menghasilkan {main}.pdf '
                         f'(mode {interaction}) — ekor {main}.log:\n'
                         f'{_latex_log_tail(workdir, main)}')
    return pdf


def convert(workdir, main='main', out=None, font='Times New Roman',
            keep=False, quiet=False, interaction='nonstopmode'):
    """Konversi penuh. `workdir` = folder berisi <main>.tex (dan PDF bila sudah
    dikompilasi). `interaction` = mode interaksi TeX (default nonstopmode —
    error recoverable tidak menghentikan konversi). Mengembalikan path DOCX
    final."""
    if interaction not in INTERACTION_MODES:
        raise SystemExit(f'[pipeline] --interaction {interaction} tidak dikenal '
                         f'— pilih dari: {", ".join(INTERACTION_MODES)}')
    workdir = os.path.abspath(workdir)
    out = os.path.abspath(out) if out else os.path.join(workdir, f'{main}_engine.docx')
    flat = main.replace('/', '_').replace('\\', '_')      # main boleh 'bab/main'
    tmp = os.path.join(workdir, f'.{flat}_tex2docx_tmp')
    os.makedirs(tmp, exist_ok=True)
    try:
        return _convert(workdir, main, out, font, tmp, keep, interaction)
    finally:
        if not keep:
            shutil.rmtree(tmp, ignore_errors=True)


def _convert(workdir, main, out, font, tmp, keep, interaction='nonstopmode'):
    pdf = _ensure_pdf(workdir, main, interaction)
    _need('pdftotext')

    # ---------- 1. make4ht: konten via engine TeX asli ----------
    # make4ht kerap exit != 0 walau ODT tetap dihasilkan lengkap (error htlatex
    # yang recoverable — mis. \insertchapterspace undefined di cls custom).
    # Kriteria sukses tahap ini = <main>.odt terbentuk. Mode --interaction
    # diinjeksi ke latex internalnya lewat config make4ht yang dibuat di tmp
    # (default make4ht = batchmode senyap — lihat _make4ht_env).
    odt = os.path.join(workdir, f'{main}.odt')
    before = os.path.getmtime(odt) if os.path.isfile(odt) else 0
    r = subprocess.run([_need('make4ht'), '-f', 'odt', f'{main}.tex'],
                       cwd=workdir, env=_make4ht_env(workdir, tmp, interaction),
                       capture_output=True, text=True)
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
    ap.add_argument('--interaction', default='nonstopmode',
                    choices=INTERACTION_MODES,
                    help='mode interaksi TeX, seperti flag asli pdflatex '
                         '(default nonstopmode: error recoverable tidak '
                         'menghentikan konversi; errorstopmode = berhenti '
                         'pada error pertama)')
    args = ap.parse_args(a)
    path = convert(args.workdir, args.main, args.out, args.font, args.keep,
                   interaction=args.interaction)
    print(f'[pipeline] selesai: {path}')


if __name__ == '__main__':
    main()
