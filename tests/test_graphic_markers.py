"""Unit test penambalan abort tex4ht pada grafik — stdlib murni, tanpa engine.
tex4ht fopen file instruksi .4og untuk tiap gambar di DVI dan EXIT KERAS bila
gagal (bad_in_file di tex4ht.c): path di luar workdir, atau FOLDER gambar tak
ada. Marker 0-byte (+ foldernya) harus dibuat proaktif/reaktif — tapi TIDAK
untuk gambar dalam workdir yang foldernya ada (tex4ht urus sendiri)."""
import tex2docx_engine as pkg  # via tests/conftest.py


def _tex(tmp_path, body):
    (tmp_path / 'main.tex').write_text(body, encoding='utf-8')


def test_marker_path_luar_workdir_dibuat(tmp_path):
    _tex(tmp_path, r'\includegraphics{../aset/gambar.png}')
    touched = pkg.pipeline._prepare_graphic_markers(str(tmp_path))
    assert touched, 'marker luar workdir wajib dibuat'
    assert (tmp_path / '../aset/gambar.png.4og').resolve().is_file()


def test_marker_folder_gambar_tak_ada_dibuat(tmp_path):
    r"""Kasus 'gambar nggak ada': pdflatex nonstopmode lolos, tex4ht abort saat
    fopen marker .4og di folder yang tidak ada. makedirs + marker wajib."""
    _tex(tmp_path, r'\includegraphics{figures/gambar-hilang.png}')
    touched = pkg.pipeline._prepare_graphic_markers(str(tmp_path))
    assert touched
    assert (tmp_path / 'figures').is_dir()
    assert (tmp_path / 'figures/gambar-hilang.png.4og').is_file()


def test_gambar_dalam_workdir_folder_ada_tak_diolok(tmp_path):
    r"""Folder ada, gambar tidak ada -> tex4ht sukses menulis markernya dan
    t4ht hanya warning; engine TIDAK boleh mengolok-olok folder user."""
    figures = tmp_path / 'figures'
    figures.mkdir()
    _tex(tmp_path, r'\includegraphics{figures/nam salah.png}')
    assert pkg.pipeline._prepare_graphic_markers(str(tmp_path)) == []
    assert not (figures / 'nam salah.png.4og').exists()


def test_gambar_ada_tak_diolok(tmp_path):
    figures = tmp_path / 'figures'
    figures.mkdir()
    (figures / 'ok.png').write_bytes(b'png')
    _tex(tmp_path, r'\includegraphics[width=0.5\textwidth]{figures/ok.png}')
    assert pkg.pipeline._prepare_graphic_markers(str(tmp_path)) == []
    assert not (figures / 'ok.png.4og').exists()


def test_reaktif_dari_mkout_buang_akhiran_4og(tmp_path):
    r"""Error tex4ht mencantumkan NAMA MARKER (`...png.4og') — akhiran .4og
    harus dikupas, jangan membuat marker bertumpuk `.4og.4og`."""
    touched = pkg.pipeline._prepare_graphic_markers(
        str(tmp_path), "--- error --- Can't find/open file `../luar/g.png.4og'\n")
    assert touched
    assert (tmp_path / '../luar/g.png.4og').resolve().is_file()
    assert not (tmp_path / '../luar/g.png.4og.4og').exists()


def test_workdir_tanpa_grafik_tanpa_touch(tmp_path):
    _tex(tmp_path, 'Teks biasa tanpa gambar.')
    assert pkg.pipeline._prepare_graphic_markers(str(tmp_path)) == []


def test_idempoten_tanpa_duplikat(tmp_path):
    _tex(tmp_path, r'\includegraphics{figures/a.png}')
    pertama = pkg.pipeline._prepare_graphic_markers(str(tmp_path))
    kedua = pkg.pipeline._prepare_graphic_markers(str(tmp_path))
    assert pertama and kedua == []          # sudah ada -> tidak disentuh lagi
