"""Unit test mode --interaction (nonstopmode dsb.) — stdlib murni, tanpa
engine eksternal. Menguji: runner TeX toleran vs ketat (makna asli
pdflatex -interaction=nonstopmode: exit != 0 bukan vonis), diagnosis dari
ekor .log, injeksi mode ke latex internal make4ht via config XDG, dan
penerimaan nilai CLI."""
import sys

import pytest

import tex2docx_engine as pkg  # via tests/conftest.py


# ---------- runner TeX: toleran (nonstopmode) vs ketat (errorstopmode) ----------

def test_run_tex_toleran_lanjut_pada_exit_nonzero(tmp_path):
    r = pkg.pipeline._run_tex([sys.executable, '-c', 'import sys; sys.exit(3)'],
                              str(tmp_path), strict=False)
    assert r.returncode == 3          # lanjut, exit code dilaporkan apa adanya


def test_run_tex_ketat_berhenti_pada_exit_nonzero(tmp_path):
    with pytest.raises(SystemExit):
        pkg.pipeline._run_tex([sys.executable, '-c', 'import sys; sys.exit(3)'],
                              str(tmp_path), strict=True)


def test_run_tex_sukses_senyap(tmp_path):
    r = pkg.pipeline._run_tex([sys.executable, '-c', 'pass'],
                              str(tmp_path), strict=False)
    assert r.returncode == 0


def test_strict_hanya_errorstopmode():
    assert pkg.pipeline._strict('errorstopmode') is True
    for mode in ('nonstopmode', 'batchmode', 'scrollmode'):
        assert pkg.pipeline._strict(mode) is False


# ---------- diagnosis kegagalan final dari ekor <main>.log ----------

def test_log_tail_kosong_bila_log_tak_ada(tmp_path):
    assert pkg.pipeline._latex_log_tail(str(tmp_path), 'main') == ''


def test_log_tail_baca_ekor_log(tmp_path):
    (tmp_path / 'main.log').write_text(
        '\n'.join(f'baris{i}' for i in range(30)), encoding='utf-8')
    tail = pkg.pipeline._latex_log_tail(str(tmp_path), 'main')
    assert tail.endswith('baris29')
    assert 'baris0' not in tail


# ---------- penemuan config make4ht milik user (urutan = make4ht sendiri) ----------

def test_user_config_make4ht_di_proyek(tmp_path):
    (tmp_path / '.make4ht').write_text("settings_add { custom = 'x' }",
                                       encoding='utf-8')
    path, content = pkg.pipeline._find_make4ht_user_config(
        str(tmp_path), {}, home=str(tmp_path / 'home'), xdg=str(tmp_path / 'xdg'))
    assert path.endswith('.make4ht')
    assert "custom = 'x'" in content


def test_user_config_xdg_dan_home_fallback(tmp_path):
    home = tmp_path / 'home'
    xdg = tmp_path / 'xdgconfig'
    env = {}
    # tidak ada config sama sekali
    assert pkg.pipeline._find_make4ht_user_config(
        str(tmp_path), env, home=str(home), xdg=str(xdg)) == ('', '')
    # ~/.make4ht (fallback HOME) ditemukan bila XDG kosong
    home.mkdir(parents=True, exist_ok=True)
    (home / '.make4ht').write_text('home-config', encoding='utf-8')
    path, content = pkg.pipeline._find_make4ht_user_config(
        str(tmp_path), env, home=str(home), xdg=str(xdg))
    assert path.endswith('.make4ht') and content == 'home-config'
    # XDG CONFIG (make4ht/config.lua) lebih diutamakan daripada ~/.make4ht
    cfgdir = xdg / 'make4ht'
    cfgdir.mkdir(parents=True)
    (cfgdir / 'config.lua').write_text('xdg-config', encoding='utf-8')
    path, content = pkg.pipeline._find_make4ht_user_config(
        str(tmp_path), env, home=str(home), xdg=str(xdg))
    assert path.endswith('config.lua') and content == 'xdg-config'


# ---------- injeksi mode ke latex internal make4ht ----------

def test_env_batchmode_tanpa_injeksi(tmp_path):
    """batchmode = default make4ht sendiri — tidak ada yang perlu diinjeksi,
    env dan tmp tak tersentuh (perilaku identik dengan sebelum ada opsi ini)."""
    env = pkg.pipeline._make4ht_env(str(tmp_path), str(tmp_path), 'batchmode',
                                    env={'A': 'B'})
    assert env == {'A': 'B'}
    assert not (tmp_path / 'xdg').exists()


def test_env_nonstopmode_injeksi_config_xdg(tmp_path):
    env = pkg.pipeline._make4ht_env(str(tmp_path), str(tmp_path), 'nonstopmode',
                                    env={'A': 'B'})
    cfg = tmp_path / 'xdg' / 'make4ht' / 'config.lua'
    assert env['XDG_CONFIG_HOME'] == str(tmp_path / 'xdg')
    assert 'set_settings { interaction = "nonstopmode" }' in cfg.read_text('utf-8')


def test_env_config_user_dipertahankan_dan_ditempel_dibawahnya(tmp_path):
    """Config XDG milik user tidak boleh hilang; penugasan engine di depan
    supaya preferensi eksplisit user (set_settings belakangan) tetap menang."""
    home = tmp_path / 'home'
    xdg = tmp_path / 'xdguser'
    cfgdir = xdg / 'make4ht'
    cfgdir.mkdir(parents=True)
    (cfgdir / 'config.lua').write_text("settings_add { custom = 'punya-user' }",
                                      encoding='utf-8')
    env = pkg.pipeline._make4ht_env(str(tmp_path), str(tmp_path), 'scrollmode',
                                    env={}, home=str(home), xdg=str(xdg))
    body = (tmp_path / 'xdg' / 'make4ht' / 'config.lua').read_text('utf-8')
    assert 'set_settings { interaction = "scrollmode" }' in body
    assert "custom = 'punya-user'" in body
    assert body.index('set_settings') < body.index('punya-user')


def test_env_make4ht_proyek_menghentikan_injeksi(tmp_path, capsys):
    """`.make4ht` di proyek menang atas config XDG di make4ht — injeksi tak
    mungkin; env tak boleh dialihkan dan user diberi tahu."""
    (tmp_path / '.make4ht').write_text('-- punya user', encoding='utf-8')
    env = pkg.pipeline._make4ht_env(str(tmp_path), str(tmp_path), 'nonstopmode',
                                    env={})
    assert 'XDG_CONFIG_HOME' not in env
    assert '.make4ht' in capsys.readouterr().out


# ---------- penerimaan nilai mode ----------

def test_daftar_mode_valid():
    assert set(pkg.pipeline.INTERACTION_MODES) == {
        'errorstopmode', 'batchmode', 'nonstopmode', 'scrollmode'}


def test_convert_tolak_mode_tidak_dikenal(tmp_path):
    with pytest.raises(SystemExit):
        pkg.convert(str(tmp_path), main='apa', interaction='nomode')
