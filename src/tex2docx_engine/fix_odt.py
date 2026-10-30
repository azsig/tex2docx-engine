"""
fix_odt — dua lapis perbaikan content.xml ODT jalur make4ht (tex2docx-engine):

1. MESIN MARKER NATIVE — komentar marker yang dipancarkan t2dcfg.cfg (tex4ht)
   diterjemahkan jadi automatic style ODT SEBELUM LibreOffice, sehingga
   semantik spasi/perataan dibaca native dari LaTeX, bukan diduga ulang dari
   oracle PDF:
     <!--t2dvs:Npt-->    -> fo:margin-top blok konten berikutnya (glue
                           beruntun dijumlahkan; hasil bersih <1pt dibuang)
     <!--t2dalign:X-->   -> fo:text-align blok di dalam scope, sampai
     <!--t2dalign:end-->    penutup scope (\aftergroup group/env)
     <!--t2dbrk-->       -> fo:break-before="page" blok berikutnya
   Marker intra-paragraf (mis. \\[X] dan \\\\ di tengah paragraf) dibuang —
   jarak antar-baris dalam satu paragraf tak terrepresentasikan di Word (v1).
   \\\\ polos sendiri sudah menjadi <text:line-break/> via \\Configure{newline}.

   Judul bab: t2dcfg menanam fo:text-align di style Heading-1..4 (styles.xml).
   Kehadiran atribut itu = sinyal "ditanam engine" (default ooffice.4ht tidak
   memilikinya) -> align disalin ke automatic style tiap text:h terkait
   (anti-drop saat export DOCX oleh LibreOffice).

2. BUNGKUS INLINE YATIM (perilaku lama tetap) — elemen inline langsung di
   bawah <office:text> dibungkus <text:p> agar LibreOffice tak membuangnya.

Dokumen tanpa marker: satu-satunya perubahan adalah komentar dibuang (sama
dengan parser lama yang juga membuang semua komentar) — ODT lama tetap
kompatibel.
"""
import io
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

TEXT_NS = 'urn:oasis:names:tc:opendocument:xmlns:text:1.0'
OFFICE_NS = 'urn:oasis:names:tc:opendocument:xmlns:office:1.0'
STYLE_NS = 'urn:oasis:names:tc:opendocument:xmlns:style:1.0'
# fo: di output make4ht memakai URI XSL-Format klasik (ODF 1.0), bukan
# varian xsl-fo-compatible — harus persis sama agar tak lahir prefiks ns3:.
FO_NS = 'http://www.w3.org/1999/XSL/Format'
TABLE_NS = 'urn:oasis:names:tc:opendocument:xmlns:table:1.0'

BLOCK_TAGS = {
    f'{{{TEXT_NS}}}p', f'{{{TEXT_NS}}}h', f'{{{TEXT_NS}}}list',
    f'{{{TEXT_NS}}}section', f'{{{TEXT_NS}}}table-of-content',
    f'{{{TEXT_NS}}}bibliography',
    f'{{{TABLE_NS}}}table',
}
# Blok konten yang sah menerima fo:margin-top / fo:text-align / fo:break-before.
CONTENT_TAGS = {f'{{{TEXT_NS}}}p', f'{{{TEXT_NS}}}h', f'{{{TEXT_NS}}}list'}
# Wadah float/TOC/pustaka: TIDAK mengonsumsi pending — marker menunggu blok
# konten nyata pertama di dalam / setelahnya.
CONTAINER_TAGS = {f'{{{TEXT_NS}}}section', f'{{{TEXT_NS}}}table-of-content',
                  f'{{{TEXT_NS}}}bibliography'}

RE_VS = re.compile(r't2dvs:(-?[\d.]+)pt')
RE_ALIGN = re.compile(r't2dalign:(center|left|right|end)')
RE_BRK = re.compile(r't2dbrk')
RE_T2D_ANY = re.compile(r't2d(vs|align|brk)')
RE_T2D_BYTES = re.compile(rb'<!--\s*t2d[^>]*-->')
# </text:p> LIAR tepat setelah penutup kontainer daftar (paragraf ODF tak
# pernah MEMBAWA index-body/bibliography/list — tutup-p setelahnya pasti
# sampah struktural dari pemulihan error "Missing }" \end{thebibliography}).
RE_STRAY_CLOSE_P = re.compile(
    r'(</text:(?:index-body|bibliography|list)>)\s*</text:p>')
ROOT_CLOSE = '</office:document-content>'


def _repair_content(data):
    """Perbaikan string PRA-parse untuk luka struktural ujung dokumen tex4ht
    (keduanya terverifikasi di skripsi uji, keduanya di ujung bibliografi):

    1. `</text:p>` liar setelah penutup kontainer daftar — error kronis
       "Missing } inserted" saat \\end{thebibliography} (.bbl ber-\\url)
       menyisakan satu tag tutup paragraf tanpa pembuka, DI LUAR kontainer.
    2. Apa pun SETELAH </office:document-content> — marker <!--t2dbrk--> dari
       \\clearpage yang terpanggil di \\end{document} terpancar SETELAH root
       tertutup (t4ht sudah menutup semuanya saat itu).

    make4ht biasanya menyembuhkan keduanya via reparse HTML-DOM domfilter;
    repair ini menjamin fix_odt tak bergantung pada penyembuh itu (dan tetap
    benar untuk .4oo mentah jalur tanpa-make4ht)."""
    i = data.find(ROOT_CLOSE)
    if i != -1:
        data = data[:i + len(ROOT_CLOSE)]
    return RE_STRAY_CLOSE_P.sub(r'\1', data)


class StyleRegistry:
    """Automatic style t2dP{n} (family=paragraph) dengan dedup + merge:
    blok yang sama dapat props dari beberapa pass (margin + align heading)
    tanpa merantai automatic style atas automatic style."""

    def __init__(self, root):
        self.root = root
        self.by_key = {}   # (parent, props-tuple) -> name
        self.by_name = {}  # name -> paragraph-properties element

    def apply(self, el, props):
        if not props:
            return
        cur = el.get(f'{{{TEXT_NS}}}style-name')
        if cur in self.by_name:  # el sudah pakai t2d style -> gabung atribut
            pp = self.by_name[cur]
            for k, v in props.items():
                pp.set(f'{{{FO_NS}}}{k}', v)
            return
        parent_name = cur or 'Standard'
        key = (parent_name, tuple(sorted(props.items())))
        name = self.by_key.get(key)
        if name is None:
            autostyles = self.root.find(f'{{{OFFICE_NS}}}automatic-styles')
            if autostyles is None:
                body = self.root.find(f'{{{OFFICE_NS}}}body')
                autostyles = ET.Element(f'{{{OFFICE_NS}}}automatic-styles')
                self.root.insert(list(self.root).index(body), autostyles)
            name = f't2dP{len(self.by_name) + 1}'
            st = ET.SubElement(autostyles, f'{{{STYLE_NS}}}style', {
                f'{{{STYLE_NS}}}name': name,
                f'{{{STYLE_NS}}}family': 'paragraph',
                f'{{{STYLE_NS}}}parent-style-name': parent_name,
            })
            pp = ET.SubElement(st, f'{{{STYLE_NS}}}paragraph-properties')
            for k, v in props.items():
                pp.set(f'{{{FO_NS}}}{k}', v)
            self.by_key[key] = name
            self.by_name[name] = pp
        el.set(f'{{{TEXT_NS}}}style-name', name)


def parse_with_comments(xml_text):
    """ElementTree default MENGHAPUS semua komentar — marker butuh mereka."""
    parser = ET.XMLParser(target=ET.TreeBuilder(insert_comments=True))
    return ET.fromstring(xml_text, parser=parser)


def planted_heading_aligns(styles_xml):
    """Scan styles.xml: Heading-1..4 yang fo:text-align-nya ditanam t2dcfg.
    Return {nama style: nilai align}.

    Agnostik-prefix: make4ht memakai URI fo BERBEDA per bagian (styles.xml =
    xsl-fo-compatible, content.xml = XSL/Format klasik) — cari atribut
    text-align apa pun namanya."""
    out = {}
    if not styles_xml:
        return out
    try:
        root = ET.fromstring(styles_xml)
    except ET.ParseError:
        return out
    for st in root.iter(f'{{{STYLE_NS}}}style'):
        name = st.get(f'{{{STYLE_NS}}}name')
        if not (name and re.fullmatch(r'Heading-[1-4]', name)):
            continue
        for pp in st.findall(f'{{{STYLE_NS}}}paragraph-properties'):
            align = next((v for k, v in pp.attrib.items()
                          if k.rsplit('}', 1)[-1] == 'text-align'), None)
            if align:
                out[name] = align
    return out


def run(SRC, DST):
    zin = zipfile.ZipFile(SRC)
    content = _repair_content(zin.read('content.xml').decode('utf-8'))

    # Prefiks wajib dipristinakan SEBELUM serialisasi: content.xml buatan
    # make4ht tidak selalu mendeklarasikan fo: — tanpa ini ET menulis
    # "ns3:break-before" (namespace sah tapi membingungkan LibreOffice/Word).
    for prefix, uri in (
        ('office', OFFICE_NS), ('style', STYLE_NS), ('text', TEXT_NS),
        ('fo', FO_NS), ('table', TABLE_NS),
        ('draw', 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0'),
    ):
        ET.register_namespace(prefix, uri)
    for _, (prefix, uri) in ET.iterparse(io.StringIO(content), events=['start-ns']):
        ET.register_namespace(prefix, uri)

    root = parse_with_comments(content)
    office_text = root.find(f'.//{{{OFFICE_NS}}}text')
    if office_text is None:
        print('[fix_odt] <office:text> tidak ketemu — tidak ada yang diubah')
        return

    stats = dict(margin=0, align=0, brk=0, drop_inline=0, drop_vs=0,
                 hardened=0, comments=0)
    reg = StyleRegistry(root)

    # ---- pass A: buang SEMUA komentar; tandai \newpage di ujung paragraf ----
    # (komentar debug tex4ht 'l. N' ikut dibuang — parser lama juga membuang
    # semua komentar, jadi ini setara perilaku sebelumnya.)
    parent = {c: p for p in root.iter() for c in p}

    # Klasifikasi marker DI DALAM paragraf berdasar posisi (marker HCode polos
    # dari t2dcfg jatuh pada titik sisip aktual TeX):
    #   LEADING  (tanpa konten sebelum)  -> berlaku UNTUK BLOK INI
    #             (vs -> margin-top blok ini; align -> align blok ini + buka
    #              scope; brk -> break-before blok ini)
    #   TRAILING (tanpa konten sesudah)  -> berlaku UNTUK BLOK BERIKUTNYA
    #             (vs/brk -> pending; align:X -> buka scope; align:end ->
    #              tutup scope SETELAH blok ini)
    #   TENGAH   (konten dua sisi)       -> intra-baris (\\[X] dsb.) — buang
    # Sasaran marker = PARAGRAF LANGSUNG pembawanya (par). Jangan panjat ke
    # ancestor top-level: paragraf di dalam text:section (env center/quote)
    # atau list-item punya ancestor section/list — memakai ancestor sebagai
    # "blok" membuat SEMUA marker di dalamnya tersesat sebagai drop-inline
    # (skripsi uji: 794 marker align/vs lenyap padahal posisinya sah).
    leading = {}     # id(text:p/h) -> list marker utk blok itu
    open_scope = {}  # id(text:p/h) -> nilai align (buka di trailing blok ini)
    close_after = set()   # id(text:p/h) — align:end trailing: blok ini terakhir scope
    trailing_vs = {}      # id(blok) -> [pt] pending utk blok berikutnya
    trailing_brk = set()  # id() blok yang diakhiri <!--t2dbrk-->
    # text:p yang isinya HANYA marker: HCode di vmode memang melahirkan
    # paragraf kosong pembawa komentar (skripsi uji: +300 paragraf hampa).
    # Paragraf ini transparan — markernya masuk state machine pass B,
    # elemennya dibuang total (bukan sekadar dikosongkan).
    marker_only = set()
    for node in list(root.iter()):
        if node.tag is not ET.Comment:
            continue
        stats['comments'] += 1
        txt = node.text or ''
        par = parent.get(node)
        if par is None:
            continue
        if par is office_text:
            continue  # marker ANTAR blok — urusan state machine pass B
        if par.tag in CONTAINER_TAGS or par.tag == f'{{{TEXT_NS}}}list-item':
            # marker antar-blok DI DALAM kontainer (section/TOC/pustaka/
            # list-item) — DIBIARKAN agar walk() rekursif membacanya.
            # (Dulu masuk klasifikasi inline dan DIHAPUS di sini: scope
            # align yang dibuka di dalam env tak pernah tertutup di level
            # state — skripsi uji: \centering kepala bab melumuri seluruh
            # isi bab, semua Heading-2 jadi center padahal kiri di PDF.)
            continue
        if not RE_T2D_ANY.search(txt):
            par.remove(node)
            continue
        if par.tag in CONTENT_TAGS and _is_marker_only(par):
            marker_only.add(id(par))
        is_head = _is_leading(par, node)
        is_tail = _is_trailing(par, node)
        block = par
        blk_ok = par.tag in CONTENT_TAGS
        m = RE_VS.search(txt)
        if m:
            val = float(m.group(1))
            if is_head and blk_ok:
                leading.setdefault(id(block), []).append(('vs', val))
            elif is_tail and blk_ok:
                trailing_vs.setdefault(id(block), []).append(val)
            else:
                stats['drop_inline'] += 1
        elif RE_ALIGN.search(txt):
            kind = RE_ALIGN.search(txt).group(1)
            if kind == 'end':
                if is_head and blk_ok:
                    close_after.discard(id(block))  # scope tutup sebelum blok ini
                    leading.setdefault(id(block), []).append(('alignend', None))
                elif is_tail and blk_ok:
                    close_after.add(id(block))
                else:
                    stats['drop_inline'] += 1
            else:
                if is_head and blk_ok:
                    leading.setdefault(id(block), []).append(('align', kind))
                elif is_tail and blk_ok:
                    open_scope[id(block)] = kind
                else:
                    stats['drop_inline'] += 1
        elif RE_BRK.search(txt):
            if is_head and blk_ok:
                leading.setdefault(id(block), []).append(('brk', None))
            elif is_tail and blk_ok:
                if blk_ok:
                    trailing_brk.add(id(block))
            else:
                stats['drop_inline'] += 1
        par.remove(node)

    # ---- pass B: state machine atas isi <office:text>, REKURSIF ----
    # Dua sumber marker BERGABUNG di sini: (a) komentar antar blok (anak
    # langsung wadah — state di bawah), (b) marker leading/trailing di dalam
    # paragraf (pass A). Glue LaTeX menjumlah, jadi vs dari kedua sumber
    # dijumlahkan sebelum ditanam sebagai margin-top.
    # Walker MENURUN ke text:section (env center/quote/float), text:list
    # (itemnya) dan table-of-content/bibliography: mayoritas \centering dan
    # \vspace dokumen nyata hidup DI DALAM environment tersebut — tanpa
    # rekursi markernya mati semua (skripsi uji: align=0, margin=7).
    # Marker sebelum blok konten PERTAMA dibuang: halaman 1 sudah mulai.
    state = {'vs': [], 'align': None, 'brk': False, 'seen': False}
    CELL_TAG = f'{{{TABLE_NS}}}table-cell'
    LIST_ITEM_TAG = f'{{{TEXT_NS}}}list-item'

    def target_of(el):
        """text:list tak menerima style paragraf — sasaran paragraf pertama."""
        if el.tag == f'{{{TEXT_NS}}}list':
            for p in el.iter(f'{{{TEXT_NS}}}p'):
                return p
            return None
        return el

    def consume(el):
        """Terapkan semua sumber marker pada blok konten `el`.

        Scope align = SATU state (`state['align']`) untuk SEMUA sumber
        marker: antar-blok (walk), leading/trailing paragraf (pass A),
        marker-only (absorb). Dulu tail align:end hanya menull-kan
        `intra_scope` padahal scope dibuka lewat `state['align']` — scope
        tak pernah tertutup dan \centering bocor lintas bab (semua
        Heading-2 skripsi uji jadi center padahal kiri di PDF)."""
        tgt = target_of(el)
        props = {}
        vs_vals = list(state['vs']) + [v for k, v in leading.get(id(el), [])
                                       if k == 'vs']
        blk_align = None
        for k, v in leading.get(id(el), []):
            if k == 'align':
                blk_align = v
            elif k == 'brk':
                if state['seen']:
                    props['break-before'] = 'page'
                    stats['brk'] += 1
            elif k == 'alignend':
                state['align'] = None   # penutup sebelum blok ini
        if blk_align:
            state['align'] = blk_align  # leading align juga membuka scope
        if state['seen']:
            if vs_vals:
                net = sum(vs_vals)
                if net >= 1.0:
                    props['margin-top'] = f'{round(net, 3)}pt'
                    stats['margin'] += 1
                else:
                    stats['drop_vs'] += 1
            if state['brk']:
                props.setdefault('break-before', 'page')
                stats['brk'] += 1
        state['seen'] = True
        if state['align']:
            props['text-align'] = state['align']
            stats['align'] += 1
        if tgt is not None:
            reg.apply(tgt, props)
        state['vs'] = []
        state['brk'] = False
        # transisi scope dari marker TRAILING blok ini (berlaku blok berikut)
        if id(el) in open_scope:
            state['align'] = open_scope[id(el)]
        if id(el) in close_after:
            state['align'] = None

    def absorb_marker_only(el, container):
        """Paragraf hampa pembawa marker: markernya masuk state antar-blok,
        elemennya dibuang (HCode vmode t4ht melahirkannya otomatis)."""
        for k, v in leading.pop(id(el), []):
            if k == 'vs':
                state['vs'].append(v)
            elif k == 'align':
                state['align'] = v
            elif k == 'alignend':
                state['align'] = None
            elif k == 'brk':
                state['brk'] = True
        trailing_vs.pop(id(el), None)
        trailing_brk.discard(id(el))
        open_scope.pop(id(el), None)
        close_after.discard(id(el))
        container.remove(el)

    def walk(container):
        for el in list(container):
            if el.tag is ET.Comment:
                # komentar antar blok = marker
                txt = el.text or ''
                container.remove(el)
                m = RE_VS.search(txt)
                if m:
                    state['vs'].append(float(m.group(1)))
                    continue
                m = RE_ALIGN.search(txt)
                if m:
                    state['align'] = (None if m.group(1) == 'end'
                                      else m.group(1))
                    continue
                if RE_BRK.search(txt):
                    state['brk'] = True
                continue
            if el.tag in CONTENT_TAGS:
                if id(el) in marker_only:
                    absorb_marker_only(el, container)
                else:
                    consume(el)
                    # trailing marker milik blok ini -> pending blok konten
                    # BERIKUTNYA (\newpage tanpa blank line, glue sisa group)
                    for v in trailing_vs.pop(id(el), []):
                        state['vs'].append(v)
                    if id(el) in trailing_brk:
                        state['brk'] = True
                        trailing_brk.discard(id(el))
                if el.tag == f'{{{TEXT_NS}}}list':
                    walk(el)   # marker di dalam item paragraf
                continue
            if el.tag in CONTAINER_TAGS or el.tag == LIST_ITEM_TAG:
                walk(el)       # section/TOC/pustaka/list-item
                continue
            # table & frame: tidak konsumsi pending

    walk(office_text)

    # ---- pass B2: buang paragraf hampa murni (tanpa anak, teks whitespace)----
    # Artefak rantai tex4ht/xtpipes pada dokumen besar (skripsi uji: ~450
    # paragraf '       ' di dalam section env center + ~70 di list-item).
    # LaTeX tidak bisa melahirkan paragraf whitespace-only — aman dibuang.
    # PARENT table:table-cell dikecualikan: ODF mewajibkan >= 1 paragraf
    # per sel tabel.
    for p in list(root.iter(f'{{{TEXT_NS}}}p')):
        if len(p) or (p.text or '').strip():
            continue
        par = parent.get(p)
        if par is None or par.tag == CELL_TAG:
            continue
        par.remove(p)

    # ---- pass C: harden heading (align ditanam cfg -> tiap text:h) ----
    aligns = planted_heading_aligns(zin.read('styles.xml'))
    if aligns:
        for h in root.iter(f'{{{TEXT_NS}}}h'):
            st = h.get(f'{{{TEXT_NS}}}style-name')
            if st in aligns:
                reg.apply(h, {'text-align': aligns[st]})
                stats['hardened'] += 1

    # ---- lapis lama: bungkus inline yatim ----
    # Sekarang juga DI DALAM text:section: env center/quote jalur tex4ht
    # menjadi section berisi <text:p begin-env-p/> + SPAN YATIM + <text:p
    # end-env-p/> — section harus berisi blok (ODDF), tanpa dibungkus
    # LibreOffice membuang span-nya saat export DOCX.
    n_wrapped = 0
    kinds = set()

    def wrap_orphans(container):
        nonlocal n_wrapped
        runs = []

        def flush(next_block):
            nonlocal n_wrapped
            if not runs:
                return
            for el in runs:
                kinds.add(el.tag.split('}')[-1])
                container.remove(el)
            p = ET.Element(f'{{{TEXT_NS}}}p')
            for el in runs:
                p.append(el)
            if next_block is not None:
                container.insert(list(container).index(next_block), p)
            else:
                container.append(p)
            n_wrapped += len(runs)
            runs.clear()

        for el in list(container):
            if el.tag in BLOCK_TAGS:
                flush(el)
            else:
                runs.append(el)
        flush(None)

    wrap_orphans(office_text)
    for sec in list(root.iter(f'{{{TEXT_NS}}}section')):
        wrap_orphans(sec)

    data = {n: zin.read(n) for n in zin.namelist()}
    zin.close()
    out = ET.tostring(root, encoding='unicode').encode('utf-8')
    data['content.xml'] = RE_T2D_BYTES.sub(b'', out)  # strip defensif sisa
    with zipfile.ZipFile(DST, 'w', zipfile.ZIP_DEFLATED) as z:
        for n, b in data.items():
            z.writestr(n, b)
    print(f'[fix_odt] native: margin={stats["margin"]} align={stats["align"]} '
          f'brk={stats["brk"]} hardened={stats["hardened"]} '
          f'(drop: inline={stats["drop_inline"]} vs={stats["drop_vs"]}, '
          f'komentar={stats["comments"]}); {n_wrapped} inline yatim dibungkus '
          f'(jenis: {sorted(kinds)})')


def _is_marker_only(par):
    """True bila text:p tidak punya konten apa pun selain komentar marker
    (teks & tail whitespace saja, anak elemen semuanya komentar) — paragraf
    hampa yang dilahirkan t4ht untuk membawa HCode marker di vmode."""
    if (par.text or '').strip():
        return False
    for ch in par:
        if ch.tag is not ET.Comment:
            return False
        if (ch.tail or '').strip():
            return False
    return True


def _is_leading(par, node):
    """True bila tidak ada konten (teks / elemen) sebelum `node` di `par` —
    marker menempel di AWAL blok -> berlaku untuk blok ini."""
    if (par.text or '').strip():
        return False
    for ch in list(par)[:list(par).index(node)]:
        if ch.tag is not ET.Comment and ((ch.tail or '').strip() or True):
            return False
    return True


def _is_trailing(par, node):
    """True bila tidak ada konten (elemen / teks tail) setelah `node` di `par`
    — yakni marker memang menempel di ujung blok."""
    if (node.tail or '').strip():
        return False
    siblings = list(par)
    for ch in siblings[siblings.index(node) + 1:]:
        if ch.tag is not ET.Comment:
            return False
        if (ch.tail or '').strip():
            return False
    return True


def main(argv=None):
    a = sys.argv[1:] if argv is None else argv
    if len(a) < 2:
        sys.exit('pemakaian: fix_odt <in.odt> <out.odt>')
    run(a[0], a[1])


if __name__ == '__main__':
    main()
