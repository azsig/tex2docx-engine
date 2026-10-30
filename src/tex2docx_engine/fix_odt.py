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
    f'{{{TABLE_NS}}}table',
}
# Blok konten yang sah menerima fo:margin-top / fo:text-align / fo:break-before.
CONTENT_TAGS = {f'{{{TEXT_NS}}}p', f'{{{TEXT_NS}}}h', f'{{{TEXT_NS}}}list'}
# Wadah float/TOC: TIDAK mengonsumsi pending — marker menunggu blok konten
# nyata pertama di dalam / setelahnya.
CONTAINER_TAGS = {f'{{{TEXT_NS}}}section', f'{{{TEXT_NS}}}table-of-content'}

RE_VS = re.compile(r't2dvs:(-?[\d.]+)pt')
RE_ALIGN = re.compile(r't2dalign:(center|left|right|end)')
RE_BRK = re.compile(r't2dbrk')
RE_T2D_ANY = re.compile(r't2d(vs|align|brk)')
RE_T2D_BYTES = re.compile(rb'<!--\s*t2d[^>]*-->')


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
    content = zin.read('content.xml').decode('utf-8')

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

    def top_ancestor(el):
        cur = el
        while parent.get(cur) is not None and parent[cur] is not office_text:
            cur = parent[cur]
        return cur

    trailing_brk = set()  # id() blok top-level yang diakhiri <!--t2dbrk-->
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
        if RE_BRK.search(txt) and _is_trailing(par, node):
            block = top_ancestor(node)
            if block is not office_text and block.tag in CONTENT_TAGS:
                trailing_brk.add(id(block))
        elif RE_T2D_ANY.search(txt):
            stats['drop_inline'] += 1  # vs/align intra-blok — dibuang (v1)
        par.remove(node)

    # ---- pass B: state machine atas anak top-level <office:text> ----
    # Marker sebelum blok konten PERTAMA dibuang: halaman 1 sudah mulai
    # (\chapter memicu \clearpage internal + \vspace*{50pt} — keduanya tak
    # bermakna di awal dokumen Word dan break-before berisiko halaman kosong).
    state = {'vs': [], 'align': None, 'brk': False, 'seen': False}

    def target_of(el):
        """text:list tak menerima style paragraf — sasaran paragraf pertama."""
        if el.tag == f'{{{TEXT_NS}}}list':
            for p in el.iter(f'{{{TEXT_NS}}}p'):
                return p
            return None
        return el

    for el in list(office_text):
        if el.tag is not ET.Comment:
            if el.tag in CONTENT_TAGS:
                tgt = target_of(el)
                props = {}
                if state['seen']:
                    if state['vs']:
                        net = sum(state['vs'])
                        if net >= 1.0:
                            props['margin-top'] = f'{round(net, 3)}pt'
                            stats['margin'] += 1
                        else:
                            stats['drop_vs'] += 1
                    if state['brk']:
                        props['break-before'] = 'page'
                        stats['brk'] += 1
                state['seen'] = True
                if state['align']:
                    props['text-align'] = state['align']
                    stats['align'] += 1
                if tgt is not None:
                    reg.apply(tgt, props)
                state['vs'] = []
                state['brk'] = False
                # align TIDAK direset: scope berlaku sampai t2dalign:end
            elif (el.tag in CONTAINER_TAGS and state['align']):
                # env \begin{center}/quote -> text:section; scope align
                # membungkus section -> perataan berlaku ke paragraf di
                # dalamnya (container tak menerima style paragraf sendiri).
                for p in el.iter(f'{{{TEXT_NS}}}p'):
                    reg.apply(p, {'text-align': state['align']})
                    stats['align'] += 1
            # container (section/TOC) & table & frame: tidak konsumsi pending
            continue
        # komentar top-level = marker antar blok
        txt = el.text or ''
        office_text.remove(el)
        m = RE_VS.search(txt)
        if m:
            state['vs'].append(float(m.group(1)))
            continue
        m = RE_ALIGN.search(txt)
        if m:
            state['align'] = None if m.group(1) == 'end' else m.group(1)
            continue
        if RE_BRK.search(txt):
            state['brk'] = True

    # \newpage di ujung paragraf -> break untuk blok konten BERIKUTNYA
    # (\newpage user kerap tanpa blank line; marker jatuh di dalam <text:p>
    # terakhir, tapi maksudnya jelas: konten selanjutnya di halaman baru.)
    prev = None
    for el in list(office_text):
        if el.tag in CONTENT_TAGS:
            if prev is not None and id(prev) in trailing_brk:
                reg.apply(el, {'break-before': 'page'})
                stats['brk'] += 1
                trailing_brk.discard(id(prev))
            prev = el

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
