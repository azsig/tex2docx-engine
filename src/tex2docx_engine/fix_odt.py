"""
fix_odt — Bungkus elemen inline yatim (text:span, text:bookmark, dst.) yang berada
langsung di bawah <office:text> menjadi <text:p>, agar LibreOffice tidak membuangnya
saat konversi ODT->DOCX. Terjadi pada template yang list-environment-nya rusak di
jalur tex4ht (mis. nomencl + \\renewcommand undefined).
"""
import io
import sys
import zipfile
import xml.etree.ElementTree as ET

TEXT_NS = 'urn:oasis:names:tc:opendocument:xmlns:text:1.0'
BLOCK_TAGS = {
    f'{{{TEXT_NS}}}p', f'{{{TEXT_NS}}}h', f'{{{TEXT_NS}}}list',
    f'{{{TEXT_NS}}}section', f'{{{TEXT_NS}}}table-of-content',
    '{urn:oasis:names:tc:opendocument:xmlns:table:1.0}table',
}


def run(SRC, DST):
    zin = zipfile.ZipFile(SRC)
    content = zin.read('content.xml').decode('utf-8')

    for _, (prefix, uri) in ET.iterparse(io.StringIO(content), events=['start-ns']):
        ET.register_namespace(prefix, uri)

    root = ET.fromstring(content)
    office_text = root.find(f'.//{{urn:oasis:names:tc:opendocument:xmlns:office:1.0}}text')
    if office_text is None:
        print('[fix_odt] <office:text> tidak ketemu — tidak ada yang diubah')
        return

    state = {'n_wrapped': 0}
    runs, kinds = [], set()

    def flush(next_block):
        if not runs:
            return
        for el in runs:
            kinds.add(el.tag.split('}')[-1])
            office_text.remove(el)
        p = ET.Element(f'{{{TEXT_NS}}}p')
        for el in runs:
            p.append(el)
        if next_block is not None:
            office_text.insert(list(office_text).index(next_block), p)
        else:
            office_text.append(p)
        state['n_wrapped'] += len(runs)
        runs.clear()

    for el in list(office_text):
        if el.tag in BLOCK_TAGS:
            flush(el)
        else:
            runs.append(el)
    flush(None)

    data = {n: zin.read(n) for n in zin.namelist()}
    zin.close()
    data['content.xml'] = ET.tostring(root, encoding='unicode').encode('utf-8')
    with zipfile.ZipFile(DST, 'w', zipfile.ZIP_DEFLATED) as z:
        for n, b in data.items():
            z.writestr(n, b)
    print(f'[fix_odt] {state["n_wrapped"]} elemen inline yatim dibungkus paragraf (jenis: {sorted(kinds)})')


def main(argv=None):
    a = sys.argv[1:] if argv is None else argv
    if len(a) < 2:
        sys.exit('pemakaian: fix_odt <in.odt> <out.odt>')
    run(a[0], a[1])


if __name__ == '__main__':
    main()
