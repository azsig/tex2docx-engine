# tex2docx-engine

**LaTeX → DOCX conversion that doesn't re-implement LaTeX — it runs it.**

Most LaTeX→Word converters translate LaTeX *syntax* and guess the layout. Give
them a real thesis class (`skripsi`, university templates, custom `.cls`) and the
result falls apart: wrong pagination, lost centering, missing vertical spacing,
broken lists. This engine takes the opposite route:

1. **Run the real engines.** `pdflatex` compiles the document (your `.cls`
   included) into a PDF + `.toc/.lot/.lof/.nls` aux files. `make4ht`/`tex4ht`
   runs the same source again for content + structure, then LibreOffice turns
   it into a DOCX skeleton.
2. **Treat the PDF as the layout oracle.** Whatever your template did —
   `\centering`, `\vskip`, custom macros nobody has heard of — the engine only
   sees its *effect*, measured as geometry: page boundaries, line edges and
   midpoints, y-coordinates.
3. **Graft the measured layout onto the DOCX.** Every layout decision is
   matched to a DOCX paragraph by content identity and planted as a native
   OOXML property (`w:pageBreakBefore`, `w:jc`, `w:spacing`, …).

```
<main>.tex ──pdflatex──> <main>.pdf (layout oracle) + .toc/.lot/.lof/.nls
      │
      └──make4ht -f odt──> <main>.odt ──fix_odt──> ──LibreOffice──> raw .docx
                                                              │
   raw.docx ──polish──> typography ──graft (PDF oracle)──> <main>_engine.docx
```

## The universality rule

**Zero word lists. Zero per-template hardcodes.** Early prototypes detected
bylines via words like *oleh/by*, NIM patterns, *Diajukan/Submitted* — those all
came out. Every formatting decision now comes from one of two sources:

- **PDF geometry** — e.g. a block is centered iff its line left-edges vary and
  midpoints cluster; a vertical gap is deliberate iff it exceeds the page's
  median line advance; a paragraph keeps its first-line indent iff its first
  PDF line starts ~35pt off the margin. Works for any language, any class,
  any macro: the engine reads *what the TeX engine actually did*, not what
  words the document contains.
- **Structural aux files** — TOC/LOT/LOF lists are regenerated from `.toc`
  `.lot` `.lof` (entries + page numbers from the real run), nomenclature
  tables from `.nls` + group names parsed out of your `.cls`.

Remaining template knowledge is one constant (`TOC_HEADS`, the localized names
of the list sections) — configurable, and the engine degrades gracefully
(stale lists stay as-is) when it doesn't match.

## Usage

External requirements: TeX Live (with `make4ht`/`tex4ht`), LibreOffice
(`soffice`), poppler-utils (`pdftotext`). No Python dependencies — stdlib only.

```bash
pip install git+https://github.com/<you>/tex2docx-engine.git

# CLI — compile if needed, then convert
tex2docx path/to/project --main skripsi --out skripsi_engine.docx
```

```python
from tex2docx_engine import convert
convert('path/to/project', main='skripsi', out='skripsi_engine.docx')
```

Each layer is also callable on its own (`tex2docx_engine.polish/typography/graft`),
e.g. when your app already produces the raw DOCX and the PDF separately.

## What the oracle measures

| Oracle | PDF source | Planted as |
|---|---|---|
| Page breaks | page boundaries (`pdftotext -layout`) | `w:pageBreakBefore` |
| Centering | line midpoint clustering in **PDF points** (`pdftotext -bbox`), column = median across pages; exact-text match preferred so regenerated TOC entries (`KATA PENGANTARxiv`) never steal the centering | `w:jc w:val="center"` |
| Bold titles | font weight in the PostScript name (`pdftohtml -xml`: `*-Medi`, `*-Bold`, `CMBX*`) | `<w:b/>` on runs |
| Vertical spacing | y-gap above each line vs page median (`pdftotext -bbox`) | `w:spacing w:before/after` (twips = pt × 20) |
| First-line indent | first line x-offset from margin | strips false `w:firstLine` |
| Lists | `.toc/.lot/.lof/.nls` entries + page numbers | regenerated paragraphs |

## Honest limitations

- **Single-column documents** (theses, reports, books). Two-column classes
  (IEEEtran & friends) interleave columns in `pdftotext` output — not yet
  handled.
- **Mid-paragraph page breaks are skipped** (paragraph integrity is preserved
  at the cost of pagination fidelity there).
- **Vertical gaps ≥ 40pt are skipped** — usually figure/table regions with no
  text in the text bbox, which can't be measured from text geometry.
- Right-alignment (signature blocks) is not yet transplanted.
- Bold is only re-planted on paragraphs the centering oracle matched (chapter
  titles, front-matter titles, bylines); run-in bold inside body text comes
  through make4ht as usual.
- Lists of tables/figures without the aux files present are left untouched.

## Status

Developed and validated against an 85-page university thesis (custom UGM
`skripsidtntf.cls`): 44/53 page breaks planted, all 12 chapter/front-matter
titles (BAB I–VI, KATA PENGANTAR, DAFTAR ISI/TABEL/GAMBAR/LAMBANG, DAFTAR
PUSTAKA) re-centered *and* re-bolded straight from PDF font evidence, all
abstract bylines centered by geometry alone in two languages (id/en), 148
vertical gaps measured from the PDF (0.5 cm `\vspace` → 350 twips,
`\vskip 6pt` → 181 twips, …), regenerated TOC/LOT/LOF/nomenclature lists with
correct page numbers.

## License

MIT
