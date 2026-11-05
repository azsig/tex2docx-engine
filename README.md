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

### Interaction mode (`--interaction`, default `nonstopmode`)

Same contract as the real TeX flag, applied to every engine the pipeline runs:

- **`nonstopmode`** (default) — recoverable TeX errors don't kill the
  conversion. This is what `-interaction=nonstopmode` actually does: pdflatex
  recovers and still writes the PDF *while exiting 1*, so exit codes are only
  reported as warnings; the final artifacts (PDF, ODT, DOCX) decide whether the
  run passed. A fatal failure (no PDF at all) is diagnosed from the tail of
  `<main>.log`.
- **`errorstopmode`** — strict: the first nonzero engine exit aborts, like the
  interactive TeX stopping at the first error.
- **`batchmode` / `scrollmode`** — same tolerance as nonstopmode; the mode is
  passed through to the engines verbatim.

The mode reaches `pdflatex` as `-interaction=<mode>` and also the LaTeX run
*inside* `make4ht` — which has no CLI flag for it and silently uses `batchmode`.
The engine injects it through a generated make4ht config (`XDG_CONFIG_HOME`
points into the engine's temp dir; your own `~/.config/make4ht/config.lua`
content is preserved below the injected assignment, so an explicit preference
of yours still wins). One carve-out: if a `.make4ht` file exists in the project
(or a parent), make4ht prioritizes it and ignores the XDG config — the pipeline
prints a notice and leaves that file in charge; `pdflatex` still honors the
flag.

**Why the mode alone isn't the whole story:** the make4ht chain has stages
after LaTeX that don't understand `nonstopmode`. For every graphic in the DVI,
`tex4ht` `fopen`s its `.4og` instruction file and *dies hard*
(`exit(EXIT_FAILURE)` in tex4ht.c) when that open fails — a figure **directory
that doesn't exist** (typo'd path, missing folder) or a path **outside** the
project directory, both legal for pdflatex. The engine pre-creates the 0-byte
marker (+ its directory) at every path the document references (proactive scan
of `\includegraphics` + reactive log parse) so the conversion keeps going past
the bad figure, like pdflatex does. `t4ht`, the packaging stage, never
hard-exits on images by itself and is run with `-g` so a failed image
conversion doesn't stop the operations after it.

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
