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
      └──make4ht -f odt -c t2dcfg──> <main>.odt (content + native markers)
                                        │
                     fix_odt (markers → automatic styles) ──LibreOffice──> raw .docx
                                                              │
   raw.docx ──polish──> typography ──graft (PDF oracle)──> <main>_engine.docx
```

## Native semantics layer (v0.3)

The oracle above still has a blind spot: `make4ht` *discards* LaTeX semantics
(`\vspace`, `\vskip`, `\centering`, `\\`) on the way to ODT, so the PDF had to
re-guess what the source said explicitly. v0.3 adds a tex4ht config shipped
inside the package (`t2dcfg.cfg`, loaded via `TEXINPUTS`) that makes tex4ht
**emit those semantics as XML comment markers at the exact point they occur**.
`fix_odt.py` then translates markers into ODT automatic styles *before*
LibreOffice — spacing and alignment are now read natively from the TeX run,
not inferred from PDF geometry:

| LaTeX source | Marker in `content.xml` | Planted as |
|---|---|---|
| `\vspace{x}` `\vspace*{x}` `\bigskip` `\medskip` `\smallskip` `\\[x]` | `<!--t2dvs:14.2pt-->` | `fo:margin-top` on the next content block (consecutive glue **sums**; net < 1 pt dropped) |
| `\vskip x` (intercepted after `\begin{document}`, internal-mode `\vskip` passes through) | `<!--t2dvs:Npt-->` | same |
| `\centering` / `\raggedright` / `\raggedleft` | `<!--t2dalign:center\|left\|right-->` … `<!--t2dalign:end-->` | `fo:text-align` scope until the enclosing group ends (`\aftergroup`; the end-marker is emitted in *every* context — swallowing it lets a chapter-heading `\centering` leak across the whole chapter) |
| `\newpage` `\clearpage` `\cleardoublepage` | `<!--t2dbrk-->` | `fo:break-before="page"` |
| `\\` (plain newline) | — (direct) | `<text:line-break/>` |
| Chapter title alignment | — (style-level) | class intent read from `\meaning\@makechapterhead`, re-issued as `fo:text-align` on Heading-1 + hardened per `text:h` |

Markers are position-classified: a marker at the *head* of a paragraph applies
to that block, at the *tail* to the next block, mid-paragraph (`\\[x]` inside a
line) is dropped — line spacing inside one paragraph has no Word equivalent.
The state machine walks the document **recursively** — into `text:section`
(`center`/`quote`/float environments), lists and index containers — because
that is where most `\centering`/`\vspace` of a real document live; without
the recursion the validation thesis lost all 103 alignment markers. Alignment
scope is **one state across all marker sources** (between-block comments,
paragraph head/tail, marker-only paragraphs): an earlier design kept them in
two variables, so an `align:end` at a paragraph tail never closed a scope
opened between blocks — a chapter head's `\centering` leaked until the end of
the chapter and every `\section` title came out centered (they are left in
the PDF).
Empty paragraphs that t4ht spawns to carry vmode markers are collapsed into the
state machine, not left as blank lines; pure whitespace-only paragraphs (a
tex4ht/xtpipes artifact — ~450 of them inside the validation thesis's
`center` environments) are removed too, except inside table cells where ODF
requires at least one paragraph.

Markers are emitted as **raw DVI specials** (`\ht:special{t4ht=…}`), not via
`\HCode`: `\HCode` executes `\leavevmode` in vertical mode (tex4ht.sty), and a
marker fired at an *internal* `\vskip` (list setup `\vskip\topsep`,
`\endtrivlist`, …) then starts a paragraph mid-list — LaTeX's list state
breaks with `Something's wrong—perhaps a missing \item` (76 in the validation
thesis), `thebibliography` never starts, and the cascade ends in unbalanced
XML. The raw special carries the identical payload with zero effect on TeX's
mode.

The result is **hybrid by default**: native markers win for spacing/alignment;
the PDF oracle keeps what only it can know (pagination, gaps with no marker
such as `\addvspace`/`\vfill`, ≥ 40 pt float regions). In `graft`, a paragraph
that already carries a native `w:before` is skipped by the vertical-spacing
oracle (no double-planting, no zeroing).

### Running without a PDF toolchain

```bash
tex2docx path/to/project --main skripsi --no-pdf
```

Skips `pdflatex` and `graft` entirely — for machines with TeX Live +
LibreOffice but no poppler/pdflatex. Page breaks then come only from
`\newpage`-family markers and tex4ht's per-chapter breaks; spacing and
alignment are fully native.

### Robustness notes (learned the hard way, on the 85-page validation thesis)

- t4ht **can** emit unbalanced XML, and the pipeline no longer judges the run
  by the log alone: on `Unbalanced Tag` it inspects the produced ODT and
  accepts it only if `content.xml` parses as strict XML and the text covers
  ≥ 60 % of the PDF oracle. In practice the big generator of unbalanced output
  on the validation thesis was our own marker layer firing `\HCode` at
  internal `\vskip`s (see above); since markers moved to raw specials the
  thesis compiles with **zero structural errors** — no heal, clean ODF,
  LibreOffice imports it whole. The heal/repair machinery (domfilter's
  HTML-DOM reparse, `fix_odt`'s pre-parse string repair for a stray
  `</text:p>` after `</text:bibliography>` and markers past the closed root)
  stays in as a safety net for unhealed input; mid-document wounds (which
  truncate content — an earlier build came out 748/1625 paragraphs) still
  fail validation and trigger a retry with cleaned t4ht intermediates.
- Two "errors" in the log are **benign noise** on this class of documents:
  `Missing } inserted` at `\end{thebibliography}` (27×; tex4ht's `\bibitem`
  handling recovers on its own — the entries come through fine) and
  `Command \insertchapterspace undefined` (a `\renewcommand` in the class
  over a command only `tocloft`/pdflatex defines). Neither blocks a complete,
  balanced ODT.
- **Silent truncation is measured, not assumed.** In hybrid mode the ODT's
  text content must cover ≥ 60 % of the PDF oracle's text (measured: broken
  builds score 27–55 %, complete ones ~96 %). This catches cut-offs that
  leave balanced XML and no log warning — the validation thesis once stopped
  quietly at a figure on page 290 of 329.
- `\includegraphics{../dir/figure.png}` (paths **outside** the project
  directory) is legal for pdflatex, but tex4ht refuses to *create* its
  `.4og` marker there and aborts at that page — the ODT still comes out
  looking balanced, just silently cut off at the figure (the validation
  thesis lost BAB V–VI this way). The pipeline pre-creates the 0-byte marker
  at the path the document itself references (proactive scan + reactive log
  parse) and retries.
- LibreOffice ODT→DOCX on documents with ~900 embedded Math objects can
  abort mid-import (exit 0, partial DOCX). Block-count validation alone is
  not enough — one abort kept 89 % of blocks and 97 % of text yet dropped
  the bibliography entries — so the pipeline additionally checks that the
  DOCX still contains the fingerprint of the ODT's **last non-table
  paragraph**, and retries with a fresh profile (up to 3×).

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
- **Native markers don't cover everything**: `\addvspace`, `\vfill`,
  class-internal spacing and `\\[x]` *inside* a paragraph line carry no marker
  — in hybrid mode the PDF oracle fills those in, in `--no-pdf` mode they fall
  back to defaults. Mid-chapter page breaks are oracle-only; `--no-pdf` breaks
  come from `\newpage` + chapter starts only.
- **Section-title alignment** (`Heading-2/3`, i.e. `\section`/`\subsection`)
  is not marker-read in v1: hybrid mode keeps the oracle for it, `--no-pdf`
  keeps the class default.
- Right-alignment (signature blocks) is not yet transplanted.
- Bold is only re-planted on paragraphs the centering oracle matched (chapter
  titles, front-matter titles, bylines); run-in bold inside body text comes
  through make4ht as usual.
- Lists of tables/figures without the aux files present are left untouched.

## Status

Developed and validated against an 85-page university thesis (custom UGM
`skripsidtntf.cls`). With the v0.3 native layer the full document converts
end-to-end: make4ht runs with **zero structural errors** (balanced XML at
source, no heal), LibreOffice imports it whole — bibliography entries
included (25/25, previously lost), supervisor-approval table included — and
the final DOCX carries 46 oracle page breaks, 9/9 chapter/front-matter
titles centered *and* bold (BAB I–VI, KATA PENGANTAR, DAFTAR LAMBANG,
DAFTAR PUSTAKA) via the Heading-1 native layer, section titles correctly
**left** (as in the PDF), 139 native alignment scopes, 215 native vertical
spacings, and regenerated TOC/LOT/LOF/nomenclature lists with correct page
numbers. Versus the pre-native build: +5 200 characters of recovered text
(bibliography), centering 220 vs 134 paragraphs, `w:before` 267 vs 190.

## License

MIT
