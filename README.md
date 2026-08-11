# MarkdownFactory

Every document becomes Markdown — and back out again as Word. One double-clickable `MarkdownFactory.exe`, four tabs — markitdown-gui, unlimited-ocr-gui, PDF Tools and To Word merged into a single program (MangaFactory-style), built for people who should never have to see a terminal.

![Python](https://img.shields.io/badge/python-3.12-blue) ![License](https://img.shields.io/badge/license-MIT-green) ![Version](https://img.shields.io/badge/version-1.4-orange)

---

## For users (the whole manual)

1. Double-click `MarkdownFactory.exe`. Your browser opens.
2. Everything you make goes into the folder in the **Save to** box at the top — `MarkdownFactory-Export` on your Desktop unless you change it. **Open** shows it.
3. Pick a tab:
   - **Documents** — Word, PowerPoint, Excel, HTML, PDFs with real text. Works immediately on any PC. PDFs go through a dedicated PDF reader that keeps headings, tables, columns and links; drop a PDF that turns out to be a scan and the tab says so instead of handing back an empty file.
   - **Scans & Photos (OCR)** — scanned PDFs, photos of paperwork, screenshots. Needs an NVIDIA graphics card. The first time, click the one **Set up OCR** button and let it download what it needs (a few GB, one time only); after that it just works.
   - **PDF Tools** — **Merge** several PDFs into one (drag the rows to reorder, click a thumbnail to flip through a file first), **Split** one PDF into the pages you want, or **Compress** a scan-heavy PDF down to a sensible size.
   - **To Word** — the way back out. Drop a **PDF** to get an editable `.docx`, or drop a **Markdown** file to turn a note into a properly formatted Word document. Works immediately on any PC. The **Pages** box applies to PDFs only: leave it empty for the whole file, or type `1-3, 5`.
4. Drag files in.

Each Documents/OCR conversion produces **one self-contained `.md` file**. Charts and figures found in scans are stored inside it, so the file still shows its pictures after you move, copy, or email it — no image folder to keep track of. (The OCR tab's **Images** dropdown can switch to a separate folder if you prefer.)

PDF Tools and To Word write their finished files into the same **Save to** folder rather than your Downloads folder, so splitting a 40-page document doesn't bury you in browser download prompts. Nothing is uploaded anywhere — all four engines run entirely on your computer.

**Which tab do I want?** Documents to go *into* Markdown; To Word to come back *out* of it. For a PDF specifically: Documents if the text is headed for notes or a website, To Word if a person has to edit it and send it back. Either way the PDF needs real text in it — if you can't select a sentence with the mouse in a PDF reader, it's a scan, and **Scans & Photos** is the tab that can read it. Documents and To Word both check for this and say so — one instead of writing you an empty `.md`, the other instead of handing you a Word file full of pictures of words.

Markdown files keep their structure: headings become real Word headings (so Word's navigation pane and any table of contents work), lists stay nested, tables become Word tables, links stay clickable, and code sits in a shaded monospaced block. Obsidian's own syntax is handled too — task boxes become ☐/☑, callouts get a bold lead-in, `==highlights==` come out highlighted, and a wikilink keeps its visible label. Two things can't survive the trip: note properties (frontmatter) are dropped, because Word has nowhere to put them, and an embedded image shows as a `[image: name]` placeholder, because only the `.md` file itself was handed over — not the pictures around it.

---

## How it works

A single Python file serves one local page (port 8770) with four panes:

- **Documents tab** runs Microsoft's markitdown **in-process**, and PyInstaller bundles it fully into the exe — zero setup, works offline.

  **PDFs take their own road since 1.4**, through [pdf-inspector](https://github.com/firecrawl/pdf-inspector) — Firecrawl's PDF engine, a Rust core with Python bindings, bundled the same way. It is a local library and nothing else: no model, no API key, no network call, so the promise above that all four engines run on your own computer still holds.
  - It **classifies before it extracts**, in tens of milliseconds. A PDF whose pages are pictures is a scan, and the tab now says so and points at Scans & Photos — markitdown would return an empty string, which arrives as a `.md` that looks converted until you open it. To Word has turned scans away since 1.3; this is Documents catching up, and on every PDF in the test corpus the two tabs agree about which files are scans.
  - What comes out is **position-aware** rather than a flat dump: font-size ratios become real headings, columns are read in the right order, tables come out as tables, and links survive. On a one-page test PDF markitdown found no headings at all; pdf-inspector found the title and both section headings. It is also 20–40× faster (a 28-page report: 137 ms against 3.3 s).
  - A **mixed** PDF — some real pages, some scanned — converts, and a note at the top of the `.md` names the pages that were pictures, so nothing goes missing silently.
  - If the engine isn't aboard, or refuses a file (encrypted, truncated, not really a PDF), the PDF **falls back to markitdown** exactly as in 1.3 and the fallback is noted on the console. 1.3's behaviour is the floor: 1.4 can be better than the old road, never worse. Every other format — docx, pptx, xlsx, html — is untouched and byte-for-byte identical to 1.3.
- **Scans & Photos tab** cannot be bundled the same way (CUDA torch alone is ~4 GB), so it manages its own runtime instead:
  - On startup it looks for a working CUDA Python: first the dev venv (`~/.venvs/unlimited-ocr`), then its own runtime at `%LOCALAPPDATA%\MarkdownFactory\ocr-runtime`.
  - If neither exists it shows a one-click setup card: downloads the official embeddable Python (11 MB), installs pip, then `torch 2.8.0+cu129` + the OCR deps into that private runtime — every step streamed to a progress log on the page. No admin rights, no system Python touched; delete the folder to uninstall.
  - With a runtime available it launches the **Unlimited-OCR GUI server as a background worker** on port 8768 and embeds its page in the tab — so the OCR pane *is* the proven standalone app, checkpointing, ETA, cancel button and all. The worker script is bundled as data in the exe; since 1.1 the copy in this folder (`unlimited-ocr-worker.py`) is the single source of truth for it, because the standalone Unlimited-OCR GUI project has been shelved.
  - No NVIDIA GPU → the tab says so in plain English; the other tabs are unaffected.
  - The ~6 GB model downloads on first OCR use into the standard Hugging Face cache, with progress shown in the pane's banner.
- **PDF Tools tab** (new in 1.1, absorbed from the standalone PDF Tools app) does its work in the browser on **pdf-lib** and **pdf.js**. Those are ~1.9 MB, so they are served from `vendor/` and loaded the first time the tab is opened — the other tabs never pay for them. The finished PDF is posted back to `/save-pdf` and written into the shared export folder; a name that already exists gets a ` (2)` suffix rather than overwriting, same as the Markdown side.
- **To Word tab** (new in 1.3) is bundled into the exe like the Documents tab, so it works offline on any PC with no setup. It takes two kinds of file down two very different roads that share one queue, one row shape and one endpoint.

  **PDFs** run through **pdf2docx**, the one conversion measured in seconds *per page* rather than per file:
  - `POST /convert-word` only *starts* the job and hands back an id; the row then polls `GET /word/state?job=…`. A long PDF therefore can't die on a request timeout, and the browser gets real progress instead of a spinner.
  - pdf2docx offers no progress callback — it narrates itself onto the root logger (`[3/4] Parsing pages...`, `(2/16) Page 2`). A `logging.Handler` reads that back into the job. Conversions hold a lock and run one at a time, which is what makes a single root-logger handler safe (and is honest anyway: the engine is single-threaded, and a queue that pretends otherwise just makes every row slower). Reading and writing are two passes over every page, so each is half the progress bar.
  - A PDF whose pages carry no extractable text is a **scan**. pdf2docx would happily return a Word file of page images, so the tab checks with PyMuPDF first and points at Scans & Photos instead.
  - The **Pages** box takes the same `1-3, 5, 8-10` grammar as the Split tool. It's read at drop time, not at conversion time — otherwise editing it mid-queue would silently change rows that were already waiting.

  **Markdown** is parsed by **markdown-it-py** and written by **python-docx**, which is already aboard for the PDF path:
  - The token stream is flat — open/close markers rather than a tree — so the writer walks it with two small stacks (open lists, blockquote depth) instead of recursing. Headings map to Word's own `Heading 1-6`, lists to `List Bullet`/`List Number` 1–3 (deeper nesting clamps), quotes to `Quote`, tables to `Table Grid`, and code to shaded `Consolas` paragraphs. Links become real `w:hyperlink` relationships, so they stay clickable — including inside table cells.
  - The preset is `commonmark` plus `table` and `strikethrough`, deliberately **not** `gfm-like`: that turns on linkify, which needs a package that isn't aboard and raises at parse time if it's missing.
  - Obsidian's dialect is normalised before parsing, because none of it is CommonMark: frontmatter is stripped, a wikilink collapses to its visible label, and an embed becomes an image node with an angle-bracketed destination (`![x](<Pasted image 1.png>)`) — Obsidian filenames nearly always contain spaces, which a bare CommonMark destination can't hold. Task boxes, callout headers and `==highlights==` are handled at render time so they work anywhere, including inside bold text and table cells.
  - It finishes in milliseconds, so it needs none of the progress machinery — but it takes the same road anyway, for one code path and one row shape, and simply arrives at `done` on the first poll.

  Cost: the exe goes from 74 MB to 134 MB, because pdf2docx reaches OpenCV through `common/algorithm.py` on its main import path and drags NumPy and PyMuPDF with it. None of those three can join the spec's EXCLUDES list. (markdown-it-py is pure Python and rounds to nothing.) Launch time is unchanged at ~2 s — the exe unpacks lazily and both engines are imported only once a file is actually dropped.

Port notes: shell 8770 (scans up to 8779 if busy, so double-launching is harmless), OCR worker 8768.

## The look

Same design language as [MangaFactory](../MangaFactory/README%20v2.0.md) — Space Grotesk over JetBrains Mono, 2px corners, flat underline tabs, hairline-bordered cards, and no gradients, glows or coloured shadows anywhere. The palette is lifted straight off `MarkDownFactory Logo 1.0.png`: black field `#000000`, wordmark red `#ff3131`, frame gold `#ffde59`, rule blue `#38b6ff`. The header is a miniature of the logo — gold frame, red wordmark, blue rule — and since 1.2 the browser-tab favicon is the same three marks drawn at 32px, with stroke weights heavier than the logo's so none of them rasterise away at 16.

Each tab owns one of the logo's marks and re-points a single `--accent` custom property, so every control inside a pane picks up that tab's colour for free: Documents is red, Scans & Photos gold, PDF Tools blue. The logo carries three inks and a page, so the fourth tab takes the page — To Word is `--page` `#f2f0ee`, which is also the only tab whose output is paper-shaped. All four are bright enough that filled buttons take near-black text — white would fail AA on every one. Every text/background pair in the theme measures at or above 4.8:1 — which is why the "(PDFs only)" hint beside **Pages** sits at `--muted` like its label and relies on the parenthetical, rather than a dimmer grey, to read as secondary.

The webfonts load non-render-blocking, so an offline launch never waits on the network; without it the page falls back to Segoe UI and Consolas and everything still lines up. The one seam is the OCR pane, which embeds the worker's own page and keeps the worker's own styling.

## Building the exe

One time: `pip install "markitdown[all]" pdf2docx markdown-it-py pdf-inspector pyinstaller`, then double-click **build.bat** (it checks that the bundled data — `unlimited-ocr-worker.py` and `vendor/` — is present and runs PyInstaller with [MarkdownFactory.spec](MarkdownFactory.spec)). Output: `dist\MarkdownFactory.exe`.

Dev mode, no build needed:

```bash
python "MarkdownFactory 1.4.py"
```

Flags: `--port N`, `--no-browser`, `--ocr-no-model` (start the worker without loading the model — UI testing on GPU-less machines).

## Version history

- **1.4** — **A real PDF engine on the Documents tab.** PDFs now go through [pdf-inspector](https://github.com/firecrawl/pdf-inspector) (Firecrawl's MIT-licensed Rust engine with PyO3 bindings, bundled in-process like the rest — a local library, so no model, no API key and no network call). Two things change. **It classifies before it extracts:** a PDF whose pages are pictures is turned away with the same point-at-Scans-&-Photos sentence To Word has given since 1.3, where markitdown used to hand back a **0-byte `.md`** that looked converted until you opened it. **And what comes out is position-aware:** font-size ratios become real headings, columns are read in order, tables stay tables and links survive — on the test corpus markitdown produced zero headings and zero links on every file, and on a credit-card statement it emitted interleaved-column garbage (`aaccccoouunntt`) where pdf-inspector produced clean words. Extraction is also 20–40× faster (28-page report: 137 ms against 3.3 s). A **mixed** PDF converts and opens with a note naming the pages that were pictures. If the engine is missing or refuses a file (encrypted, truncated, not really a PDF) the PDF **falls back to markitdown** exactly as in 1.3 — verified byte-identical — and every non-PDF format is untouched. The Documents rows also grew the To Word tab's `.note` line, so a rejection is read rather than hovered. Exe 134 MB → 136 MB (+2.34 MiB); launch time unchanged, measured at 3.1 s for both 1.3 and 1.4 on the same machine, since the engine is imported lazily on first use.
- **1.3** — **To Word** tab, the way back out of Markdown. Born-digital PDFs are rebuilt as editable `.docx` by pdf2docx, bundled in-process like the Documents tab, with an optional page range and page-by-page progress polled from a background job; scanned PDFs are detected and sent to the Scans & Photos tab rather than converted into a Word file of page images. Markdown files become properly styled Word documents via markdown-it-py + python-docx — real headings, nested lists, tables, quotes, clickable links, shaded code, and the Obsidian dialect (frontmatter, wikilinks, embeds, callouts, task boxes, `==highlights==`). Exe 74 MB → 134 MB (OpenCV + NumPy + PyMuPDF); launch time unchanged.
- **1.2** — Browser-tab favicon: the logo reduced to the three marks that survive 16px (gold frame, red wordmark bar, blue rule on black), carried as an inline SVG data URI like the other factory apps. The tab previously showed the generic page icon.
- **1.1** — PDF Tools tab (merge / split / compress) absorbed from the standalone app, saving into the shared export folder instead of Downloads; **Save to** moved above the tabs since it now serves two tabs; OCR worker script decoupled from the shelved Unlimited-OCR GUI project; UI rebuilt in MangaFactory's design language on the logo's own palette.
- **1.0** — Documents + Scans & Photos tabs.

## License

MIT. Built on [markitdown](https://github.com/microsoft/markitdown), [pdf-inspector](https://github.com/firecrawl/pdf-inspector), [baidu/Unlimited-OCR](https://github.com/baidu/Unlimited-OCR), [pdf-lib](https://github.com/Hopding/pdf-lib), [pdf.js](https://github.com/mozilla/pdf.js), [pdf2docx](https://github.com/ArtifexSoftware/pdf2docx), [python-docx](https://github.com/python-openxml/python-docx) and [markdown-it-py](https://github.com/executablebooks/markdown-it-py); not affiliated with Microsoft, Firecrawl, Baidu, Mozilla or Artifex.
