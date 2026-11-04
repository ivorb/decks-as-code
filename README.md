# Decks as Code

PowerPoint files are binary blobs. You can't diff them, you can't review them in a
pull request, and you can't reuse the code that already lives three folders over in
the same repo. So every deck gets rebuilt from scratch, and the good one is always
`readout_final_v3_REAL.pptx` in somebody's inbox.

This repo is a working alternative. A deck is one Markdown file. Styling lives in
one CSS file. A build step turns them into a website, a PDF, or a single HTML file
you can email to someone who will double-click it and never know the difference.

Read `decks/decks-as-code/slides.md` to see it. That file is a deck about this,
written in this. It is published at
**<https://ivorb.github.io/decks-as-code/>** by the build in this repo.

## Run it locally

No build step, no `npm install`. Serve the repo over HTTP:

```powershell
cd decks-as-code
python -m http.server 8931
# then open http://localhost:8931/
```

Serving it matters. Opening `index.html` straight off disk won't work, because the
pages fetch `decks.json` and `slides.md` and browsers block that on `file://`. The
built `standalone/` files are exempt, since everything is inlined into them.

**Controls:** `←/→` navigate, `↓` vertical slides, `S` speaker notes,
`Esc` slide overview, `F` fullscreen, `?` all shortcuts.

---

## Layout

```
decks-as-code/
├── index.html                 landing page, lists every deck
├── deck.html                  the player: deck.html?d=<slug>
├── convert.html               .pptx to localizable HTML, in the browser
├── decks.json                 generated listing, don't hand-edit
├── assets/                    theme.css (brand), pptx-slide.css (converted slides),
│                              deck-i18n.js (language swap, fit-to-box)
├── docs/string-table.md       the localization contract
├── tools/
│   ├── build-manifest.mjs     regenerates decks.json
│   ├── build-site.mjs         builds dist/: offline site plus single-file decks
│   ├── import-pptx.mjs        converts a .pptx into a Markdown deck folder
│   ├── pptx-to-html.mjs       converts a .pptx into an HTML deck folder (pptx-html.mjs)
│   ├── translate.mjs          translates a deck's strings.json with an LLM
│   ├── translate-core.mjs     hashing, validation, batching, retry (used by translate)
│   └── llm-provider.mjs       provider interface: copilot, openai-compatible, mock
└── decks/
    └── <slug>/
        ├── slides.md          the deck. The only file you edit day to day.
        ├── deck.json          title, description, author, date, tags, theme
        └── images/            optional, referenced as ![](images/x.png)
        (decks converted to HTML have slides.html and strings.json instead of slides.md)
```

| URL | Serves |
| --- | --- |
| `/` | Landing page listing all decks |
| `/deck.html?d=<slug>` | That deck, full screen |
| `/deck.html?d=<slug>&lang=de` | An HTML deck in a translation (needs `strings.de.json`) |
| `/decks/<slug>/deck.pdf` | PDF export, once built (off by default) |

---

## Add a deck

```powershell
mkdir decks\my-talk
# write decks\my-talk\slides.md, and optionally decks\my-talk\deck.json
node tools\build-manifest.mjs
```

It shows up on the landing page, at `/deck.html?d=my-talk`.

`deck.json` is optional. Without it the deck is titled from its first Markdown
heading. The fields:

```json
{
  "title": "My Talk",
  "description": "One line for the landing page card.",
  "author": "Your Name",
  "date": "2026-09-01",
  "tags": ["readout"],
  "theme": "night"
}
```

`theme` is any stock Reveal.js theme (`night`, `black`, `white`, `league`,
`solarized`, and so on). `assets/theme.css` layers brand styling over whichever you
pick.

---

## Import an existing PowerPoint

You do not have to retype decks you already have. Open the site and click
**Import a .pptx** on the home page, or use the command line:

```powershell
node tools\import-pptx.mjs "C:\path\to\talk.pptx" --slug my-talk
node tools\build-manifest.mjs
```

A `.pptx` is a ZIP of XML, so this reads the file directly. No PowerPoint, no
network, no dependencies.

Both routes run the same converter, so they produce the same `slides.md`. The
browser version does the whole conversion locally — your file is never uploaded —
and lets you preview the result before you keep it. From there you can download a
ZIP to unpack into `decks/`, or, in Chromium browsers, write straight into your
`decks/` folder. The import page needs to be served (see below); it will not work
if you open it from disk.

What comes across:

| Recovered | Notes |
| --- | --- |
| Titles | From the title placeholder |
| Bullets | Nesting levels preserved |
| Bold, italic, links | Inline formatting |
| Tables | Converted to Markdown tables |
| Code | Monospace text boxes become fenced code blocks, indentation intact |
| Images | Written to `decks/<slug>/images/`, SVG preferred over raster |
| Speaker notes | Become `Note:` blocks |
| Slide order | Read from the presentation, not filename order |

What does not, because it is drawing instructions rather than content: layout and
positioning, animations and transitions, charts, SmartArt, themes and masters.
The importer lists the slides it could not fully handle when it finishes.

Treat the output as a first draft. A slide deck built for narration usually has
more text on it than one built for a repo, so expect to cut.

Options:

| Flag | Effect |
| --- | --- |
| `--slug <name>` | Folder name (defaults to the filename) |
| `--title`, `--author`, `--date`, `--theme` | Override `deck.json` fields |
| `--no-images` | Text only, skip the media |
| `--dry-run` | Print the first 40 lines, write nothing |
| `--force` | Overwrite an existing deck folder |
| `--out <dir>` | Write somewhere other than `decks/<slug>` |

---

## Convert a PowerPoint to localizable HTML

A second path, for when the slides must look like the original and every word has to
be translatable. Instead of Markdown, a `.pptx` becomes positioned HTML: one fixed
canvas per slide, shapes placed where PowerPoint put them, and all text kept as real
text.

Open `convert.html` on the served site (or `node tools/pptx-to-html.mjs talk.pptx
--slug my-talk`). It reads the file in your browser, never uploads it, and lets you:

- **Preview** the slides, with speaker notes.
- **Look through** the generated `slides.html` and the string table, and see what the
  converter could not handle.
- **Download** a single HTML file you can double-click, or the deck folder as a ZIP.

The folder holds `slides.html`, `strings.json` and `images/`. `strings.json` lists
every paragraph, speaker note and alt text under a stable ID, in the format described
in [`docs/string-table.md`](docs/string-table.md); translations are `strings.<lang>.json`
files in the same folder, and `?lang=<code>` on the deck or standalone file shows one.
Longer translated text shrinks to fit its box (down to 60%); add `&review` to the
address to outline boxes that still overflow.

What is kept: every paragraph (including hidden-slide text and notes), bold, italic,
underline, links, bullets and numbering, fonts and colors, positions, rotation, groups,
pictures with cropping, tables, gradients, and the common shapes and connectors with
arrowheads. SmartArt is rendered from its stored drawing when the file has one.

What is not: text inside pictures (it cannot be translated), charts (shown as a
placeholder), EMF/WMF pictures (browsers cannot draw them), animations, transitions,
and shadows or other effects. Shapes with an unusual preset are drawn as rectangles.
Everything the converter skipped or approximated is listed under "Needs a look".

Code paragraphs (monospace text) stay on the slide but are left out of the string
table, so nothing translates them.

---

## Translate a deck

`tools/translate.mjs` translates a deck's string table (`decks/<slug>/strings.json`,
see [`docs/string-table.md`](docs/string-table.md)) into `strings.<lang>.json` files.
Node 18+, no dependencies.

```powershell
node tools\translate.mjs my-talk --to de,pt-BR,ar            # translate
node tools\translate.mjs my-talk --to de --dry-run           # show what would change
node tools\translate.mjs my-talk --to de --force             # retranslate everything
node tools\translate.mjs my-talk --to de --provider mock     # offline fake, no LLM
node --test                                                  # run the tests (no network)
```

- **Incremental.** Only strings that are missing or whose source text changed (the
  `src` hash no longer matches) are sent to the model. `--force` redoes all of them.
- **Validated.** Every translation must keep the same inline tags (in matching pairs),
  keep `href` values, and be non-empty. A failing batch is retried once with the error
  fed back; strings that still fail are skipped and listed, the file stays valid, and
  the viewer falls back to the source text for them. Exit code `2` means "done, but
  some strings were skipped"; re-run to retry them.
- **Provider.** `--provider copilot` (default) runs the GitHub Copilot CLI in
  non-interactive mode, so it uses the models in your Copilot subscription: no API key,
  nothing to configure beyond signing in once by running `copilot`. Tools, MCP servers and
  custom instructions are switched off, so the model can only answer. Optional:
  `--model <name>` (otherwise the CLI's default), `LLM_REASONING_EFFORT` (default `low`),
  `COPILOT_BIN` if `copilot` is not on your PATH. The five-slide sample takes under a
  minute per language; a 50-slide deck takes a few minutes.
  `--provider openai` talks to any OpenAI-compatible
  `/chat/completions` endpoint. Configure it with environment variables (the key is
  never a flag, never logged, never written): `LLM_API_KEY` (or `OPENAI_API_KEY`),
  `LLM_MODEL` (or `--model`), and optionally `LLM_BASE_URL` (or `--base-url`; default
  `https://api.openai.com/v1`). For Azure OpenAI / Foundry use the
  `https://<resource>.openai.azure.com/openai/v1` base URL, the deployment name as the
  model, and `LLM_AUTH_HEADER=api-key`. Set `LLM_JSON_MODE=off` if a host rejects
  `response_format`. 429s honour `retry-after`. Providers live in
  `tools/llm-provider.mjs`; adding one means returning `{ name, complete({messages}) }`.
  GitHub Models was retired on 2026-07-30, so it is not supported.

---

## Writing slides

`slides.md` is ordinary Markdown with three separator conventions:

| Syntax | Meaning |
| --- | --- |
| `---` on its own line | New slide |
| `--` on its own line | New vertical sub-slide: optional depth you can skip live |
| `Note:` | Everything after it becomes a speaker note (press `S`) |

Fenced code blocks get real syntax highlighting. `<pre class="mermaid">` blocks
render as diagrams, so there are no image files to keep in sync with the thing they
illustrate.

Per-slide styling uses HTML comment attributes:

```md
<!-- .slide: data-background-color="#1a1a2e" -->
```

---

## Publishing and sharing

`node tools/build-site.mjs` produces `dist/`:

| Output | What it's for |
| --- | --- |
| `dist/standalone/<slug>.html` | A whole deck in one file. Double-click to open. No server, no internet. The easiest thing to email or paste into a chat. |
| `dist/` | The full site with every dependency vendored. Serve it from any static host or file share. Makes no external requests. |
| `dist/decks/<slug>/deck.pdf` | PDF export. |

**`standalone/` is the only part you can double-click.** `index.html` and `deck.html`
read `decks.json` and `slides.md` at runtime, and browsers block those reads on a
`file://` URL, so those two pages need a server even after they are built. Open them
from your file system and they now say so rather than failing silently. To run the
whole site locally, serve `dist/` on a port as shown under "Run it locally".

`.github/workflows/publish-decks.yml` runs on every push and pull request. It builds
`dist/` and uploads the `deck-standalone` artifact. On `main` it also deploys `dist/`
to GitHub Pages.

**PDF export is dormant.** The workflow no longer exports PDFs on push or pull
request. The steps are still there: run the workflow manually (Actions → Publish
decks → Run workflow) and tick `export_pdfs` to get a `deck-pdfs` artifact.

**GitHub Pages.** This repo is public, so Pages is enabled and every push to `main`
republishes <https://ivorb.github.io/decks-as-code/>. Settings → Pages → Source is set
to "GitHub Actions".

**Sharing from a private repo.** If you copy this into a private repo, Pages needs a
paid plan. Artifacts still work: open the latest workflow run, download
`deck-standalone`, unzip, and open the HTML. Anyone with repo access can do the same.

Add a deck without running `build-manifest.mjs` and the `--check` step fails the
build, rather than quietly leaving your deck off the listing.

---

## Dependency versions

Reveal.js, highlight.js, and Mermaid load from jsDelivr at exact pinned versions, set
as constants at the top of `deck.html`. Floating majors like `reveal.js@5` are
avoided on purpose: a minor release can change how an already-approved deck renders.
To upgrade, bump the versions in `deck.html` and `tools/build-site.mjs` together, then
look at the decks.

They load as classic scripts rather than ES modules, also on purpose. decktape drives
the global `Reveal.initialize()` API to export PDFs, and plain script bundles can be
copied into `dist/vendor/`, while Mermaid's ESM build fetches lazy chunks that can't.

`build-site.mjs` downloads those exact files into `dist/vendor/` and caches them in
`.asset-cache/`, so the built output needs no CDN access at all. Handy if jsDelivr is
blocked on your network. It also strips the Google Fonts imports that ship inside the
Reveal themes, since `assets/theme.css` replaces both font stacks anyway.

---

## Why bother

- **Reviewable.** Changes are a readable text diff in a PR, not a 4 MB attachment.
- **Reusable.** Code, configs, and README snippets already in the repo go straight
  into a slide instead of being screenshotted.
- **Consistent.** Branding is `assets/theme.css`. Restyle every deck by editing one
  file.
- **Durable.** Plain text still opens in ten years, and it greps.
- **Live code.** Highlighted, copyable, and correct, instead of a blurry screenshot
  of somebody's editor.

### Trade-offs

- Placing a free-floating element exactly where you want it is harder than dragging a
  box in PowerPoint. This suits structured, content-driven decks better than
  design-led ones.
- Non-technical contributors need a Markdown on-ramp. A template deck and a web
  editor (github.dev works) covers most of it.
- Someone has to own the theme CSS. Budget a small one-time setup.
- PDF export runs headless Chrome. Automated here, but it's one more moving part.

---

## Escape hatches

- **PDF.** Every deck gets one from the build. Locally:
  `npx decktape reveal "http://localhost:8931/deck.html?d=<slug>" deck.pdf`
- **Single file.** `node tools/build-site.mjs` writes a self-contained
  `dist/standalone/<slug>.html`.
- **PPTX.** If someone genuinely needs a `.pptx`, run that deck's `slides.md` through
  Marp: `npx @marp-team/marp-cli slides.md --pptx`
- **Hosting.** It's static HTML. Any file share, intranet host, or GitHub Pages.

---

## The three real options

| Tool | Best for | Effort | Native PPTX export |
| --- | --- | --- | --- |
| Marp | Simplest thing that works. Markdown in, PDF or PPTX out. | Lowest | Yes |
| Reveal.js | Most control, custom HTML and CSS. This setup. | Low to medium | No, PDF only |
| Slidev | Polished themes, live coding, developer audiences | Medium | No, PDF or PNG |

If the main worry is being able to back out later, pilot Marp. If it's control and a
distinctive look, stay on Reveal.js. The `slides.md` files here are close to portable
between the two, mostly differing in front-matter and separators.

---

## Suggested rollout

1. Pilot with one real deck. This one, or an upcoming team readout.
2. Put brand colors, fonts, and a title-slide layout into `assets/theme.css`.
3. Create a `deck-template` repo so starting a deck is one click.
4. Share builds through the workflow artifacts, or enable Pages if the repo goes
   public.
5. Get feedback after two or three decks before pushing it org-wide.
