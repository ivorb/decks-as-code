# Decks as Code

### A plain Markdown file → a real presentation

Ivor Berry · Learning Content Org

Note: This whole deck is one .md file in git.

---

## Why move off PPTX

- Binary files → useless diffs, painful merges
- Copy/paste drift between decks
- Version "final_v3_REAL.pptx"
- No reuse of content that already lives in repos

---

## What "as code" gives you

| PPTX | Decks as code |
| --- | --- |
| Binary blob | Plain text |
| Email attachments | PR review |
| Manual restyle | One theme file |
| Screenshots of code | Live, highlighted code |

---

## Content is just Markdown

```md
## My slide title

- point one
- point two

---
```

That's it. New slide = `---`.

---

## Real code, really highlighted

```python
def summarize(chunks: list[str]) -> str:
    prompt = "\n".join(chunks)
    return client.responses.create(
        model="gpt-4o-mini",
        input=prompt,
    ).output_text
```

No more blurry screenshots.

--

### Vertical slides work too

Press ↓ for optional depth — skip it live, keep it for the read-later crowd.

---

## Diagrams from text

<pre class="mermaid">
flowchart LR
  A[slides.md] --> B[Build]
  B --> C[Static HTML]
  C --> D[GitHub Pages]
  C --> E[PDF export]
</pre>

---

## Workflow

1. Edit `slides.md`
2. `git commit` → open a PR
3. Teammates review the *text*
4. Merge → auto-published site + PDF

---

## Tooling options

- **Reveal.js** — what you're looking at. Max control, HTML/CSS.
- **Slidev** — Vue-based, dev-friendly, great themes, live coding.
- **Marp** — simplest. Markdown in, PDF/PPTX out.
- **Quarto** — best if slides sit next to data/notebooks.

---

## Escape hatches

- Export to **PDF** for anyone who wants a file
- Marp can export **.pptx** if leadership insists
- Host free on **GitHub Pages**

---

# Questions?

`slides.md` · one file · in git
