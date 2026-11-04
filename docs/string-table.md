# String table contract

The interface between the PPTX converter and everything that localizes a deck
(translate CLI, language switcher, review view). Version 1.

The converter emits one **source table** per deck. A translation is a separate
file keyed by string ID. The slide HTML never changes per language; the viewer
swaps text by ID.

## Files

```
decks/<slug>/
├── slides.html            generated; every localizable node carries data-i18n="<id>"
├── strings.json           source table (written by the converter)
└── strings.<lang>.json    one per target language (written by the translate CLI)
```

## Source table: `strings.json`

```json
{
  "version": 1,
  "sourceLang": "en",
  "strings": [
    { "id": "s1.sh2.p1", "slide": 1, "kind": "text", "text": "Decks as <b>Code</b>" }
  ]
}
```

| Field | Notes |
| --- | --- |
| `id` | Unique, stable key, see below. |
| `slide` | 1-based position in presentation order. |
| `kind` | `text`, `table`, `notes`, `alt`. Lets translators treat alt text and notes differently. |
| `text` | One whole paragraph as an HTML fragment. Never a single run. |
| `context` | Optional. Short hint for translators, e.g. the slide title. Never rendered. |

Strings are listed in document order: slide by slide, shapes in z-order, then
notes. Empty paragraphs are not emitted.

### IDs

| Pattern | Source |
| --- | --- |
| `s<slide>.sh<shapeId>.p<n>` | Paragraph `n` (1-based) of the shape whose `cNvPr id` is `shapeId` |
| `s<slide>.sh<shapeId>.r<row>c<col>.p<n>` | Paragraph in a table cell |
| `s<slide>.sh<shapeId>.alt` | Alt text of a shape or picture |
| `s<slide>.L.sh<shapeId>.p<n>` | Text from a layout or master shape drawn on that slide (logos, labels) |
| `s<slide>.notes.p<n>` | Speaker note paragraph |

Shape IDs come from the PPTX, so they survive reordering shapes within a slide.
Slide numbers do not survive reordering slides; re-running the converter after
reordering means re-translating.

Paragraphs set entirely in a monospace font are code. They stay in the HTML but
are left out of the table, so nothing translates them.

### Inline markup in `text`

The only tags allowed are `<b>`, `<i>`, `<u>`, `<sup>`, `<sub>`, `<code>`, `<br>`
and `<a href="...">`. Everything else is escaped (`&lt;`, `&gt;`, `&amp;`).

Translators must keep every tag, in matching pairs, and must not change `href`
values. Tag order may change when word order does. A translation that fails this
check is invalid.

## Translation: `strings.<lang>.json`

```json
{
  "version": 1,
  "lang": "de",
  "sourceLang": "en",
  "strings": {
    "s1.sh2.p1": { "text": "Decks als <b>Code</b>", "src": "9f3a1c2e" }
  }
}
```

- `lang` is a BCP 47 tag (`de`, `pt-BR`, `ar`).
- `src` is the hash of the source `text` the translation was made from: FNV-1a
  32-bit over the UTF-8 bytes, lowercase hex, 8 characters. It is not
  cryptographic. When the source text changes, `src` no longer matches and that
  one string is stale and gets re-translated; the rest are kept.
- A missing or stale string falls back to the source text in the viewer. The
  deck never renders blank.
- Right-to-left languages (`ar`, `he`, `fa`, `ur`) set `dir="rtl"` on the deck
  root, along with `lang`.

## Rules for consumers

1. Never edit `strings.json` by hand. Fix the PPTX and re-run the converter.
2. Validate before accepting a translation: same ID set (extras are errors),
   matching tags and `href`s, non-empty `text`.
3. Text inside images, and charts rendered as images, is not in the table and is
   not translated.

See `docs/examples/` for a small matching pair.
