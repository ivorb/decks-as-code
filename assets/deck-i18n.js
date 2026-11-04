/*
 * Swaps localized text into a deck converted from PowerPoint.
 *
 * Every localizable node carries data-i18n="<id>" (or data-i18n-alt for image
 * alt text). A translation is a table { strings: { "<id>": { text } } }. Anything
 * missing keeps its source text, so a deck never renders blank.
 * See docs/string-table.md.
 *
 * A classic script, shared by deck.html and the standalone single-file build.
 */
(function () {
  var RTL = { ar: 1, he: 1, fa: 1, ur: 1 };
  var ALLOWED = { B: 1, I: 1, U: 1, SUP: 1, SUB: 1, CODE: 1, BR: 1, A: 1 };

  /** Rebuild a translation from the contract's inline tags only; drop everything else. */
  function safeFragment(html) {
    var tpl = document.createElement('template');
    tpl.innerHTML = html;
    var frag = document.createDocumentFragment();
    (function clean(src, dest) {
      for (var i = 0; i < src.childNodes.length; i++) {
        var n = src.childNodes[i];
        if (n.nodeType === 3) dest.appendChild(document.createTextNode(n.textContent));
        else if (n.nodeType === 1 && ALLOWED[n.tagName]) {
          var el = document.createElement(n.tagName);
          if (n.tagName === 'A') {
            var href = n.getAttribute('href') || '';
            if (/^(https?:|mailto:)/i.test(href)) el.setAttribute('href', href);
          }
          clean(n, el);
          dest.appendChild(el);
        } else if (n.nodeType === 1) clean(n, dest);
      }
    })(tpl.content, frag);
    return frag;
  }

  function plainText(html) {
    var tpl = document.createElement('template');
    tpl.innerHTML = html;
    return tpl.content.textContent;
  }

  /** Apply a translation table. `source` restores the original text first. */
  function apply(table, lang) {
    document.querySelectorAll('[data-i18n]').forEach(function (el) {
      if (el.dataset.i18nSrc === undefined) el.dataset.i18nSrc = el.innerHTML;
      var hit = table && table.strings && table.strings[el.dataset.i18n];
      if (hit && hit.text) el.replaceChildren(safeFragment(hit.text));
      else el.innerHTML = el.dataset.i18nSrc;
    });
    document.querySelectorAll('[data-i18n-alt]').forEach(function (el) {
      if (el.dataset.i18nSrc === undefined) el.dataset.i18nSrc = el.alt;
      var hit = table && table.strings && table.strings[el.dataset.i18nAlt];
      el.alt = hit && hit.text ? plainText(hit.text) : el.dataset.i18nSrc;
    });
    if (lang) {
      document.documentElement.lang = lang;
      document.documentElement.dir = RTL[lang.split('-')[0]] ? 'rtl' : 'ltr';
    }
  }

  function langFromUrl() {
    var lang = new URLSearchParams(window.location.search).get('lang');
    return lang && /^[a-z]{2,3}(-[A-Za-z0-9]+)*$/.test(lang) ? lang : null;
  }

  var FIT_FLOOR = 0.6;
  var FIT_STEP = 0.05;

  /**
   * Longer text (a translation, or a fallback font that runs wide) must not spill
   * out of its box and over its neighbours. Paragraph sizes are written as
   * calc(var(--fit, 1) * Npx), so shrinking a box's text is one custom property.
   * Boxes that still overflow at the floor are flagged with data-overflow.
   * Only visible slides can be measured, so this runs per slide on demand.
   */
  function fitSlide(slide) {
    if (!slide) return;
    slide.querySelectorAll('.sh').forEach(function (sh) {
      var tx = sh.querySelector(':scope > .tx');
      if (!tx) return;
      var limit = parseFloat(sh.style.minHeight);
      if (!limit) return;
      tx.style.removeProperty('--fit');
      sh.removeAttribute('data-overflow');
      var scale = 1;
      while (tx.offsetHeight > limit + 1 && scale > FIT_FLOOR) {
        scale = Math.round((scale - FIT_STEP) * 100) / 100;
        tx.style.setProperty('--fit', String(scale));
      }
      if (tx.offsetHeight > limit + 1) sh.setAttribute('data-overflow', 'true');
    });
    slide.dataset.fitted = '1';
  }

  /** Fit the current slide now and on every navigation; call once after Reveal starts. */
  function autoFit(reveal) {
    // ?review outlines boxes whose text still overflows after shrinking.
    if (new URLSearchParams(window.location.search).has('review')) document.documentElement.classList.add('deck-review');
    var run = function () { fitSlide(reveal.getCurrentSlide()); };
    reveal.on('ready', run);
    reveal.on('slidechanged', run);
    reveal.on('resize', function () {
      var slide = reveal.getCurrentSlide();
      if (slide) delete slide.dataset.fitted;
    });
    if (reveal.isReady && reveal.isReady()) run();
  }

  /** Forget measurements, e.g. after the language changed. */
  function refit(reveal) {
    document.querySelectorAll('[data-fitted]').forEach(function (s) { delete s.dataset.fitted; });
    if (reveal) fitSlide(reveal.getCurrentSlide());
  }

  function languageName(code) {
    try {
      var label = new Intl.DisplayNames([code], { type: 'language' }).of(code);
      return label ? label.charAt(0).toUpperCase() + label.slice(1) : code;
    } catch (e) {
      return code;
    }
  }

  /**
   * A language <select> in the corner of the deck. Hidden when there is nothing
   * to choose. Options are named in their own language ("Français"), the source
   * language comes first, and the choice is mirrored into ?lang= so the address
   * can be shared.
   *
   *   source     the deck's own language tag
   *   languages  translation tags that have a table
   *   load       async (tag) => table
   *   reveal     the Reveal instance, to re-fit text after a swap
   */
  function mountPicker(o) {
    var langs = (o.languages || []).filter(function (l) { return l !== o.source; });
    if (!langs.length) return null;

    var box = document.createElement('div');
    box.className = 'deck-lang';
    var select = document.createElement('select');
    select.setAttribute('aria-label', 'Language');
    [o.source].concat(langs).forEach(function (tag) {
      var opt = document.createElement('option');
      opt.value = tag;
      opt.textContent = languageName(tag === o.source ? tag.split('-')[0] : tag);
      select.appendChild(opt);
    });
    box.appendChild(select);
    document.body.appendChild(box);

    var chosen = langFromUrl();
    select.value = chosen && langs.indexOf(chosen) >= 0 ? chosen : o.source;

    select.addEventListener('change', async function () {
      var tag = select.value;
      var table = null;
      if (tag !== o.source) {
        try { table = await o.load(tag); } catch (e) { table = null; }
      }
      apply(table, tag);
      if (o.reveal) refit(o.reveal);
      var url = new URL(window.location.href);
      if (tag === o.source) url.searchParams.delete('lang'); else url.searchParams.set('lang', tag);
      window.history.replaceState(null, '', url);
    });
    return select;
  }

  window.DeckI18n = {
    apply: apply, safeFragment: safeFragment, langFromUrl: langFromUrl, fitSlide: fitSlide, autoFit: autoFit, refit: refit, mountPicker: mountPicker,
  };
})();
