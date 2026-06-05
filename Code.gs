// ─── Menu & Sidebar ──────────────────────────────────────────────────────────

function onOpen() {
  DocumentApp.getUi()
    .createMenu('Thesaurus')
    .addItem('Open Thesaurus', 'showSidebar')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('WordHippo Thesaurus')
    .setWidth(320);
  DocumentApp.getUi().showSidebar(html);
}

// ─── Document Interaction ─────────────────────────────────────────────────────

function getSelectedWord() {
  const selection = DocumentApp.getActiveDocument().getSelection();
  if (!selection) return '';

  const elements = selection.getRangeElements();
  if (!elements.length) return '';

  const el = elements[0];
  if (!el.getElement().asText) return '';

  const text = el.getElement().asText().getText();
  return text.substring(el.getStartOffset(), el.getEndOffsetInclusive() + 1).trim();
}

function replaceSelectedWord(newWord) {
  const selection = DocumentApp.getActiveDocument().getSelection();
  if (!selection) return { success: false, error: 'No text is currently selected.' };

  const elements = selection.getRangeElements();
  if (!elements.length) return { success: false, error: 'No text is currently selected.' };

  const el = elements[0];
  const textEl = el.getElement().asText();
  textEl.deleteText(el.getStartOffset(), el.getEndOffsetInclusive());
  textEl.insertText(el.getStartOffset(), newWord);

  return { success: true };
}

// ─── WordHippo Fetch & Parse ──────────────────────────────────────────────────

const WORDHIPPO_URLS = {
  synonyms:  word => `https://www.wordhippo.com/what-is/another-word-for/${word}.html`,
  antonyms:  word => `https://www.wordhippo.com/what-is/the-opposite-of/${word}.html`,
  rhymes:    word => `https://www.wordhippo.com/what-is/words-that-rhyme-with/${word}.html`,
  sentences: word => `https://www.wordhippo.com/what-is/sentences-with-the-word/${word}.html`,
};

function fetchThesaurus(word, type) {
  const slug = encodeURIComponent(word.toLowerCase().trim().replace(/\s+/g, '-'));
  const urlFn = WORDHIPPO_URLS[type] || WORDHIPPO_URLS.synonyms;
  const url = urlFn(slug);

  // Note: CacheService.getUserCache() was removed because it throws an
  // uncatchable PERMISSION_DENIED ScriptError after re-login. The sidebar's
  // in-memory cache handles repeat lookups within a single session.

  let response;
  try {
    response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      headers: {
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'en-US,en;q=0.5',
      }
    });
  } catch (e) {
    return { error: 'Network error: ' + e.message };
  }

  const code = response.getResponseCode();
  if (code === 404) return { error: `No results found for "${word}".` };
  if (code !== 200) return { error: `WordHippo returned status ${code}.` };

  const html = response.getContentText();
  if (type === 'sentences') return parseSentences(html, word);
  return parseSynonymPage(html);
}

// ─── HTML Parsers ─────────────────────────────────────────────────────────────

function parseSynonymPage(html) {
  // Verified structure from live site:
  //   <div class="wordtype">Adjective</div>
  //   <div class="relatedwords">
  //     <div class="wb"\n><a href="joyful.html">joyful</a></div>
  //     ...
  //   </div>
  // Note: "wordblock" only appears in the CSS, not the HTML body.

  const sections = [];

  // Split on each wordtype div; index 0 is pre-content (CSS/nav), skip it.
  const parts = html.split(/<div[^>]+class=["']wordtype["'][^>]*>/i);

  for (let i = 1; i < parts.length; i++) {
    const part = parts[i];

    // POS label is text before the first </div>.
    const posMatch = part.match(/^([\s\S]*?)<\/div>/i);
    const pos = posMatch ? stripTags(posMatch[1]).trim() : '';

    // Section description, e.g. "A person who lacks sense or judgment".
    const descMatch = part.match(/<div[^>]+class=["']tabdesc["'][^>]*>([\s\S]*?)<\/div>/i);
    const desc = descMatch ? stripTags(descMatch[1]).trim() : '';

    // Each word lives in <div class="wb"><a href="...">word</a></div>.
    // The div tag may have a newline before its closing >, e.g. <div class="wb" \n>.
    const words = [];
    const wbRegex = /<div[^>]+class=["']wb["'][^>]*>[\s\S]*?<a[^>]*>([^<]+)<\/a>[\s\S]*?<\/div>/gi;
    let m;
    while ((m = wbRegex.exec(part)) !== null) {
      const w = m[1].trim();
      if (w && isWordLike(w)) words.push(w);
    }

    if (words.length > 0) {
      sections.push({ pos: pos || 'General', desc, words: dedupe(words).slice(0, 50) });
    }
  }

  if (sections.length === 0) {
    return { error: 'No results found. The page structure may need a parser update.', debugHint: extractDebugHint(html) };
  }

  return { sections };
}

function parseSentences(html, word) {
  const sentences = [];
  // Sentences live in <div class="sentence"> or <li> within a results list.
  const sentenceRegex = /<div[^>]+class=["'][^"']*sentence[^"']*["'][^>]*>([\s\S]*?)<\/div>/gi;
  let m;
  while ((m = sentenceRegex.exec(html)) !== null) {
    const s = stripTags(m[1]).trim();
    if (s.length > 10) sentences.push(s);
    if (sentences.length >= 20) break;
  }
  return sentences.length > 0
    ? { sentences }
    : { error: `No example sentences found for "${word}".` };
}

// ─── Utilities ────────────────────────────────────────────────────────────────

function stripTags(html) {
  return html.replace(/<[^>]+>/g, '').replace(/&amp;/g, '&').replace(/&nbsp;/g, ' ').replace(/&#\d+;/g, '').trim();
}

function isWordLike(str) {
  // Accept single words and short phrases; reject nav links, long strings, numbers.
  if (str.length > 60) return false;
  if (/^\d+$/.test(str)) return false;
  const skipTerms = /^(more|less|back|next|previous|home|about|contact|privacy|terms|all\s|see\s)/i;
  return !skipTerms.test(str);
}

function dedupe(arr) {
  return [...new Set(arr.map(w => w.toLowerCase()))].map(w =>
    arr.find(orig => orig.toLowerCase() === w)
  );
}

// Returns a small snippet of parsed text to help diagnose parser failures.
function extractDebugHint(html) {
  const stripped = stripTags(html).replace(/\s+/g, ' ').substring(0, 400);
  return stripped;
}

// ─── Debug helper (run from Apps Script editor, check Execution Log) ──────────

function debugFetch() {
  const word = 'happy';
  const url = WORDHIPPO_URLS.synonyms(word);
  const response = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    headers: {
      'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
      'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
      'Accept-Language': 'en-US,en;q=0.5',
    }
  });

  Logger.log('HTTP status: ' + response.getResponseCode());

  const html = response.getContentText();
  Logger.log('Page length (chars): ' + html.length);

  // Show 500 chars of context around the first "wordblock" occurrence
  const wbIdx = html.indexOf('wordblock');
  if (wbIdx === -1) {
    Logger.log('wordblock NOT FOUND in HTML — the class name may have changed');
  } else {
    Logger.log('wordblock found at index ' + wbIdx);
    Logger.log('--- 500 CHARS AROUND FIRST wordblock ---');
    Logger.log(html.substring(Math.max(0, wbIdx - 100), wbIdx + 400));
  }

  // Show 500 chars around first "relatedwords"
  const rwIdx = html.indexOf('relatedwords');
  if (rwIdx === -1) {
    Logger.log('relatedwords NOT FOUND in HTML');
  } else {
    Logger.log('--- 500 CHARS AROUND FIRST relatedwords ---');
    Logger.log(html.substring(Math.max(0, rwIdx - 100), rwIdx + 400));
  }

  // Find first actual <div class="wordtype"> HTML tag (CSS uses div.wordtype { — different)
  const wtHtmlIdx = html.indexOf('<div class="wordtype"');
  if (wtHtmlIdx === -1) {
    Logger.log('<div class="wordtype"> NOT FOUND in HTML body');
  } else {
    Logger.log('<div class="wordtype"> found at index ' + wtHtmlIdx);
    Logger.log('--- 800 CHARS AROUND FIRST <div class="wordtype"> ---');
    Logger.log(html.substring(Math.max(0, wtHtmlIdx - 100), wtHtmlIdx + 700));
  }

  // Run the current parser and show results
  const parsed = parseSynonymPage(html);
  Logger.log('--- PARSE RESULT ---');
  Logger.log(JSON.stringify(parsed, null, 2).substring(0, 3000));
}
