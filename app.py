import streamlit as st

# Set page config must be the very first Streamlit command
st.set_page_config(page_title="Presify - PDF to PowerPoint", layout="wide")

# ---------------------------------------------------------------------------
# SSL workaround for NLTK downloads only (kept local, no global side effects)
# ---------------------------------------------------------------------------
import ssl
import nltk


def _safe_nltk_download(pkg: str) -> None:
    try:
        nltk.download(pkg, quiet=True)
    except Exception:
        # Try once more with unverified SSL (for corporate / outdated certs)
        try:
            _orig = ssl._create_default_https_context
            ssl._create_default_https_context = ssl._create_unverified_context
            nltk.download(pkg, quiet=True)
            ssl._create_default_https_context = _orig
        except Exception:
            pass


import io
import re
import base64
import traceback
from collections import Counter
from datetime import datetime

import PyPDF2
from nltk.corpus import stopwords
from nltk.tokenize import sent_tokenize, word_tokenize
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from sklearn.feature_extraction.text import TfidfVectorizer


# ---------------------------------------------------------------------------
# NLTK bootstrap (with graceful fallback)
# ---------------------------------------------------------------------------
@st.cache_resource
def bootstrap_nltk():
    needed = [
        ("tokenizers/punkt", "punkt"),
        ("tokenizers/punkt_tab", "punkt_tab"),  # NLTK >= 3.8.2
        ("corpora/stopwords", "stopwords"),
    ]
    for path, pkg in needed:
        try:
            nltk.data.find(path)
        except LookupError:
            _safe_nltk_download(pkg)
    try:
        sent_tokenize("Hello world. This is a test.")
        stopwords.words("english")
        return True
    except Exception:
        return False


NLTK_READY = bootstrap_nltk()
STOPWORDS = set(stopwords.words("english")) if NLTK_READY else set(
    "a an the and or but if while of in on at to from by for with as is are was were "
    "be been being this that these those it its i you he she we they them us our "
    "your their his her not no so do does did done have has had can could should would "
    "may might will just also more most much many such only over under into out about".split()
)


# ---------------------------------------------------------------------------
# Themes
# ---------------------------------------------------------------------------
THEMES = {
    "Midnight Mint": {
        "bg": RGBColor(0x10, 0x10, 0x2E),
        "panel": RGBColor(0x1A, 0x1A, 0x4A),
        "primary": RGBColor(0x7D, 0xE6, 0xBC),     # mint
        "accent": RGBColor(0xFC, 0x81, 0x6E),      # coral
        "text": RGBColor(0xFF, 0xFF, 0xFF),
        "muted": RGBColor(0xB8, 0xB8, 0xD4),
    },
    "Royal Gold": {
        "bg": RGBColor(0x16, 0x1B, 0x33),
        "panel": RGBColor(0x21, 0x29, 0x4A),
        "primary": RGBColor(0xF5, 0xC8, 0x6B),     # gold
        "accent": RGBColor(0xE8, 0x6A, 0x6A),      # red-coral
        "text": RGBColor(0xFA, 0xFA, 0xFA),
        "muted": RGBColor(0xBE, 0xC3, 0xD6),
    },
    "Minimal Light": {
        "bg": RGBColor(0xFA, 0xFA, 0xF7),
        "panel": RGBColor(0xEE, 0xEC, 0xE4),
        "primary": RGBColor(0x1F, 0x3A, 0x5F),     # navy
        "accent": RGBColor(0xC0, 0x55, 0x3C),      # terracotta
        "text": RGBColor(0x22, 0x22, 0x22),
        "muted": RGBColor(0x66, 0x66, 0x66),
    },
    "Ocean Breeze": {
        "bg": RGBColor(0x0B, 0x1E, 0x3F),
        "panel": RGBColor(0x14, 0x2A, 0x52),
        "primary": RGBColor(0x6E, 0xC6, 0xFF),     # sky
        "accent": RGBColor(0xFF, 0xB7, 0x4D),      # amber
        "text": RGBColor(0xFF, 0xFF, 0xFF),
        "muted": RGBColor(0xAE, 0xC4, 0xE0),
    },
}


# ---------------------------------------------------------------------------
# PDF text extraction
# ---------------------------------------------------------------------------
def extract_text_from_pdf(pdf_file):
    try:
        pdf_file.seek(0)
        reader = PyPDF2.PdfReader(pdf_file)
        pages = []
        for page in reader.pages:
            try:
                pages.append(page.extract_text() or "")
            except Exception:
                pages.append("")
        meta = {"title": "", "author": "", "page_count": len(pages)}
        try:
            info = reader.metadata or {}
            if info:
                meta["title"] = (info.get("/Title") or "").strip()
                meta["author"] = (info.get("/Author") or "").strip()
        except Exception:
            pass
        return "\n".join(pages), meta
    except Exception as e:
        st.error(f"Error extracting text from PDF: {e}")
        return "", {"title": "", "author": "", "page_count": 0}


def clean_text(text: str) -> str:
    text = re.sub(r"-\n", "", text)              # join hyphenated wraps
    text = re.sub(r"[ \t]+", " ", text)
    lines = []
    for ln in text.splitlines():
        s = ln.strip()
        if not s:
            continue
        # Skip page-number / running-header only lines
        if re.fullmatch(r"(page\s*)?\d{1,4}(\s*/\s*\d{1,4})?", s.lower()):
            continue
        lines.append(s)
    return re.sub(r"\s+", " ", " ".join(lines)).strip()


# ---------------------------------------------------------------------------
# Sentence + keyword utilities
# ---------------------------------------------------------------------------
_SENT_FALLBACK = re.compile(r"(?<=[.!?])\s+(?=[A-Z0-9])")


def split_sentences(text: str):
    text = text.strip()
    if not text:
        return []
    if NLTK_READY:
        try:
            return [s.strip() for s in sent_tokenize(text) if s.strip()]
        except Exception:
            pass
    return [s.strip() for s in _SENT_FALLBACK.split(text) if s.strip()]


def tokenize_words(text: str):
    if NLTK_READY:
        try:
            return [w.lower() for w in word_tokenize(text) if w.isalpha()]
        except Exception:
            pass
    return [w.lower() for w in re.findall(r"[A-Za-z][A-Za-z\-']+", text)]


def top_keywords(text: str, n: int = 8):
    words = [w for w in tokenize_words(text) if w not in STOPWORDS and len(w) > 2]
    if not words:
        return []
    return [w for w, _ in Counter(words).most_common(n)]


def titlecase_phrase(phrase: str) -> str:
    small = {"of", "the", "a", "an", "and", "or", "in", "on", "for", "to", "with", "by"}
    out = []
    for i, p in enumerate(phrase.split()):
        out.append(p if (i > 0 and p.lower() in small) else p.capitalize())
    return " ".join(out)


# ---------------------------------------------------------------------------
# Title detection
# ---------------------------------------------------------------------------
def detect_title(text: str, meta: dict, fallback_filename: str) -> str:
    if meta.get("title"):
        t = meta["title"].strip()
        if 3 < len(t) < 120:
            return t

    head = text[:1500]
    for ln in head.split("."):
        ln = ln.strip()
        if 5 < len(ln) < 100:
            words = ln.split()
            cap_ratio = sum(1 for w in words if w[:1].isupper()) / max(1, len(words))
            if cap_ratio >= 0.5 and len(words) <= 14:
                return ln

    base = re.sub(r"[_\-]+", " ", fallback_filename.rsplit(".", 1)[0])
    return titlecase_phrase(base).strip() or "Document Summary"


# ---------------------------------------------------------------------------
# Section detection & segmentation
# ---------------------------------------------------------------------------
HEADING_PATTERNS = [
    re.compile(r"^\s*(?:chapter|section|part)\s+[\dIVXLCM]+\s*[:\-\.\u2013\u2014]?\s*(.+?)\s*$", re.I),
    re.compile(r"^\s*\d+(?:\.\d+)*\s+([A-Z][A-Za-z0-9 ,&'\-/]{2,80})\s*$"),
    re.compile(r"^\s*([A-Z][A-Z0-9 ,&'\-/]{3,60})\s*$"),  # ALL CAPS HEADING line
    re.compile(
        r"^\s*(introduction|background|overview|abstract|summary|methodology|methods?|approach|"
        r"results?|findings?|analysis|discussion|implications?|recommendations?|conclusions?|"
        r"references?|literature\s+review|case\s+study|key\s+takeaways?)\s*[:\-]?\s*$",
        re.I,
    ),
]


def find_heading_positions(raw_text: str):
    """Return list of (start_idx, end_idx_of_line, heading_text) sorted by position."""
    positions = []
    for m in re.finditer(r"[^\n]+", raw_text):
        line = m.group(0).strip()
        if not line or len(line) > 120:
            continue
        for pat in HEADING_PATTERNS:
            hm = pat.match(line)
            if hm:
                heading = (hm.group(1) if hm.groups() else line).strip(" :-\u2013\u2014").strip()
                heading = re.sub(r"\s+", " ", heading)
                if 3 <= len(heading) <= 80:
                    positions.append((m.start(), m.end(), titlecase_phrase(heading)))
                break
    deduped, seen = [], set()
    for s, e, h in positions:
        key = h.lower()
        if key in seen:
            continue
        seen.add(key)
        deduped.append((s, e, h))
    return deduped


def segment_by_headings(raw_text: str):
    headings = find_heading_positions(raw_text)
    if len(headings) < 3:
        return []
    sections = []
    for i, (s, e, h) in enumerate(headings):
        body_start = e
        body_end = headings[i + 1][0] if i + 1 < len(headings) else len(raw_text)
        body = clean_text(raw_text[body_start:body_end])
        if len(body.split()) >= 25:
            sections.append((h, body))
    if len(sections) > 8:
        # Keep the 8 most substantial, then restore document order.
        order_idx = {h: i for i, (h, _) in enumerate(sections)}
        sections.sort(key=lambda x: len(x[1]), reverse=True)
        sections = sections[:8]
        sections.sort(key=lambda x: order_idx.get(x[0], 999))
    return sections


def auto_segment_by_topics(cleaned: str, target_sections: int = 6):
    """Fallback when no real headings exist: split sentences into ~equal chunks,
    label each chunk with its distinctive TF-IDF keywords."""
    sentences = split_sentences(cleaned)
    if not sentences:
        return [("Overview", cleaned)]
    target_sections = max(3, min(target_sections, max(3, len(sentences) // 10)))
    chunk_size = max(1, len(sentences) // target_sections)
    chunks = []
    for i in range(0, len(sentences), chunk_size):
        chunk = " ".join(sentences[i:i + chunk_size]).strip()
        if chunk:
            chunks.append(chunk)
    chunks = chunks[:target_sections]

    labels = []
    try:
        vec = TfidfVectorizer(
            stop_words="english",
            max_features=400,
            ngram_range=(1, 2),
            token_pattern=r"(?u)\b[A-Za-z][A-Za-z\-']{2,}\b",
        )
        mat = vec.fit_transform(chunks)
        feats = vec.get_feature_names_out()
        for i in range(mat.shape[0]):
            row = mat.getrow(i).toarray().ravel()
            top_idx = row.argsort()[::-1][:4]
            terms = [feats[j] for j in top_idx if row[j] > 0]
            labels.append(titlecase_phrase(" & ".join(terms[:2])) if terms else f"Topic {i + 1}")
    except Exception:
        labels = [f"Topic {i + 1}" for i in range(len(chunks))]

    if labels:
        labels[0] = "Introduction & Overview"
    if len(labels) > 1:
        labels[-1] = "Conclusions & Outlook"
    return list(zip(labels, chunks))


# ---------------------------------------------------------------------------
# Extractive summarisation per section (meeting-ready content)
# ---------------------------------------------------------------------------
BOILERPLATE_RE = re.compile(
    r"^(figure|fig\.?|table|see\s+|refer\s+to|references?|bibliography|appendix|"
    r"copyright|©|all\s+rights\s+reserved|http[s]?://|www\.)",
    re.I,
)

# Inline citations: (Smith et al., 2020), (Smith, 2020, p.5), [12], [1-3], (Fig. 3), (p. 5), (2020)
_CITATION_RE = re.compile(
    r"\s*(?:"
    r"\(\s*[A-Z][A-Za-z\-']+(?:\s+et\s+al\.?|\s+(?:and|&)\s+[A-Z][A-Za-z\-']+)?"
    r"\s*,?\s*\d{4}[a-z]?(?:\s*,\s*p+\.?\s*\d+(?:\s*[-\u2013]\s*\d+)?)?\s*\)"
    r"|\[\s*\d{1,3}(?:\s*[\-,\u2013]\s*\d{1,3})*\s*\]"
    r"|\((?:Fig(?:ure)?|Table|Tab\.?|Eq(?:uation)?|Section|Sec\.?|Chapter|Ch\.?|Appendix|App\.?)\s*[\dIVX\.\-]+\)"
    r"|\(p+\.?\s*\d+(?:\s*[-\u2013]\s*\d+)?\)"
    r"|\(\s*\d{4}[a-z]?\s*\)"
    r"|\(\s*(?:see(?:\s+also)?|cf\.?)\s+[^)]{1,80}\)"
    r")",
    re.I,
)

# Filler / hedging openings that don't add information in a bullet
_LEAD_FILLER_RE = re.compile(
    r"^\s*(?:"
    # Discourse connectives
    r"however|moreover|furthermore|additionally|in\s+addition|also|besides|"
    r"therefore|thus|hence|consequently|as\s+a\s+result|"
    r"in\s+conclusion|to\s+conclude|to\s+summari[sz]e|in\s+summary|overall|on\s+the\s+whole|"
    r"in\s+particular|specifically|notably|importantly|interestingly|clearly|essentially|"
    # "It is/was/has been/should be/can be ... [to do] that"
    r"it\s+(?:is|was|has\s+been|should\s+be|can\s+be|may\s+be|must\s+be)\s+"
    r"(?:important|worth|note[d]?|clear|evident|interesting|essential|critical|crucial|useful|"
    r"shown|seen|observed|argued|noted|found|established|recognised|recognized)"
    r"(?:\s+(?:to\s+(?:note|mention|emphasi[sz]e|highlight|point\s+out|observe)))?\s+that|"
    # Bare "Note that / Recall that"
    r"note\s+that|notice\s+that|please\s+note\s+that|recall\s+that|"
    # "As shown above/below/in Figure N/Table N"
    r"as\s+(?:noted|mentioned|stated|shown|discussed|seen|described|illustrated|reported|depicted)"
    r"(?:\s+(?:above|below|earlier|previously|before|"
    r"in\s+(?:the\s+)?(?:above|preceding|following|figure|fig\.?|table|tab\.?|section|sec\.?|chapter|ch\.?)\s*[\dIVX\.\-]*))?,?|"
    # "(As shown )?in Figure N," / "In Table 2,"
    r"in\s+(?:figure|fig\.?|table|tab\.?|equation|eq\.?|section|sec\.?|chapter|ch\.?|appendix|app\.?)\s*[\dIVX\.\-]+,?|"
    # "In this paper / In the present study"
    r"in\s+(?:this|the\s+(?:current|present))\s+(?:paper|study|report|article|section|chapter|document|work|analysis),?|"
    # First-person prefixes: "We <verb>", "The authors <verb>", "This paper <verb>", optionally trailing "that"
    r"we\s+(?:will\s+|now\s+|also\s+|further\s+|hereby\s+)?[a-z]+(?:\s+that)?|"
    r"the\s+authors?\s+(?:also\s+|further\s+)?[a-z]+(?:\s+that)?|"
    r"this\s+(?:paper|study|report|article|section|chapter|document|work|analysis)\s+"
    r"(?:also\s+|further\s+|now\s+)?[a-z]+(?:\s+that)?|"
    # "The purpose/aim/goal/objective of this X is to"
    r"the\s+(?:purpose|aim|goal|objective)\s+of\s+(?:this|our|the\s+(?:current|present))\s+[a-z]+\s+(?:is|was)\s+to"
    r")\b[,:\s]*",
    re.I,
)

# Residue cleanup: a stray "that " left behind after removing "argues that"
_RESIDUAL_THAT_RE = re.compile(r"^\s*that\s+", re.I)

# Sentences whose meaning depends on an antecedent that's not in the bullet
_ANAPHORA_BAD_RE = re.compile(r"^\s*(?:this|that|these|those|it|such|they|he|she)\s+(?:is|are|was|were|has|have|will|would|may|might|can|could|should|do|does|did|seems|appears|provides|shows|gives|offers|leads|results|enables|allows|requires)\b", re.I)
_REFS_BAD_RE = re.compile(r"\b(?:as\s+shown\s+in|see\s+(?:above|below|table|figure|fig\.?|tab\.?)|in\s+the\s+(?:above|preceding|following)|the\s+(?:above|below|following|preceding)\s+(?:table|figure|equation|section))\b", re.I)

# Punchy fact patterns: percentages, currency, large numbers, years, units
_STAT_RE = re.compile(
    r"(\d+(?:[\.,]\d+)?\s*%"
    r"|\$\s?\d[\d,\.]*"
    r"|\b(?:19|20)\d{2}(?:[-\u2013](?:19|20)?\d{2})?\b"
    r"|\b\d+(?:[\.,]\d+)?\s*(?:million|billion|trillion|thousand|gw|mw|kw|tw|km|kg|tons?|tonnes?|mph|kph|°c|°f)\b"
    r"|\b\d+(?:[\.,]\d+)?\s*x\b"
    r"|\bfold\b)",
    re.I,
)


def polish_sentence(s: str) -> str:
    """Clean a raw sentence into meeting-ready phrasing.
    Strips citations, normalises punctuation, removes hedging/filler openers."""
    if not s:
        return ""
    s = s.replace("\u2018", "'").replace("\u2019", "'")
    s = s.replace("\u201C", '"').replace("\u201D", '"')
    s = s.replace("\u2013", "-").replace("\u2014", " - ")
    # Strip inline citations (repeat to catch adjacent ones)
    prev = None
    while prev != s:
        prev = s
        s = _CITATION_RE.sub("", s)
    # Strip stacked filler openers (e.g. "Moreover, it is important to note that ...")
    prev = None
    while prev != s:
        prev = s
        s = _LEAD_FILLER_RE.sub("", s).strip()
    # Drop residual leading "that " ("argues that policy ..." -> "policy ...")
    s = _RESIDUAL_THAT_RE.sub("", s).strip()
    # Tidy spacing & trailing punctuation
    s = re.sub(r"\s+([,.;:?!])", r"\1", s)
    s = re.sub(r"\s+", " ", s).strip()
    s = s.strip(" .,;:")
    if s:
        s = s[0].upper() + s[1:]
    return s


def _sentence_quality(sentence: str) -> float:
    n = len(sentence)
    if n < 35 or n > 320:
        return 0.0
    if BOILERPLATE_RE.match(sentence):
        return 0.0
    alpha = sum(c.isalpha() for c in sentence)
    if alpha / max(1, n) < 0.55:
        return 0.0
    if 80 <= n <= 180:
        return 1.0
    return 0.7


def _passes_quality(s: str) -> bool:
    """Stricter gate for meeting-ready bullets."""
    if not s or len(s) < 30:
        return False
    if _ANAPHORA_BAD_RE.match(s):
        return False
    if _REFS_BAD_RE.search(s):
        return False
    words = s.split()
    if len(words) < 6:
        return False
    # Require a verb-like content: at least one lowercase letter run
    if not re.search(r"[a-z]{3,}", s):
        return False
    return True


def _stat_bonus(s: str) -> float:
    matches = _STAT_RE.findall(s)
    return min(0.6, 0.18 * len(matches))


def _textrank_scores(tfidf):
    """PageRank over cosine-similarity sentence graph. Returns (scores, sim_matrix)."""
    import numpy as np
    from sklearn.metrics.pairwise import cosine_similarity

    n = tfidf.shape[0]
    if n == 0:
        return np.zeros(0), np.zeros((0, 0))
    sim = cosine_similarity(tfidf)
    np.fill_diagonal(sim, 0.0)
    row_sums = sim.sum(axis=1, keepdims=True)
    row_sums[row_sums == 0] = 1.0
    M = sim / row_sums
    scores = np.ones(n) / n
    teleport = np.ones(n) / n
    damping = 0.85
    for _ in range(30):
        new_scores = (1 - damping) * teleport + damping * (M.T @ scores)
        if np.allclose(new_scores, scores, atol=1e-6):
            scores = new_scores
            break
        scores = new_scores
    return scores, sim


def _mmr_select(rel_scores, sim, k, lambda_=0.7):
    """Maximal Marginal Relevance: pick k items balancing relevance and diversity."""
    import numpy as np

    n = len(rel_scores)
    if n == 0 or k <= 0:
        return []
    selected = [int(np.argmax(rel_scores))]
    while len(selected) < k:
        best_i, best_v = None, -1e18
        for i in range(n):
            if i in selected:
                continue
            redundancy = max(sim[i, j] for j in selected) if selected else 0.0
            score = lambda_ * rel_scores[i] - (1 - lambda_) * redundancy
            if score > best_v:
                best_v, best_i = score, i
        if best_i is None:
            break
        selected.append(best_i)
    return selected


def summarise_section(section_text: str, num_points: int = 5):
    """TextRank-centrality + MMR diversity + polish + stat boost.
    Produces a small set of high-signal, meeting-ready sentences."""
    sentences = split_sentences(section_text)
    if not sentences:
        return []

    # Polish first, then quality-gate
    polished = [(i, polish_sentence(s)) for i, s in enumerate(sentences)]
    candidates = [(i, p) for i, p in polished if _passes_quality(p) and _sentence_quality(p) > 0]

    # Loosen the gate if too few survive
    if len(candidates) < max(2, num_points // 2):
        candidates = [(i, p) for i, p in polished if p and len(p) >= 30 and not _ANAPHORA_BAD_RE.match(p)]
    if not candidates:
        candidates = [(i, p) for i, p in polished if p and len(p) >= 25]
    if not candidates:
        return []

    if len(candidates) <= num_points:
        # Preserve document order
        candidates.sort(key=lambda x: x[0])
        return [p for _, p in candidates]

    texts = [p for _, p in candidates]

    try:
        import numpy as np

        vec = TfidfVectorizer(
            stop_words="english",
            max_features=800,
            ngram_range=(1, 2),
            token_pattern=r"(?u)\b[A-Za-z][A-Za-z\-']{2,}\b",
            sublinear_tf=True,
        )
        tfidf = vec.fit_transform(texts)
        tr_scores, sim = _textrank_scores(tfidf)

        # Combine centrality + stat bonus + length sweet spot
        combined = np.zeros(len(candidates))
        last_idx = len(candidates) - 1
        for k, (_, p) in enumerate(candidates):
            score = float(tr_scores[k]) + _stat_bonus(p) + 0.08 * _sentence_quality(p)
            if k == 0 or k == last_idx:
                score *= 1.08  # mild boost for topic / closing sentence
            combined[k] = score

        picked_idx = _mmr_select(combined, sim, num_points, lambda_=0.7)
        # Restore document order for narrative flow
        picked_idx.sort(key=lambda i: candidates[i][0])
        return [candidates[i][1] for i in picked_idx]

    except Exception:
        # Frequency-only fallback (still polished)
        freqs = Counter(w for w in tokenize_words(section_text) if w not in STOPWORDS and len(w) > 3)
        scored = []
        for orig_i, p in candidates:
            toks = [w for w in tokenize_words(p) if w in freqs]
            score = sum(freqs[w] for w in toks) / max(1, len(p.split()) ** 0.6)
            score = score * _sentence_quality(p) + _stat_bonus(p)
            scored.append((orig_i, p, score))
        scored.sort(key=lambda x: x[2], reverse=True)
        picked = sorted(scored[:num_points], key=lambda x: x[0])
        return [p for _, p, _ in picked]


def compress_bullet(sentence: str, max_chars: int = 150) -> str:
    """Final tightening for a polished sentence -> bullet."""
    s = polish_sentence(sentence).rstrip(" .;:,")
    if not s:
        return ""
    if len(s) <= max_chars:
        return s
    cut = s[:max_chars]
    for sep in [";", ":", " — ", " - ", ","]:
        idx = cut.rfind(sep)
        if idx > max_chars * 0.55:
            return cut[:idx].rstrip()
    idx = cut.rfind(" ")
    return (cut[:idx].rstrip() if idx > 0 else cut) + "\u2026"


# Back-compat shim (older callers used to_bullet_phrase)
def to_bullet_phrase(sentence: str, max_chars: int = 150) -> str:
    return compress_bullet(sentence, max_chars=max_chars)


def deduplicate_across_sections(sections):
    """Drop bullets in later sections that are near-duplicates of earlier ones."""
    seen_sets = []
    out = []
    for title, bullets in sections:
        kept = []
        for b in bullets:
            toks = {w for w in tokenize_words(b) if w not in STOPWORDS and len(w) > 3}
            if not toks:
                kept.append(b)
                continue
            dup = False
            for prev in seen_sets:
                if not prev:
                    continue
                overlap = len(toks & prev) / max(1, len(toks | prev))
                if overlap >= 0.6:
                    dup = True
                    break
            if not dup:
                kept.append(b)
                seen_sets.append(toks)
        if kept:
            out.append((title, kept))
    return out


# ---------------------------------------------------------------------------
# PPTX building (custom slides on blank layout)
# ---------------------------------------------------------------------------
SLIDE_W_IN = 13.333
SLIDE_H_IN = 7.5


def _set_bg(slide, color: RGBColor):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def _add_rect(slide, x, y, w, h, color: RGBColor):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = color
    shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def _add_oval(slide, x, y, w, h, color: RGBColor):
    shp = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = color
    shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def _add_text(slide, x, y, w, h, text, *, size=18, bold=False, color=None,
              align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, font_name="Calibri"):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = Inches(0)
    tf.margin_right = Inches(0)
    tf.margin_top = Inches(0)
    tf.margin_bottom = Inches(0)
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = font_name
    if color is not None:
        run.font.color.rgb = color
    return tb


def _add_footer(slide, theme, doc_title, slide_num, total):
    _add_rect(slide,
              Inches(0.6), Inches(SLIDE_H_IN - 0.55),
              Inches(SLIDE_W_IN - 1.2), Emu(9525),
              theme["muted"])
    _add_text(slide,
              Inches(0.6), Inches(SLIDE_H_IN - 0.45),
              Inches(8), Inches(0.3),
              doc_title[:80],
              size=10, color=theme["muted"], align=PP_ALIGN.LEFT)
    _add_text(slide,
              Inches(SLIDE_W_IN - 2.2), Inches(SLIDE_H_IN - 0.45),
              Inches(1.6), Inches(0.3),
              f"{slide_num} / {total}",
              size=10, color=theme["muted"], align=PP_ALIGN.RIGHT)


def _add_slide_header(slide, theme, title_text):
    _add_rect(slide, Inches(0.6), Inches(0.55), Inches(0.12), Inches(0.55), theme["accent"])
    tb = slide.shapes.add_textbox(Inches(0.95), Inches(0.45),
                                   Inches(SLIDE_W_IN - 1.5), Inches(0.85))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0)
    tf.margin_top = Inches(0)
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = title_text.upper()
    run.font.size = Pt(30)
    run.font.bold = True
    run.font.name = "Calibri"
    run.font.color.rgb = theme["primary"]
    _add_rect(slide, Inches(0.95), Inches(1.30), Inches(1.4), Inches(0.06), theme["accent"])


def _add_bullets(slide, theme, points):
    tb = slide.shapes.add_textbox(
        Inches(0.95), Inches(1.7),
        Inches(SLIDE_W_IN - 1.9), Inches(SLIDE_H_IN - 2.7),
    )
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0)
    tf.margin_top = Inches(0)

    first = True
    for pt in points:
        if not pt:
            continue
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = PP_ALIGN.LEFT
        p.space_after = Pt(12)
        p.line_spacing = 1.25

        marker = p.add_run()
        marker.text = "\u25CF  "  # ●
        marker.font.size = Pt(14)
        marker.font.name = "Calibri"
        marker.font.color.rgb = theme["accent"]

        body = p.add_run()
        body.text = pt
        body.font.size = Pt(18)
        body.font.name = "Calibri"
        body.font.color.rgb = theme["text"]


# --- Specific slide builders -------------------------------------------------
def slide_title(prs, theme, title, subtitle, meta):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _set_bg(slide, theme["bg"])
    _add_oval(slide, Inches(-1.5), Inches(-1.5), Inches(4.5), Inches(4.5), theme["panel"])
    _add_oval(slide, Inches(SLIDE_W_IN - 2.5), Inches(SLIDE_H_IN - 2.5),
              Inches(4.0), Inches(4.0), theme["panel"])
    _add_rect(slide, Inches(0.9), Inches(2.4), Inches(0.5), Inches(0.08), theme["accent"])
    _add_text(slide, Inches(0.9), Inches(2.0), Inches(8), Inches(0.4),
              "PRESENTATION", size=14, bold=True, color=theme["accent"])
    _add_text(slide, Inches(0.9), Inches(2.6), Inches(SLIDE_W_IN - 2), Inches(2.2),
              title, size=48, bold=True, color=theme["primary"])
    if subtitle:
        _add_text(slide, Inches(0.9), Inches(4.9), Inches(SLIDE_W_IN - 2), Inches(0.8),
                  subtitle, size=20, color=theme["text"])
    meta_bits = []
    if meta.get("author"):
        meta_bits.append(meta["author"])
    if meta.get("page_count"):
        meta_bits.append(f"{meta['page_count']} pages")
    meta_bits.append(datetime.now().strftime("%b %Y"))
    _add_text(slide, Inches(0.9), Inches(6.4), Inches(SLIDE_W_IN - 2), Inches(0.4),
              "  \u00B7  ".join(meta_bits), size=12, color=theme["muted"])
    return slide


def slide_toc(prs, theme, sections, doc_title, slide_num, total):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _set_bg(slide, theme["bg"])
    _add_slide_header(slide, theme, "Contents")

    half = (len(sections) + 1) // 2

    def render_col(items, start_idx, left):
        tb = slide.shapes.add_textbox(left, Inches(1.8),
                                       Inches(SLIDE_W_IN / 2 - 1.2),
                                       Inches(SLIDE_H_IN - 2.7))
        tf = tb.text_frame
        tf.word_wrap = True
        first = True
        for j, name in enumerate(items):
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            p.space_after = Pt(14)
            num = p.add_run()
            num.text = f"{start_idx + j:02d}  "
            num.font.size = Pt(22)
            num.font.bold = True
            num.font.color.rgb = theme["accent"]
            num.font.name = "Calibri"
            body = p.add_run()
            body.text = name
            body.font.size = Pt(18)
            body.font.color.rgb = theme["text"]
            body.font.name = "Calibri"

    render_col(sections[:half], 1, Inches(0.95))
    if len(sections) > half:
        render_col(sections[half:], half + 1, Inches(SLIDE_W_IN / 2 + 0.3))

    _add_footer(slide, theme, doc_title, slide_num, total)
    return slide


def slide_section_divider(prs, theme, number, title, doc_title, slide_num, total):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _set_bg(slide, theme["bg"])
    _add_rect(slide, Inches(0), Inches(0), Inches(4.2), Inches(SLIDE_H_IN), theme["panel"])
    _add_rect(slide, Inches(4.2), Inches(0), Inches(0.08), Inches(SLIDE_H_IN), theme["accent"])
    _add_text(slide, Inches(0.6), Inches(2.0), Inches(3.4), Inches(0.5),
              "SECTION", size=14, bold=True, color=theme["primary"])
    _add_text(slide, Inches(0.6), Inches(2.5), Inches(3.4), Inches(2.0),
              f"{number:02d}", size=120, bold=True, color=theme["accent"], align=PP_ALIGN.LEFT)
    _add_rect(slide, Inches(4.8), Inches(2.8), Inches(0.8), Inches(0.08), theme["accent"])
    _add_text(slide, Inches(4.8), Inches(3.0), Inches(SLIDE_W_IN - 5.5), Inches(2.5),
              title, size=44, bold=True, color=theme["primary"])
    _add_footer(slide, theme, doc_title, slide_num, total)
    return slide


def slide_content(prs, theme, title, bullets, doc_title, slide_num, total):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _set_bg(slide, theme["bg"])
    _add_slide_header(slide, theme, title)
    _add_bullets(slide, theme, bullets)
    _add_footer(slide, theme, doc_title, slide_num, total)
    return slide


def slide_key_terms(prs, theme, terms, doc_title, slide_num, total):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _set_bg(slide, theme["bg"])
    _add_slide_header(slide, theme, "Key Terms")

    chip_h = Inches(0.55)
    chip_pad = Inches(0.18)
    chip_x_start = Inches(0.95)
    max_x = Inches(SLIDE_W_IN - 0.95)

    cur_x = chip_x_start
    cur_y = Inches(2.0)
    for term in terms:
        label = term.upper() if len(term) <= 22 else term[:22].upper() + "\u2026"
        chip_w = Inches(max(1.4, 0.13 * len(label) + 0.6))
        if cur_x + chip_w > max_x:
            cur_x = chip_x_start
            cur_y = cur_y + chip_h + chip_pad
        shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cur_x, cur_y, chip_w, chip_h)
        shp.fill.solid()
        shp.fill.fore_color.rgb = theme["panel"]
        shp.line.color.rgb = theme["accent"]
        shp.line.width = Pt(1)
        tf = shp.text_frame
        tf.margin_left = Inches(0.15)
        tf.margin_right = Inches(0.15)
        tf.margin_top = Inches(0.05)
        tf.margin_bottom = Inches(0.05)
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = label
        run.font.size = Pt(14)
        run.font.bold = True
        run.font.name = "Calibri"
        run.font.color.rgb = theme["primary"]
        cur_x = cur_x + chip_w + chip_pad

    _add_footer(slide, theme, doc_title, slide_num, total)
    return slide


def slide_takeaways(prs, theme, takeaways, doc_title, slide_num, total):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _set_bg(slide, theme["bg"])
    _add_slide_header(slide, theme, "Key Takeaways")
    _add_bullets(slide, theme, takeaways)
    _add_footer(slide, theme, doc_title, slide_num, total)
    return slide


def slide_thank_you(prs, theme, doc_title):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _set_bg(slide, theme["bg"])
    _add_oval(slide, Inches(SLIDE_W_IN - 3.5), Inches(-1.5),
              Inches(5.0), Inches(5.0), theme["panel"])
    _add_oval(slide, Inches(-1.5), Inches(SLIDE_H_IN - 3.5),
              Inches(4.0), Inches(4.0), theme["panel"])
    _add_text(slide, Inches(0.9), Inches(2.8), Inches(SLIDE_W_IN - 2), Inches(1.5),
              "Thank You", size=72, bold=True, color=theme["primary"])
    _add_rect(slide, Inches(0.9), Inches(4.2), Inches(0.8), Inches(0.08), theme["accent"])
    _add_text(slide, Inches(0.9), Inches(4.4), Inches(SLIDE_W_IN - 2), Inches(0.6),
              f"Generated from \u201C{doc_title}\u201D", size=18, color=theme["text"])
    return slide


# --- Top-level builder -------------------------------------------------------
def build_presentation(doc_title, subtitle, sections, key_terms, takeaways, theme, meta):
    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W_IN)
    prs.slide_height = Inches(SLIDE_H_IN)

    total = 1 + 1 + 2 * len(sections) + (1 if key_terms else 0) + (1 if takeaways else 0) + 1

    slide_title(prs, theme, doc_title, subtitle, meta)
    slide_num = 2
    slide_toc(prs, theme, [s for s, _ in sections], doc_title, slide_num, total)

    for i, (sec_title, bullets) in enumerate(sections, start=1):
        slide_num += 1
        slide_section_divider(prs, theme, i, sec_title, doc_title, slide_num, total)
        slide_num += 1
        slide_content(prs, theme, sec_title, bullets, doc_title, slide_num, total)

    if key_terms:
        slide_num += 1
        slide_key_terms(prs, theme, key_terms, doc_title, slide_num, total)

    if takeaways:
        slide_num += 1
        slide_takeaways(prs, theme, takeaways, doc_title, slide_num, total)

    slide_thank_you(prs, theme, doc_title)
    return prs


# ---------------------------------------------------------------------------
# Pipeline helpers
# ---------------------------------------------------------------------------
def build_takeaways(cleaned_text, sections_raw, max_items=5):
    """Pick one strong, distinct insight per section using the polished pipeline."""
    picks = []
    seen_sets = []
    for _, body in sections_raw:
        if not isinstance(body, str):
            continue
        candidates = summarise_section(body, num_points=2)
        for s in candidates:
            bullet = compress_bullet(s, max_chars=170)
            if not bullet or len(bullet) < 25:
                continue
            toks = {w for w in tokenize_words(bullet) if w not in STOPWORDS and len(w) > 3}
            if any(len(toks & prev) / max(1, len(toks | prev)) >= 0.55 for prev in seen_sets):
                continue
            picks.append(bullet)
            seen_sets.append(toks)
            break  # one per section
        if len(picks) >= max_items:
            break
    if not picks:
        picks = [compress_bullet(s, 170) for s in summarise_section(cleaned_text, max_items)]
        picks = [p for p in picks if p]
    return picks


def get_download_link(pptx_bytes, filename):
    b64 = base64.b64encode(pptx_bytes).decode()
    return (
        f'<a class="download-button" download="{filename}" '
        f'href="data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,{b64}">'
        "&#8681;  Download PowerPoint</a>"
    )


# ---------------------------------------------------------------------------
# Streamlit UI
# ---------------------------------------------------------------------------
def load_css():
    try:
        with open("style.css", "r") as f:
            st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)
    except Exception:
        pass


def main():
    load_css()
    st.title("Presify")
    st.caption("Turn dense PDFs into clean, on-brand PowerPoint decks.")

    with st.sidebar:
        st.header("Settings")
        theme_name = st.selectbox("Theme", list(THEMES.keys()), index=0)
        num_sections_hint = st.slider("Target number of sections", 3, 8, 6)
        bullets_per_slide = st.slider("Bullets per slide", 3, 7, 5)
        custom_title = st.text_input("Override title (optional)", "")
        custom_subtitle = st.text_input("Subtitle (optional)",
                                        "An executive summary of the report")
        include_key_terms = st.checkbox("Include 'Key Terms' slide", value=True)
        include_takeaways = st.checkbox("Include 'Key Takeaways' slide", value=True)

    uploaded = st.file_uploader("Upload a PDF", type="pdf")
    if uploaded is None:
        st.info("Upload a PDF to begin. Tip: text-based PDFs (not scanned images) work best.")
        return

    st.success(f"Loaded **{uploaded.name}**  ({uploaded.size / 1024:.1f} KB)")

    if not st.button("Generate Presentation", type="primary"):
        return

    progress = st.progress(0)
    status = st.empty()
    theme = THEMES[theme_name]

    try:
        status.text("1/6  Extracting text from PDF\u2026")
        raw_text, meta = extract_text_from_pdf(uploaded)
        progress.progress(15)

        if not raw_text.strip():
            st.error("Couldn't extract any text. The PDF may be scanned or image-based.")
            return

        status.text("2/6  Cleaning and normalising text\u2026")
        cleaned = clean_text(raw_text)
        progress.progress(30)

        status.text("3/6  Detecting title and sections\u2026")
        doc_title = custom_title.strip() or detect_title(raw_text, meta, uploaded.name)

        sections_raw = segment_by_headings(raw_text)
        if not sections_raw:
            sections_raw = auto_segment_by_topics(cleaned, target_sections=num_sections_hint)
        sections_raw = sections_raw[:num_sections_hint]
        progress.progress(50)

        status.text("4/6  Summarising each section\u2026")
        sections = []
        for sec_title, body in sections_raw:
            sentences = summarise_section(body, num_points=bullets_per_slide)
            bullets = [compress_bullet(s, max_chars=150) for s in sentences if s]
            bullets = [b for b in bullets if b and len(b) >= 25]
            if not bullets:
                # Last-ditch fallback: pick first usable polished sentence
                for s in split_sentences(body):
                    cand = compress_bullet(s, 150)
                    if cand and len(cand) >= 25:
                        bullets = [cand]
                        break
            if bullets:
                sections.append((sec_title, bullets))

        # Drop bullets that repeat across sections
        sections = deduplicate_across_sections(sections)
        progress.progress(75)

        if not sections:
            st.error("Could not produce meaningful sections from this document.")
            return

        status.text("5/6  Extracting key terms and takeaways\u2026")
        key_terms = top_keywords(cleaned, n=12) if include_key_terms else []
        takeaways = build_takeaways(cleaned, sections_raw) if include_takeaways else []
        progress.progress(85)

        status.text("6/6  Designing slides\u2026")
        prs = build_presentation(
            doc_title=doc_title,
            subtitle=custom_subtitle.strip(),
            sections=sections,
            key_terms=key_terms,
            takeaways=takeaways,
            theme=theme,
            meta=meta,
        )
        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        progress.progress(100)
        status.text("Done.")

        out_name = f"{uploaded.name.rsplit('.', 1)[0]}.pptx"
        st.markdown(get_download_link(buf.getvalue(), out_name), unsafe_allow_html=True)
        st.success(f"Generated **{out_name}** with {len(sections)} sections.")

        with st.expander("Preview content"):
            st.markdown(f"### {doc_title}")
            if custom_subtitle:
                st.markdown(f"_{custom_subtitle}_")
            for i, (sec, bullets) in enumerate(sections, start=1):
                st.markdown(f"**{i}. {sec}**")
                for b in bullets:
                    st.markdown(f"- {b}")
            if key_terms:
                st.markdown("**Key Terms:** " + ", ".join(key_terms))
            if takeaways:
                st.markdown("**Takeaways:**")
                for t in takeaways:
                    st.markdown(f"- {t}")
    except Exception as e:
        st.error(f"Something went wrong: {e}")
        st.code(traceback.format_exc())


if __name__ == "__main__":
    main()

