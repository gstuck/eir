#!/usr/bin/env python3
"""
Convos with Colleagues — page generator.

Reads `convos.xlsx` (sheet "Convos", header on row 2) and regenerates every
convo detail page (e.g. right-answers.html) plus index.html.

Run locally:
    pip install pandas openpyxl
    python generate_pages.py

The GitHub Action (.github/workflows/build.yml) runs this automatically
whenever `convos.xlsx` is updated in the repo.

What this script DOES touch:
    - One HTML file per convo, named from a slug derived from "Short Title"
    - index.html (the homepage, since theme groupings come from the spreadsheet)

What this script DOES NOT touch:
    - about.html, anatomy.html, tips.html (hand-edited)
    - styles.css
    - PDFs in /pdfs/

If you rename a "Short Title" in the spreadsheet, the slug changes and a
new HTML file is created. The old one stays behind until you delete it
from the repo manually.
"""

import os
import re
import sys
from urllib.parse import quote

try:
    import pandas as pd
except ImportError:
    print("ERROR: pandas not installed. Run: pip install pandas openpyxl")
    sys.exit(1)


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

SPREADSHEET_PATH = os.environ.get("CWC_SPREADSHEET", "convos.xlsx")
SHEET_NAME = "Convos"
HEADER_ROW = 1  # 0-indexed; the actual header is on spreadsheet row 2
OUTPUT_DIR = os.environ.get("CWC_OUTPUT_DIR", ".")

# Order of themes on the homepage. Any theme found in the spreadsheet but
# not listed here is appended to the bottom in alphabetical order.
THEME_ORDER = [
    "Deepening Mathematical Understanding",
    "Fostering Student Discourse",
    "Focusing on the Real World",
]

# Locked slugs for existing convos, so URLs and PDF filenames stay stable
# even if you tweak the "Short Title" in the spreadsheet. Keyed by Convo #.
# If you add a 12th convo (or beyond), it auto-slugs from its Short Title.
# To rename an existing convo's URL: update the slug here AND rename its
# PDFs in /pdfs/ to match.
LOCKED_SLUGS = {
    1: "sequencing-student-work",
    2: "right-answers",
    3: "wrong-answers",
    4: "get-started",
    5: "physical-classroom",
    6: "student-to-student",
    7: "complex-problems",
    8: "voices-and-perspectives",
    9: "rabbit-holes",
    10: "sensitive-topics",
    11: "open-ended-questions",
}

# Map theme → list of (lesson name, lesson URL) tuples used as the Spark
# anchor for that theme. Update if you add new themes.
THEME_LESSONS = {
    "Deepening Mathematical Understanding": [
        ("Hair Today, Gone Tomorrow",
         "https://www.citizenmath.com/lessons/_template.html?slug=hair-today-gone-tomorrow"),
    ],
    "Fostering Student Discourse": [
        ("Big Foot Conspiracy",
         "https://www.citizenmath.com/lessons/_template.html?slug=big-foot-conspiracy"),
    ],
    "Focusing on the Real World": [
        ("Big Foot Conspiracy",
         "https://www.citizenmath.com/lessons/_template.html?slug=big-foot-conspiracy"),
        ("Seeking Shelter",
         "https://www.citizenmath.com/lessons/_template.html?slug=seeking-shelter"),
        ("Coupon Clipping",
         "https://www.citizenmath.com/lessons/_template.html?slug=coupon-clipping"),
    ],
}


# ---------------------------------------------------------------------------
# Reusable building blocks
# ---------------------------------------------------------------------------

SITE_HEADER = '''  <header class="site-header">
    <div class="site-header-inner">
      <a href="index.html" class="site-logo">
        <svg class="site-logo-mark" viewBox="0 0 36 36" width="22" height="22" aria-hidden="true">
          <circle cx="18" cy="18" r="4.5" fill="none" stroke="currentColor" stroke-width="1"/>
          <circle cx="18" cy="6" r="2.2" fill="currentColor"/>
          <circle cx="28.4" cy="12" r="2.2" fill="currentColor"/>
          <circle cx="28.4" cy="24" r="2.2" fill="currentColor"/>
          <circle cx="18" cy="30" r="2.2" fill="currentColor"/>
          <circle cx="7.6" cy="24" r="2.2" fill="currentColor"/>
          <circle cx="7.6" cy="12" r="2.2" fill="currentColor"/>
        </svg>
        <span>Convos with Colleagues</span>
      </a>
      <nav class="site-nav">
        <a href="about.html">About</a>
        <a href="anatomy.html">Anatomy of a Convo</a>
        <a href="tips.html">Tips for a Great Convo</a>
      </nav>
    </div>
  </header>'''


SITE_FOOTER = '''  <footer class="site-footer">
    <p>A professional learning resource. Use freely.</p>
  </footer>'''


SPARK_ICON_SVG = ('<svg viewBox="0 0 24 24" width="11" height="11" fill="none" '
                  'stroke="currentColor" stroke-width="1.5" stroke-linecap="round" '
                  'stroke-linejoin="round" aria-hidden="true">'
                  '<polyline points="7 3 4 13 11 13 8 21 20 9 13 9 16 3 7 3"/></svg>')


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def slugify(text):
    """Convert a string into a URL-friendly slug.
    'Right Answers' -> 'right-answers'
    'Open-Ended Questions' -> 'open-ended-questions'
    """
    if not text:
        return ""
    text = text.lower().strip()
    text = re.sub(r"[''""]", "", text)  # drop fancy quotes
    text = re.sub(r"[^a-z0-9]+", "-", text)  # non-alphanumeric -> hyphen
    text = re.sub(r"^-+|-+$", "", text)  # trim hyphens
    return text


def extract_vimeo_id(url):
    if not url:
        return None
    m = re.search(r"vimeo\.com/(\d+)", str(url))
    return m.group(1) if m else None


def safe(val):
    """Coerce pandas value to a clean string, or '' if NaN/empty."""
    if val is None:
        return ""
    if isinstance(val, float):
        # pandas float NaN
        import math
        if math.isnan(val):
            return ""
    return str(val).strip()


def render_share_url(title):
    subject = "Let's have a Conversation with a Colleague"
    body = (
        f"I came across this discussion guide and thought we could work through it together.\n\n"
        f"{title}\n\n"
        f"[paste page URL here]\n\n"
        f"It's designed for 2–5 people, 30–60 minutes. Want to pick a time?"
    )
    return f"mailto:?subject={quote(subject)}&body={quote(body)}"


# ---------------------------------------------------------------------------
# Spreadsheet loading
# ---------------------------------------------------------------------------

def load_convos():
    """Read the spreadsheet and return a list of cleaned convo dicts."""
    if not os.path.exists(SPREADSHEET_PATH):
        print(f"ERROR: Spreadsheet not found at {SPREADSHEET_PATH}")
        sys.exit(1)

    df = pd.read_excel(SPREADSHEET_PATH, sheet_name=SHEET_NAME, header=HEADER_ROW)

    # Filter to rows with a numeric Convo #
    df = df[pd.to_numeric(df["Convo #"], errors="coerce").notna()]

    # Spark Type column has a newline in the header. Map to a clean key.
    spark_type_col = next((c for c in df.columns if c.lower().startswith("spark type")), None)
    if not spark_type_col:
        print("ERROR: Couldn't find 'Spark Type' column in spreadsheet.")
        sys.exit(1)

    convos = []
    for _, row in df.iterrows():
        num = int(row["Convo #"])
        short = safe(row["Short Title"])
        # Use locked slug if this is one of the original convos; otherwise derive
        slug = LOCKED_SLUGS.get(num) or slugify(short)
        if not slug:
            print(f"ERROR: Convo #{num} has no Short Title and is not in LOCKED_SLUGS.")
            sys.exit(1)
        convos.append({
            "num": num,
            "slug": slug,
            "short_title": short,
            "title": safe(row["Title / Question"]),
            "overview": safe(row["Overview / Description"]),
            "theme": safe(row["Theme"]),
            "focus_lesson_name": safe(row["Focus Lesson Name"]),
            "focus_lesson_url": safe(row["Focus Lesson URL"]),
            "spark_type": safe(row[spark_type_col]),
            "spark_label": safe(row["Spark Label"]),
            "spark_caption": safe(row["Spark Caption"]),
            "spark_url": safe(row["Spark URL"]),
            "observe_q1": safe(row.get("Observe Q1")),
            "observe_q2": safe(row.get("Observe Q2")),
            "discuss_q1": safe(row.get("Discuss Q1")),
            "discuss_q2": safe(row.get("Discuss Q2")),
            "discuss_q3": safe(row.get("Discuss Q3")),
            "relate_q1": safe(row.get("Relate Q1")),
            "relate_q2": safe(row.get("Relate Q2")),
            "relate_q3": safe(row.get("Relate Q3")),
            "commit_q1": safe(row.get("Commit Q1")),
        })

    # Sanity-check slugs are unique
    slugs = [c["slug"] for c in convos]
    if len(slugs) != len(set(slugs)):
        from collections import Counter
        dupes = [s for s, n in Counter(slugs).items() if n > 1]
        print(f"ERROR: Duplicate slugs from Short Titles: {dupes}")
        sys.exit(1)

    return sorted(convos, key=lambda c: c["num"])


# ---------------------------------------------------------------------------
# Detail page rendering
# ---------------------------------------------------------------------------

def render_questions_ul(*qs):
    items = [q for q in qs if q]
    if not items:
        return ""
    return "\n          ".join(f"<li>{q}</li>" for q in items)


def render_section(section_id, h2, time_label, body_inner):
    return f'''
      <section class="section" id="{section_id}">
        <button class="section-header" type="button" aria-expanded="true" aria-controls="{section_id}-body">
          <h2>{h2}</h2>
          <span class="section-time">{time_label}</span>
          <span class="section-toggle" aria-hidden="true">
            <svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="6 9 12 15 18 9"/></svg>
          </span>
        </button>
        <div class="section-body" id="{section_id}-body">
{body_inner}
        </div>
      </section>'''


def render_video_spark(vimeo_id, label, caption):
    return f'''<aside class="spark-inset spark-inset-video" role="complementary">
            <div class="spark-inset-thumb">
              <img src="https://vumbnail.com/{vimeo_id}.jpg" alt="" onerror="this.style.display='none'">
              <button class="spark-play-button" type="button" data-video-id="{vimeo_id}" aria-label="Play animated video">
                <svg viewBox="0 0 24 24" width="24" height="24" fill="currentColor" aria-hidden="true"><polygon points="6 4 20 12 6 20 6 4"/></svg>
              </button>
            </div>
            <div class="spark-inset-body">
              <span class="spark-inset-label">Spark Artifact · {label}</span>
              <p class="spark-inset-caption">{caption}</p>
              <button class="spark-inset-cta" type="button" data-video-id="{vimeo_id}">
                <svg viewBox="0 0 24 24" width="13" height="13" fill="currentColor" aria-hidden="true"><polygon points="6 4 20 12 6 20 6 4"/></svg>
                Watch the video
              </button>
            </div>
          </aside>'''


def render_pdf_spark(slug, label, caption):
    pdf_href = f"pdfs/sparks/{slug}-spark.pdf"
    return f'''<aside class="spark-inset spark-inset-pdf" role="complementary">
            <a class="spark-inset-thumb spark-inset-thumb-pdf" href="{pdf_href}" target="_blank" rel="noopener" aria-label="Open the spark PDF in a new tab">
              <span class="fanned-page fanned-page-back" aria-hidden="true">
                <span class="fanned-lines"><span></span><span></span><span></span><span></span><span></span></span>
              </span>
              <span class="fanned-page fanned-page-front" aria-hidden="true">
                <span class="fanned-lines"><span></span><span></span><span></span><span></span><span></span><span></span></span>
              </span>
            </a>
            <div class="spark-inset-body">
              <span class="spark-inset-label">Spark Artifact · {label}</span>
              <p class="spark-inset-caption">{caption}</p>
              <a class="spark-inset-cta" href="{pdf_href}" target="_blank" rel="noopener">
                <svg viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M14 3v4a1 1 0 0 0 1 1h4"/><path d="M17 21H7a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h7l5 5v11a2 2 0 0 1-2 2z"/></svg>
                Open the PDF
              </a>
            </div>
          </aside>'''


def render_video_modal():
    return '''
  <div class="video-modal" id="video-modal" role="dialog" aria-modal="true" aria-label="Video player" hidden>
    <div class="video-modal-backdrop" data-close-modal></div>
    <div class="video-modal-content">
      <button class="video-modal-close" type="button" data-close-modal aria-label="Close video">
        <svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
      </button>
      <div class="video-modal-frame">
        <iframe id="video-modal-iframe" src="" allow="autoplay; fullscreen; picture-in-picture; clipboard-write; encrypted-media; web-share" referrerpolicy="strict-origin-when-cross-origin" allowfullscreen></iframe>
      </div>
    </div>
  </div>'''


def render_scripts(has_video):
    video_block = ""
    if has_video:
        video_block = '''
      var modal = document.getElementById('video-modal');
      var iframe = document.getElementById('video-modal-iframe');
      function openVideo(id) {
        iframe.src = 'https://player.vimeo.com/video/' + id + '?autoplay=1&title=0&byline=0&portrait=0&badge=0&autopause=0&player_id=0&app_id=58479';
        modal.hidden = false;
        document.body.style.overflow = 'hidden';
      }
      function closeVideo() {
        iframe.src = '';
        modal.hidden = true;
        document.body.style.overflow = '';
      }
      document.querySelectorAll('[data-video-id]').forEach(function(el) {
        el.addEventListener('click', function() { openVideo(el.getAttribute('data-video-id')); });
      });
      document.querySelectorAll('[data-close-modal]').forEach(function(el) {
        el.addEventListener('click', closeVideo);
      });
      document.addEventListener('keydown', function(e) {
        if (e.key === 'Escape' && !modal.hidden) closeVideo();
      });'''
    return f'''
  <script>
    (function() {{{video_block}
      document.querySelectorAll('.section-header').forEach(function(header) {{
        header.addEventListener('click', function() {{
          var section = header.parentElement;
          var expanded = header.getAttribute('aria-expanded') === 'true';
          header.setAttribute('aria-expanded', String(!expanded));
          section.classList.toggle('section-collapsed', expanded);
        }});
      }});
    }})();
  </script>'''


def render_detail_page(convo):
    vimeo_id = extract_vimeo_id(convo["spark_url"])
    is_video = vimeo_id is not None
    slug = convo["slug"]

    if is_video:
        spark_html = render_video_spark(vimeo_id, convo["spark_label"], convo["spark_caption"])
    else:
        spark_html = render_pdf_spark(slug, convo["spark_label"], convo["spark_caption"])

    observe_body = f'''          {spark_html}
          <ul class="questions">
          {render_questions_ul(convo['observe_q1'], convo['observe_q2'])}
          </ul>'''

    discuss_body = f'''          <ul class="questions">
          {render_questions_ul(convo['discuss_q1'], convo['discuss_q2'], convo['discuss_q3'])}
          </ul>'''

    relate_body = f'''          <ul class="questions">
          {render_questions_ul(convo['relate_q1'], convo['relate_q2'], convo['relate_q3'])}
          </ul>'''

    commit_body = f'''          <ul class="questions">
          {render_questions_ul(convo['commit_q1'])}
          </ul>'''

    sections = (
        render_section("observe", "Observe", "5–15 min", observe_body)
        + render_section("discuss", "Discuss", "10–15 min", discuss_body)
        + render_section("relate", "Relate", "10–15 min", relate_body)
        + render_section("commit", "Commit", "5–10 min", commit_body)
    )

    download_url = f"pdfs/{slug}.pdf"

    return f'''<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>{convo['title']} · Convos with Colleagues</title>
  <meta name="description" content="A roundtable discussion guide for math teachers.">
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Lora:ital,wght@0,400;0,500;1,400&family=Inter:wght@400;500&display=swap" rel="stylesheet">
  <link rel="stylesheet" href="styles.css">
</head>
<body>

{SITE_HEADER}

  <main>

    <div class="hero-meta">
      <span><span class="num">30–60</span>min</span>
      <span><span class="num">2–5</span>educators</span>
      <span>no facilitator needed</span>
    </div>

    <section class="detail-hero">
      <div class="hero-theme">{convo['theme']}</div>
      <h1>{convo['title']}</h1>
      <p class="overview">{convo['overview']}</p>
    </section>

    <div class="container">
      <section class="prep-zone" aria-labelledby="get-ready">
        <div class="prep-actions">
          <a class="prep-action-btn" href="{render_share_url(convo['title'])}">
            <svg viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z"/><polyline points="22,6 12,13 2,6"/></svg>
            Invite colleagues
          </a>
          <a class="prep-action-btn" href="{download_url}" download>
            <svg viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
            Print Convo Guide
          </a>
          <a class="prep-action-btn" href="{convo['focus_lesson_url']}" target="_blank" rel="noopener">
            <svg viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M2 12s3-7 10-7 10 7 10 7-3 7-10 7-10-7-10-7z"/><circle cx="12" cy="12" r="3"/></svg>
            Preview the focus lesson
          </a>
        </div>
      </section>
    </div>

    <div class="container">
{sections}

    </div>

  </main>

{SITE_FOOTER}
{render_video_modal() if is_video else ''}
{render_scripts(is_video)}

</body>
</html>
'''


# ---------------------------------------------------------------------------
# Index (homepage) rendering
# ---------------------------------------------------------------------------

def render_theme_lesson_block(theme):
    lessons = THEME_LESSONS.get(theme, [])
    if not lessons:
        return ""

    label = "Real-World Math lesson used for Spark:" if len(lessons) == 1 \
        else "Real-World Math lessons used for Spark:"

    eye_icon = ('<svg viewBox="0 0 24 24" width="14" height="14" fill="none" '
                'stroke="currentColor" stroke-width="1.5" stroke-linecap="round" '
                'stroke-linejoin="round">'
                '<path d="M2 12s3-7 10-7 10 7 10 7-3 7-10 7-10-7-10-7z"/>'
                '<circle cx="12" cy="12" r="3"/></svg>')

    # Build the lesson links inline
    link_html_pieces = []
    for i, (name, url) in enumerate(lessons):
        link = f'<a class="theme-lesson-link" href="{url}" target="_blank" rel="noopener"><em>{name}</em></a>'
        if i < len(lessons) - 1:
            sep_text = "," if i < len(lessons) - 2 else ", &amp;"
            link_html_pieces.append(f'{link}<span class="theme-lesson-sep">{sep_text}</span>')
        else:
            link_html_pieces.append(link)

    links_html = "\n            ".join(link_html_pieces)

    return f'''<div class="theme-lesson">
            <span class="theme-lesson-icon" aria-hidden="true">
              {eye_icon}
            </span>
            <span class="theme-lesson-label">{label}</span>
            {links_html}
            <span class="theme-lesson-credit">by Citizen Math</span>
          </div>'''


def render_card(convo):
    return f'''
          <a class="convo-card" href="{convo['slug']}.html">
            <span class="convo-short">{convo['short_title']}</span>
            <span class="convo-question">{convo['title']}</span>
            <span class="convo-spark">
              {SPARK_ICON_SVG}
              {convo['spark_label']}
            </span>
          </a>'''


def render_index(convos):
    # Group by theme, preserving spreadsheet order within each theme
    by_theme = {}
    for c in convos:
        by_theme.setdefault(c["theme"], []).append(c)

    # Order: themes in THEME_ORDER first, then any extras alphabetically
    ordered_themes = [t for t in THEME_ORDER if t in by_theme]
    extras = sorted(t for t in by_theme if t not in THEME_ORDER)
    ordered_themes.extend(extras)

    theme_sections = []
    for theme in ordered_themes:
        cards_html = "".join(render_card(c) for c in by_theme[theme])
        lesson_block = render_theme_lesson_block(theme)
        theme_sections.append(f'''
      <section class="theme-section">
        <header class="theme-header">
          <h2>{theme}</h2>
          {lesson_block}
        </header>
        <div class="convo-grid">
{cards_html}
        </div>
      </section>''')

    themes_html = "\n".join(theme_sections)

    # Stats row uses dynamic count of topics
    n_topics = len(convos)

    return f'''<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Convos with Colleagues</title>
  <meta name="description" content="Roundtable discussions for math teachers serious about their craft.">
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Lora:ital,wght@0,400;0,500;1,400&family=Inter:wght@400;500&display=swap" rel="stylesheet">
  <link rel="stylesheet" href="styles.css">
</head>
<body>

{SITE_HEADER}

  <main>
    <section class="home-hero">
      <div class="home-hero-grid">
        <div class="home-hero-content">
          <h1>Convos<br>with Colleagues</h1>
          <p class="dek">Roundtable discussions about common instructional challenges. For math teachers who are serious about their craft.</p>
        </div>
        <aside class="hero-callout">
          <div class="hero-callout-title">What's in each convo</div>
          <ul class="hero-callout-list">
            <li>
              <span class="hero-callout-icon" aria-hidden="true">
                <svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M21 11.5a8.38 8.38 0 0 1-.9 3.8 8.5 8.5 0 0 1-7.6 4.7 8.38 8.38 0 0 1-3.8-.9L3 21l1.9-5.7a8.38 8.38 0 0 1-.9-3.8 8.5 8.5 0 0 1 4.7-7.6 8.38 8.38 0 0 1 3.8-.9h.5a8.48 8.48 0 0 1 8 8v.5z"/></svg>
              </span>
              <div>
                <strong>Convo Guide</strong>
                <span>Carefully crafted questions to guide a productive discussion without a facilitator.</span>
              </div>
            </li>
            <li>
              <span class="hero-callout-icon" aria-hidden="true">
                <svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="7 3 4 13 11 13 8 21 20 9 13 9 16 3 7 3"/></svg>
              </span>
              <div>
                <strong>Spark</strong>
                <span>A short video, classroom transcript, or piece of student work to ground the discussion.</span>
              </div>
            </li>
            <li>
              <span class="hero-callout-icon" aria-hidden="true">
                <svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M2 12s3-7 10-7 10 7 10 7-3 7-10 7-10-7-10-7z"/><circle cx="12" cy="12" r="3"/></svg>
              </span>
              <div>
                <strong>Lesson Context</strong>
                <span>Each spark is contextualized in a high-quality lesson that is mathematically rigorous, conversational, and focused on the real world.</span>
              </div>
            </li>
          </ul>
        </aside>
      </div>

      <div class="hero-stats">
        <div class="hero-stat">
          <div class="hero-stat-num">{n_topics}</div>
          <div class="hero-stat-label">Topics</div>
        </div>
        <div class="hero-stat">
          <div class="hero-stat-num">30–60</div>
          <div class="hero-stat-label">Min each</div>
        </div>
        <div class="hero-stat">
          <div class="hero-stat-num">2–5</div>
          <div class="hero-stat-label">Participants</div>
        </div>
        <div class="hero-stat">
          <div class="hero-stat-num">0</div>
          <div class="hero-stat-label">Prep needed</div>
        </div>
      </div>
    </section>

    <div class="container-wide">
{themes_html}
    </div>

  </main>

{SITE_FOOTER}

</body>
</html>
'''


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    convos = load_convos()
    print(f"Loaded {len(convos)} convos from {SPREADSHEET_PATH}")

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # Generate each detail page
    for convo in convos:
        path = os.path.join(OUTPUT_DIR, f"{convo['slug']}.html")
        with open(path, "w") as f:
            f.write(render_detail_page(convo))
        print(f"  ✓ {convo['slug']}.html ({convo['spark_type'] or 'no spark type'})")

    # Generate index
    index_path = os.path.join(OUTPUT_DIR, "index.html")
    with open(index_path, "w") as f:
        f.write(render_index(convos))
    print(f"  ✓ index.html ({len(convos)} convos across "
          f"{len({c['theme'] for c in convos})} themes)")

    print(f"\nDone. {len(convos) + 1} files written to {OUTPUT_DIR}/")


if __name__ == "__main__":
    main()
