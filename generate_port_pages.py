"""
Port Pages Generator
Generates modern, responsive, SEO-optimized port landing pages matching the new
Ship Chandler Georgia theme (Georgia Blue & Red, Instrument Sans & Karla typography),
centralized CSS (assets/css/port-pages.css), centralized JS (assets/js/port-pages.js),
standardized port-hero.jpg image, navbar & footer matching index.html,
schema markup (JSON-LD), horizontal tabbed TOC, and structured article content from Excel data.
"""

import os
import re
import html
import json
import shutil
import openpyxl

# Configuration Paths
REPO_ROOT = os.path.dirname(os.path.abspath(__file__))
PORTS_DIR = os.path.join(REPO_ROOT, "ports")
CSS_FILE = os.path.join(REPO_ROOT, "assets", "css", "port-pages.css")
JS_FILE = os.path.join(REPO_ROOT, "assets", "js", "port-pages.js")
IMGS_DIR = os.path.join(REPO_ROOT, "assets", "imgs")
EXCEL_PATH = r"D:\FalconXoft Ship Chandler\Website Data\Ship_Chandler_SEO_49_Pages_FULL.xlsx"
HERO_IMG_NAME = "port-hero.jpg"

# 22 Service categories metadata matching Bissau port page layout exactly
CATEGORY_META = [
    ("professional-services", "fas fa-ship"),
    ("provisions", "fas fa-apple-alt"),
    ("bonded", "fas fa-box"),
    ("deck", "fas fa-anchor"),
    ("engine", "fas fa-cogs"),
    ("cabin", "fas fa-bed"),
    ("safety", "fas fa-hard-hat"),
    ("electrical", "fas fa-bolt"),
    ("chemicals", "fas fa-flask"),
    ("welding", "fas fa-fire"),
    ("spare-parts", "fas fa-wrench"),
    ("publications", "fas fa-book"),
    ("tools", "fas fa-tools"),
    ("technical-support", "fas fa-user-cog"),
    ("urgent", "fas fa-clock"),
    ("vessels", "fas fa-ship"),
    ("rfq", "fas fa-file-invoice"),
    ("why-choose", "fas fa-check-circle"),
    ("", "fas fa-shield-alt"),
    ("", "fas fa-handshake"),
    ("", "fas fa-paper-plane"),
    ("", "fas fa-hand-holding-heart"),
]


def ensure_hero_image():
    """
    Ensures assets/imgs/port-hero.jpg exists.
    If not, copies bg-hero.jpg to port-hero.jpg as a reliable base image.
    Also ensures logo.png is copied to root directory if needed.
    """
    target_hero = os.path.join(IMGS_DIR, HERO_IMG_NAME)
    source_hero = os.path.join(IMGS_DIR, "bg-hero.jpg")
    if not os.path.exists(target_hero) and os.path.exists(source_hero):
        shutil.copyfile(source_hero, target_hero)
        print(f"[IMAGE] Created {target_hero} from {source_hero}")

    # Ensure logo in root directory exists for convenience
    root_logo = os.path.join(REPO_ROOT, "logo.png")
    source_logo = os.path.join(IMGS_DIR, "logo.png")
    if not os.path.exists(root_logo) and os.path.exists(source_logo):
        shutil.copyfile(source_logo, root_logo)
        print(f"[IMAGE] Created root {root_logo}")


def get_clean_port_name(page_name):
    """
    Derives clean port name for the CTA heading:
    e.g. 'Bissau Ship Chandler - Guinea' -> 'Bissau Port'
         'Ship Chandler Apapa' -> 'Apapa Port'
         'Ship Chandler Dar Es Salaam' -> 'Dar Es Salaam Port'
         'Ship Chandler Onne Port Nigeria' -> 'Onne Port Nigeria'
    """
    clean = page_name
    clean = re.sub(r'\b(Ship\s*Chandlers?|Ship\s*Supply\s*Chandler|Shipchandler)\b', '', clean, flags=re.IGNORECASE)
    clean = re.sub(r'\s*-\s*.*$', '', clean).strip()
    clean = re.sub(r'\s+', ' ', clean).strip()
    if not re.search(r'\bport\b', clean, flags=re.IGNORECASE):
        clean += ' Port'
    return clean


def format_hero_title(h1, page_name=None):
    """
    Formats the H1 title with an <em> highlight matching Bissau's two-tone theme.
    e.g. 'Bissau Ship Chandler - Guinea' -> 'Bissau Ship Chandler <br><em>- Guinea</em>'
         'Ship Chandler Dar Es Salaam' -> 'Ship Chandler <br><em>Dar Es Salaam</em>'
         'Monrovia Ship Chandler' -> '<em>Monrovia</em> <br>Ship Chandler'
    """
    h1 = str(h1).strip()
    if ' - ' in h1:
        parts = h1.split(' - ', 1)
        return f"{parts[0]} <br><em>- {parts[1]}</em>"
    elif ' -' in h1:
        parts = h1.split(' -', 1)
        return f"{parts[0]} <br><em>-{parts[1]}</em>"

    m = re.match(r'^(Ship\s*Chandlers?|Shipchandler)\s+(.*)$', h1, flags=re.IGNORECASE)
    if m:
        return f"{m.group(1)} <br><em>{m.group(2)}</em>"

    m2 = re.match(r'^(.*?)\s+(Ship\s*Chandlers?|Ship\s*Supply\s*Chandler)$', h1, flags=re.IGNORECASE)
    if m2:
        return f"<em>{m2.group(1)}</em> <br>{m2.group(2)}"

    return f"<em>{h1}</em>"


def render_article_html(desc_text, clean_port):
    """
    Parses the Markdown description from Excel into semantic HTML
    matching the exact structure, category IDs, icons, pill lists, and CTA of the Bissau page.
    """
    blocks = [b.strip() for b in str(desc_text).split('\n\n') if b.strip()]
    lines = []
    cat_idx = 0
    in_category = False

    for b in blocks:
        if b.startswith('# '):
            # Skip top-level H1 (rendered in Hero section)
            continue
        elif b.startswith('## '):
            if in_category:
                lines.append('          </div>')
                lines.append('')

            heading_text = b[3:].strip()
            heading_text = heading_text.replace('\u2013', '–').replace('\u2014', '–')

            cat_id, cat_icon = CATEGORY_META[cat_idx] if cat_idx < len(CATEGORY_META) else ("", "fas fa-ship")
            cat_idx += 1

            if cat_id:
                lines.append(f'          <div class="service-category" id="{cat_id}">')
            else:
                lines.append('          <div class="service-category">')

            lines.append(f'            <h2><i class="{cat_icon}"></i> {heading_text}</h2>')
            in_category = True
        elif b.startswith('- '):
            lines.append('            <ul>')
            for item in b.split('\n'):
                item = item.strip()
                if not item:
                    continue
                val = re.sub(r'^\s*-\s*', '', item)
                val = val.replace('\u2013', '–').replace('\u2014', '–')
                lines.append(f'              <li>{val}</li>')
            lines.append('            </ul>')
        else:
            indent = '            ' if in_category else '          '
            if b.startswith('**') and b.endswith('**'):
                bold_content = b[2:-2].strip().replace('\u2013', '–').replace('\u2014', '–')
                lines.append(f'{indent}<p><strong>{bold_content}</strong></p>')
            else:
                clean_p = ' '.join(b.split()).replace('\u2013', '–').replace('\u2014', '–')
                lines.append(f'{indent}<p>{clean_p}</p>')

    if in_category:
        lines.append('          </div>')
        lines.append('')

    # Article Call to Action
    lines.append('          <!-- Article CTA -->')
    lines.append('          <div class="port-article-cta">')
    lines.append(f'            <h3>Need Supply at {clean_port}?</h3>')
    lines.append('            <p>Send us your complete vessel requisition and our team will respond with a competitive quotation.</p>')
    lines.append('            <a href="../request-quote.html" class="btn btn-fill">Request a Quote <i class="fas fa-arrow-right"></i></a>')
    lines.append('          </div>')

    return '\n'.join(lines)


HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="en">

<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
  <title>__META_TITLE__</title>
  <meta name="description"
    content="__META_DESC__">
  <meta name="keywords"
    content="__KEYWORDS__">
  <meta name="author" content="SHIP CHANDLER GEORGIA CO., LTD">

  <link rel="canonical" href="__CANONICAL_URL__">

  <!-- Open Graph Meta Tags -->
  <meta property="og:title" content="__META_TITLE__">
  <meta property="og:url" content="__CANONICAL_URL__">
  <meta property="og:description"
    content="__META_DESC__">
  <meta property="og:type" content="website">
  <meta property="og:site_name" content="Ship Chandler Georgia">
  <meta property="og:image" content="https://shipchandlergeorgia.com/assets/imgs/port-hero.jpg">

  <!-- Twitter Meta Tags -->
  <meta name="twitter:card" content="summary_large_image">
  <meta name="twitter:title" content="__META_TITLE__">
  <meta name="twitter:description"
    content="__META_DESC__">
  <meta name="twitter:image" content="https://shipchandlergeorgia.com/assets/imgs/port-hero.jpg">

  <!-- Structured Data (JSON-LD) -->
  <script type="application/ld+json">
__SCHEMA_JSON__
  </script>

  <!-- Fonts -->
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link
    href="https://fonts.googleapis.com/css2?family=Karla:ital,wght@0,300;0,400;0,500;0,600;0,700;0,800;1,400&family=Instrument+Sans:wght@400;500;600;700&display=swap"
    rel="stylesheet">

  <!-- Font Awesome -->
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css">

  <link rel="icon" type="image/png" href="../favicon.png">
  <link rel="shortcut icon" href="../favicon.ico">
  <link rel="apple-touch-icon" href="../assets/imgs/logo.png">

  <!-- Centralized Port Pages CSS -->
  <link rel="stylesheet" href="../assets/css/port-pages.css">
</head>

<body>

  <!-- HEADER -->
  <header class="masthead" id="masthead">
    <div class="wrap masthead-inner">
      <a class="brand" href="../index.html">
        <img src="../assets/imgs/logo.png" alt="Ship Chandler Georgia Logo" width="46" height="46">
        <span class="brand-copy">
          <strong>ShipChandler</strong>
          <span>Georgia</span>
        </span>
      </a>

      <button class="burger" type="button" aria-label="Toggle navigation">
        <span></span><span></span><span></span>
      </button>

      <nav class="nav-shell" id="navShell">
        <ul class="nav-set">
          <li class="nav-unit"><a class="nav-anchor" href="../index.html">Home</a></li>
          <li class="nav-unit"><a class="nav-anchor" href="../about.html">About</a></li>

          <!-- Services flyout -->
          <li class="nav-unit has-flyout">
            <a class="nav-anchor" href="../services.html">Services <span class="caret"><svg viewBox="0 0 24 24"
                  width="9" height="9" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round"
                  stroke-linejoin="round">
                  <polyline points="6 9 12 15 18 9"></polyline>
                </svg></span></a>
            <div class="flyout">
              <a href="../services.html#provisions">Food Provisions Supply</a>
              <a href="../services.html#bonded">Bonded Stores</a>
              <a href="../services.html#deck">Deck Stores</a>
              <a href="../services.html#engine">Engine Stores</a>
              <a href="../services.html#cabin">Cabin & Galley</a>
              <a href="../services.html#safety">Safety & PPE</a>
              <a href="../services.html#electrical">Electrical Stores</a>
              <a href="../services.html#chemicals">Marine Chemicals</a>
              <a href="../services.html#welding">Welding Equipment</a>
              <a href="../services.html#spare-parts">Spare Parts</a>
              <a href="../services.html#publications">Nautical Charts</a>
              <a href="../services.html#tools">Tools & Hardware</a>
              <a href="../services.html#urgent">Urgent Supply</a>
            </div>
          </li>

          <!-- Ports mega -->
          <li class="nav-unit has-mega">
            <a class="nav-anchor" href="javascript:void(0)" id="navPorts">Ports <span class="caret"><svg
                  viewBox="0 0 24 24" width="9" height="9" fill="none" stroke="currentColor" stroke-width="2.5"
                  stroke-linecap="round" stroke-linejoin="round">
                  <polyline points="6 9 12 15 18 9"></polyline>
                </svg></span></a>
            <div class="mega" id="megaPortsMenu">
              <div class="mega-top">
                <h4>Global Port Network</h4>
                <div class="mega-filter">
                  <i class="fas fa-search" style="color:var(--steel);font-size:0.8rem;"></i>
                  <input type="text" id="portSearch" placeholder="Search 49 ports...">
                </div>
              </div>
              <div class="mega-cols" id="megaPorts">
                <!-- populated via JS -->
              </div>
            </div>
          </li>

          <li class="nav-unit"><a class="nav-anchor" href="../contact.html">Contact</a></li>
        </ul>
        <a class="btn btn-fill" style="padding:10px 22px;font-size:0.78rem;" href="../request-quote.html">Get Quote</a>
      </nav>
    </div>
  </header>

  <main>
    <!-- PORT HERO -->
    <section class="port-hero">
      <img class="port-hero-bg" src="../assets/imgs/port-hero.jpg" alt="__HERO_ALT__">
      <div class="port-hero-scrim"></div>

      <div class="wrap port-hero-inner">
        <div class="port-hero-copy">
          <div class="breadcrumb">
            <a href="../index.html">Home</a>
            <span>/</span>
            <a href="../index.html#ports">Ports</a>
            <span>/</span>
            <span>__PAGE_NAME__</span>
          </div>
          <span class="kicker" style="color: var(--red);">Port Supply Profile</span>
          <h1>__HERO_TITLE__</h1>
          <p>__META_DESC__</p>
        </div>


      </div>
    </section>

    <!-- PORT ARTICLE SECTION (NEW TABBED LAYOUT) -->
    <section class="port-article-section">

      <!-- Sticky Horizontal TOC -->
      <div class="port-toc-bar-wrap">
        <div class="wrap">
          <div class="port-toc-bar">
            <a href="#professional-services">Overview</a>
            <a href="#provisions">Provisions</a>
            <a href="#bonded">Bonded Stores</a>
            <a href="#deck">Deck Stores</a>
            <a href="#engine">Engine Stores</a>
            <a href="#cabin">Cabin & Galley</a>
            <a href="#safety">Safety & PPE</a>
            <a href="#electrical">Electrical</a>
            <a href="#chemicals">Chemicals</a>
            <a href="#welding">Welding</a>
            <a href="#spare-parts">Spare Parts</a>
            <a href="#publications">Publications</a>
            <a href="#tools">Tools</a>
            <a href="#technical-support">Tech Support</a>
            <a href="#urgent">Urgent</a>
            <a href="#vessels">Vessels</a>
            <a href="#rfq">RFQ</a>
            <a href="#why-choose">Why Us</a>
          </div>
        </div>
      </div>

      <div class="wrap port-article-layout">
        <article class="port-article-content">
__RENDERED_ARTICLE__
        </article>
      </div>
    </section>
  </main>

  <!-- FOOTER -->
  <footer class="foot">
    <div class="wrap">
      <div class="foot-top">
        <div class="foot-brand">
          <a class="brand" href="../index.html">
            <img src="../assets/imgs/logo.png" alt="Logo" width="46" height="46"
              style="border-radius:8px;background:#fff;padding:3px;border:1px solid rgba(255,255,255,0.1);">
            <span class="brand-copy">
              <strong>ShipChandler</strong>
              <span>Georgia</span>
            </span>
          </a>
          <p>SHIP CHANDLER GEORGIA CO., LTD is a licensed marine supplier providing food provisions,
            food stores, bond store supply, and emergency vessel supply services across the
            Port of Poti, Georgia.</p>
        </div>

        <div class="foot-col">
          <h4>Quick Links</h4>
          <ul class="foot-links">
            <li><a href="../index.html">Home</a></li>
            <li><a href="../about.html">About Us</a></li>
            <li><a href="../services.html">Services</a></li>
            <li><a href="../contact.html">Contact</a></li>
            <li><a href="../request-quote.html">Request a Quote</a></li>
          </ul>
        </div>

        <div class="foot-col">
          <h4>Services</h4>
          <ul class="foot-links">
            <li><a href="../services.html#provisions">Provisions Supply</a></li>
            <li><a href="../services.html#bonded">Bonded Stores</a></li>
            <li><a href="../services.html#deck">Deck Stores</a></li>
            <li><a href="../services.html#engine">Engine Stores</a></li>
            <li><a href="../services.html#safety">Safety & PPE</a></li>
          </ul>
        </div>

        <div class="foot-col">
          <h4>Contact Us</h4>
          <ul class="foot-contact">
            <li><i class="fas fa-map-marker-alt"></i> Port of Poti, 52 David Aghmashenebeli Street, 4400 Poti, Georgia
            </li>
            <li><i class="fas fa-envelope"></i> <a href="mailto:sales@shipchandlergeorgia.com"
                style="color:inherit;">sales@shipchandlergeorgia.com</a></li>
            <li><i class="fas fa-user-tie"></i> Giorgi Beridze</li>
            <li>
              <a class="foot-wa" href="https://wa.me/6590190005" target="_blank" rel="noopener noreferrer"
                aria-label="Direct WhatsApp contact (+65) 9019 0005">
                <span class="foot-wa-ico" aria-hidden="true">
                  <svg viewBox="0 0 24 24" fill="currentColor">
                    <path
                      d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.5-.669-.51-.173-.008-.371-.01-.57-.01-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347m-5.421 7.403h-.004a9.87 9.87 0 01-5.031-1.378l-.361-.214-3.741.982.998-3.648-.235-.374a9.86 9.86 0 01-1.51-5.26c.001-5.45 4.436-9.884 9.888-9.884 2.64 0 5.122 1.03 6.988 2.898a9.825 9.825 0 012.893 6.994c-.003 5.45-4.437 9.884-9.885 9.884m8.413-18.297A11.815 11.815 0 0012.05 0C5.495 0 .16 5.335.157 11.892c0 2.096.547 4.142 1.588 5.945L.057 24l6.305-1.654a11.882 11.882 0 005.683 1.448h.005c6.554 0 11.89-5.335 11.893-11.893a11.821 11.821 0 00-3.48-8.413z" />
                  </svg>
                </span>
                <span class="foot-wa-tag">WhatsApp:</span>
                <span class="foot-wa-num">(+65) 9019 0005</span>
                <span class="foot-wa-arr" aria-hidden="true">&rarr;</span>
              </a>
            </li>
          </ul>
        </div>
      </div>

      <div class="foot-bottom">
        <div>&copy; 2026 SHIP CHANDLER GEORGIA CO., LTD. All rights reserved.</div>
      </div>
    </div>
  </footer>

  <!-- FLOATING WHATSAPP -->
  <a href="https://wa.me/6590190005" class="wa-float" target="_blank" rel="noopener noreferrer"
    aria-label="WhatsApp Contact (+65) 9019 0005">
    <i class="fab fa-whatsapp"></i>
  </a>

  <!-- Centralized Port Pages JavaScript -->
  <script src="../assets/js/port-pages.js"></script>
</body>

</html>"""


def escape_meta(val):
    """
    Safely prepares text for meta tags and attributes by escaping double quotes,
    without converting '&' to '&amp;' so title and descriptions remain clean and natural.
    """
    if val is None:
        return ""
    return str(val).strip().replace('"', '&quot;')


def generate_page_html(page_name, h1, meta_title, meta_desc, full_desc, slug):
    """
    Constructs complete HTML for a port page matching the exact Bissau structure,
    centralized CSS, Georgia theme, SEO tags, and schema markup.
    """
    clean_port = get_clean_port_name(page_name)
    hero_title_html = format_hero_title(h1, page_name)
    rendered_article = render_article_html(full_desc, clean_port)

    canonical_url = f"https://shipchandlergeorgia.com/ports/{slug}"
    escaped_desc = escape_meta(meta_desc)
    escaped_title = escape_meta(meta_title)
    escaped_name = escape_meta(page_name)
    escaped_h1 = escape_meta(h1)
    keywords = f"{escaped_name}, ship chandler, ship supply, provisions, marine stores, bonded stores, deck stores, engine stores, vessel supply, marine procurement"

    # Schema.org JSON-LD
    schema_data = {
        "@context": "https://schema.org",
        "@type": "WebPage",
        "name": meta_title,
        "description": meta_desc,
        "url": canonical_url,
        "breadcrumb": {
            "@type": "BreadcrumbList",
            "itemListElement": [
                {
                    "@type": "ListItem",
                    "position": 1,
                    "name": "Home",
                    "item": "https://shipchandlergeorgia.com/"
                },
                {
                    "@type": "ListItem",
                    "position": 2,
                    "name": "Ports",
                    "item": "https://shipchandlergeorgia.com/#ports"
                },
                {
                    "@type": "ListItem",
                    "position": 3,
                    "name": page_name,
                    "item": canonical_url
                }
            ]
        }
    }
    schema_json = json.dumps(schema_data, indent=4)

    page_html = HTML_TEMPLATE
    page_html = page_html.replace("__META_TITLE__", escaped_title)
    page_html = page_html.replace("__META_DESC__", escaped_desc)
    page_html = page_html.replace("__KEYWORDS__", keywords)
    page_html = page_html.replace("__CANONICAL_URL__", canonical_url)
    page_html = page_html.replace("__SCHEMA_JSON__", schema_json)
    page_html = page_html.replace("__HERO_ALT__", escaped_h1)
    page_html = page_html.replace("__PAGE_NAME__", escaped_name)
    page_html = page_html.replace("__HERO_TITLE__", hero_title_html)
    page_html = page_html.replace("__RENDERED_ARTICLE__", rendered_article)

    return page_html


def main():
    print(f"Reading Excel data from: {EXCEL_PATH}")
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
    sheet = wb['SEO Pages']

    ensure_hero_image()

    rows = list(sheet.iter_rows(values_only=True))[1:]  # Skip header row
    print(f"Found {len(rows)} port rows in Excel.")

    generated_count = 0
    generated_files = []

    for row in rows:
        num, page_name, h1, meta_title, meta_desc, full_desc = row[:6]
        if not page_name:
            continue

        slug = re.sub(r'[^a-z0-9]+', '-', str(page_name).lower()).strip('-') + '.html'
        out_path = os.path.join(PORTS_DIR, slug)

        html_content = generate_page_html(
            page_name=page_name,
            h1=h1,
            meta_title=meta_title,
            meta_desc=meta_desc,
            full_desc=full_desc,
            slug=slug
        )

        with open(out_path, 'w', encoding='utf-8') as f:
            f.write(html_content)

        generated_count += 1
        generated_files.append((num, page_name, slug))
        print(f"[{num:02d}/49] Generated: {slug} ({page_name})")

    unique_files = len(set(f[2] for f in generated_files))
    print("\n" + "=" * 60)
    print(f"SUCCESS: Processed {generated_count} rows across {unique_files} unique port page files.")
    print(f"Centralized CSS written to: {CSS_FILE}")
    print(f"Centralized JS written to:  {JS_FILE}")
    print(f"Standardized hero image:    assets/imgs/{HERO_IMG_NAME}")
    print("=" * 60)


if __name__ == "__main__":
    main()
