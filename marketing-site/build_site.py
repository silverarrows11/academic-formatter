"""Builds the FormatForge guides website from approved Markdown pages. Standard library only.

    python build_site.py            # content/*.md  ->  dist/

Every page the agents write (and you approve by merging) becomes a static HTML page. Nothing here calls an AI,
tracks visitors, or runs scripts in the visitor's browser.
"""
import datetime
import html
import json
import os
import re
import sys
from pathlib import Path
from urllib.parse import urlparse

ROOT = Path(__file__).resolve().parent
GROUPS = [  # (engine key, heading, one-line blurb)
    ("academic", "Academic papers", "Citations, reference lists and Word formatting for essays and research papers."),
    ("cleaner", "AI drafts", "Cleaning up text pasted from an AI chatbot before it goes into a document."),
    ("resume", "Resumes", "Laying out and tightening a one-page resume."),
]


# --------------------------------------------------------------------------- front matter
def parse_front_matter(text):
    m = re.match(r"\A---[ \t]*\n(.*?)\n---[ \t]*\n(.*)\Z", text.strip() + "\n", re.S)
    if not m:
        raise ValueError("missing front matter")
    meta = {}
    for line in m.group(1).splitlines():
        if ":" in line:
            k, v = line.split(":", 1)
            meta[k.strip()] = v.strip().strip("\"'")
    return meta, m.group(2)


# --------------------------------------------------------------------------- Markdown -> safe HTML
_SAFE_URL = re.compile(r"\A(https?://|mailto:|/|#)", re.I)


def _link(label, url, allowed_hosts):
    url = html.unescape(url).strip()
    if not _SAFE_URL.match(url):
        return html.escape(label)
    host = urlparse(url).hostname if url.lower().startswith("http") else None
    rel = "" if (host is None or host in allowed_hosts) else ' rel="nofollow noopener noreferrer" target="_blank"'
    return f'<a href="{html.escape(url, quote=True)}"{rel}>{label}</a>'


def inline(text, allowed_hosts=()):
    codes = []

    def stash(m):
        codes.append(f"<code>{html.escape(m.group(1))}</code>")
        return f"\x00{len(codes) - 1}\x00"

    t = re.sub(r"`([^`\n]+)`", stash, text)
    t = html.escape(t, quote=False)
    t = re.sub(r"\[([^\]\n]+)\]\(([^)\s]+)\)", lambda m: _link(m.group(1), m.group(2), allowed_hosts), t)
    t = re.sub(r"\*\*([^*\n]+)\*\*", r"<strong>\1</strong>", t)
    t = re.sub(r"(?<![\w*])\*([^*\n]+)\*(?![\w*])", r"<em>\1</em>", t)
    t = re.sub(r"(?<![\w_])_([^_\n]+)_(?![\w_])", r"<em>\1</em>", t)
    # bare URLs become links too
    t = re.sub(r'(?<!["=>])(https?://[^\s<)]+[^\s<).,;:])',
               lambda m: _link(m.group(1), m.group(1), allowed_hosts), t)
    return re.sub(r"\x00(\d+)\x00", lambda m: codes[int(m.group(1))], t)


def render_markdown(md, allowed_hosts=()):
    lines = md.replace("\r\n", "\n").split("\n")
    out, i, n = [], 0, len(lines)
    inl = lambda s: inline(s, allowed_hosts)  # noqa: E731
    while i < n:
        line = lines[i]
        if not line.strip():
            i += 1
            continue
        if line.startswith("```"):
            i += 1
            buf = []
            while i < n and not lines[i].startswith("```"):
                buf.append(lines[i])
                i += 1
            i += 1
            out.append(f"<pre><code>{html.escape(chr(10).join(buf))}</code></pre>")
            continue
        m = re.match(r"(#{1,6})\s+(.*)", line)
        if m:
            level = max(2, len(m.group(1)))      # the page title is the only h1
            text = m.group(2).strip()
            slug = re.sub(r"[^a-z0-9]+", "-", text.lower()).strip("-")
            out.append(f'<h{level} id="{html.escape(slug)}">{inl(text)}</h{level}>')
            i += 1
            continue
        if re.fullmatch(r"\s*([-*_])\1{2,}\s*", line):
            out.append("<hr>")
            i += 1
            continue
        if line.startswith(">"):
            buf = []
            while i < n and lines[i].startswith(">"):
                buf.append(lines[i][1:].lstrip())
                i += 1
            out.append(f"<blockquote>{render_markdown(chr(10).join(buf), allowed_hosts)}</blockquote>")
            continue
        if re.match(r"\s*[-*+]\s+", line) or re.match(r"\s*\d+[.)]\s+", line):
            ordered = bool(re.match(r"\s*\d+[.)]\s+", line))
            pat = r"\s*\d+[.)]\s+" if ordered else r"\s*[-*+]\s+"
            items = []
            while i < n and re.match(pat, lines[i]):
                items.append(inl(re.sub(pat, "", lines[i], count=1).strip()))
                i += 1
            tag = "ol" if ordered else "ul"
            out.append(f"<{tag}>" + "".join(f"<li>{x}</li>" for x in items) + f"</{tag}>")
            continue
        if "|" in line and i + 1 < n and re.match(r"\s*\|?\s*:?-{2,}", lines[i + 1]):
            def cells(s):
                return [c.strip() for c in s.strip().strip("|").split("|")]
            head = cells(line)
            i += 2
            rows = []
            while i < n and "|" in lines[i] and lines[i].strip():
                rows.append(cells(lines[i]))
                i += 1
            out.append("<div class='table-wrap'><table><thead><tr>" + "".join(f"<th>{inl(c)}</th>" for c in head)
                       + "</tr></thead><tbody>" + "".join("<tr>" + "".join(f"<td>{inl(c)}</td>" for c in r) + "</tr>" for r in rows)
                       + "</tbody></table></div>")
            continue
        buf = []
        while i < n and lines[i].strip() and not re.match(r"(#{1,6}\s|```|>|\s*[-*+]\s+|\s*\d+[.)]\s+)", lines[i]):
            buf.append(lines[i].strip())
            i += 1
        out.append(f"<p>{inl(' '.join(buf))}</p>")
    return "\n".join(out)


# --------------------------------------------------------------------------- page templates
CSS = """
:root{color-scheme:dark}
*{box-sizing:border-box}
html{-webkit-text-size-adjust:100%}
body{margin:0;background:#0B1220;color:#E2E8F0;font:17px/1.7 -apple-system,BlinkMacSystemFont,"Segoe UI",Inter,Roboto,Helvetica,Arial,sans-serif}
a{color:#34D399;text-underline-offset:3px}a:hover{color:#6EE7B7}
.wrap{max-width:760px;margin:0 auto;padding:0 20px}
.wide{max-width:1000px}
header.site{border-bottom:1px solid #1E293B;background:#0B1220;position:sticky;top:0;z-index:5}
header.site .wrap{display:flex;align-items:center;justify-content:space-between;gap:12px;height:64px}
.brand{display:flex;align-items:center;gap:10px;color:#fff;text-decoration:none;font-weight:800;font-size:18px}
.logo{width:30px;height:30px;border-radius:8px;background:#10B981;color:#052E16;display:grid;place-items:center;font-weight:800}
.brand small{color:#94A3B8;font-weight:600;font-size:14px}
.btn{display:inline-block;background:#10B981;color:#052E16;font-weight:700;text-decoration:none;padding:10px 18px;border-radius:10px;font-size:15px;line-height:1.2}
.btn:hover{background:#34D399;color:#052E16}
.btn.ghost{background:transparent;color:#E2E8F0;border:1px solid #334155}
main{padding:40px 0 56px}
h1{font-size:clamp(28px,5vw,40px);line-height:1.15;margin:0 0 12px;color:#fff;letter-spacing:-.02em}
h2{font-size:24px;line-height:1.3;margin:40px 0 10px;color:#fff}
h3{font-size:19px;margin:28px 0 8px;color:#fff}
p{margin:0 0 16px}ul,ol{padding-left:24px;margin:0 0 18px}li{margin-bottom:6px}
.meta{color:#94A3B8;font-size:14px;margin-bottom:28px}
.lede{color:#CBD5E1;font-size:19px;margin-bottom:8px}
blockquote{margin:18px 0;padding:14px 18px;background:#0F172A;border:1px solid #334155;border-left:4px solid #10B981;border-radius:10px;color:#CBD5E1}
blockquote p:last-child{margin:0}
code{background:#0F172A;border:1px solid #1E293B;border-radius:6px;padding:1px 6px;font:14px ui-monospace,Menlo,Consolas,monospace}
pre{background:#0F172A;border:1px solid #334155;border-radius:10px;padding:14px;overflow-x:auto}pre code{border:0;padding:0;background:none}
hr{border:0;border-top:1px solid #1E293B;margin:32px 0}
.table-wrap{overflow-x:auto;margin:0 0 18px}table{border-collapse:collapse;width:100%;font-size:15px}
th,td{border:1px solid #334155;padding:8px 10px;text-align:left}th{background:#0F172A}
.cta{margin:44px 0 8px;padding:24px;background:linear-gradient(135deg,#0F172A,#111C30);border:1px solid #10B981;border-radius:16px}
.cta h2{margin:0 0 6px;font-size:21px}.cta p{margin:0 0 14px;color:#CBD5E1;font-size:16px}
.hero{padding:8px 0 12px}
.group{margin-top:36px}.group h2{margin:0 0 4px}.group p.blurb{color:#94A3B8;margin:0 0 14px;font-size:16px}
.cards{display:grid;grid-template-columns:repeat(auto-fill,minmax(280px,1fr));gap:14px}
.card{display:flex;flex-direction:column;gap:6px;padding:18px;background:#0F172A;border:1px solid #334155;border-radius:14px;text-decoration:none;color:#E2E8F0}
.card:hover{border-color:#10B981}.card strong{color:#fff;font-size:17px;line-height:1.35}.card span{color:#94A3B8;font-size:14.5px;line-height:1.5}
.crumbs{font-size:14px;color:#94A3B8;margin-bottom:18px}.crumbs a{color:#94A3B8}
footer.site{border-top:1px solid #1E293B;padding:28px 0 40px;color:#94A3B8;font-size:14px;line-height:1.6}
footer.site a{color:#CBD5E1}
@media (max-width:520px){body{font-size:16.5px}.brand small{display:none}main{padding-top:28px}}
"""

FAVICON = ("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 64 64'%3E%3Crect width='64' height='64' rx='14' "
           "fill='%2310B981'/%3E%3Cpath d='M20 16h26v8H29v8h13v8H29v10h-9z' fill='%23052E16'/%3E%3C/svg%3E")


def _layout(cfg, title, description, canonical, body, extra_head=""):
    e = html.escape
    return f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{e(title)}</title>
<meta name="description" content="{e(description)}">
<link rel="canonical" href="{e(canonical)}">
<meta name="theme-color" content="#0B1220">
<link rel="icon" href="{FAVICON}">
<meta property="og:type" content="website">
<meta property="og:site_name" content="{e(cfg['product_name'])}">
<meta property="og:title" content="{e(title)}">
<meta property="og:description" content="{e(description)}">
<meta property="og:url" content="{e(canonical)}">
<meta name="twitter:card" content="summary">
{extra_head}<style>{CSS}</style>
</head>
<body>
<header class="site"><div class="wrap wide">
<a class="brand" href="/"><span class="logo">F</span>{e(cfg['product_name'])} <small>Guides</small></a>
<a class="btn" href="{e(cfg['app_url'])}/?utm_source=seo&amp;utm_medium=guides_header">Open {e(cfg['product_name'])}</a>
</div></header>
<main><div class="wrap">
{body}
</div></main>
<footer class="site"><div class="wrap">
<p>{e(cfg['required_disclosure'])} {e(cfg['product_name'])} does not guarantee any grade, interview or hiring outcome.</p>
<p>&copy; {datetime.date.today().year} {e(cfg['product_name'])} &bull; <a href="{e(cfg['app_url'])}">Open the app</a> &bull; <a href="/">All guides</a> &bull; <a href="mailto:{e(cfg['contact_email'])}">{e(cfg['contact_email'])}</a></p>
</div></footer>
</body>
</html>
"""


def build_page(cfg, meta, body_md):
    allowed = {urlparse(cfg["site_url"]).hostname, urlparse(cfg["app_url"]).hostname}
    e = html.escape
    slug, engine = meta["slug"], meta["engine"]
    canonical = f"{cfg['site_url']}/guides/{slug}/"
    cta_url = f"{cfg['app_url']}/?tool={engine}&utm_source=seo&utm_campaign={slug}"
    group = next((g for g in GROUPS if g[0] == engine), None)
    crumb = f'<a href="/">All guides</a>' + (f" &rsaquo; {e(group[1])}" if group else "")
    ld = {"@context": "https://schema.org", "@type": "Article", "headline": meta["title"], "description": meta["description"],
          "datePublished": meta["date"], "dateModified": meta["date"],
          "author": {"@type": "Organization", "name": cfg["product_name"]},
          "publisher": {"@type": "Organization", "name": cfg["product_name"]}, "mainEntityOfPage": canonical}
    ld_json = json.dumps(ld).replace("<", "\\u003c")
    body = (f'<div class="crumbs">{crumb}</div>\n<article>\n<h1>{e(meta["title"])}</h1>\n'
            f'<p class="meta">Updated {e(meta["date"])}</p>\n{render_markdown(body_md, allowed)}\n</article>\n'
            f'<aside class="cta"><h2>Try it on your own text</h2>'
            f'<p>The free preview shows the opening of your result and the first audit item. A one-time $1.99 pass per tool unlocks the full result. No subscription.</p>'
            f'<a class="btn" href="{e(cta_url)}">Start the free preview</a></aside>')
    return _layout(cfg, f"{meta['title']} | {cfg['product_name']}", meta["description"], canonical, body,
                   f'<script type="application/ld+json">{ld_json}</script>\n')


def build_index(cfg, pages):
    e = html.escape
    parts = ['<section class="hero"><h1>Guides for papers, AI drafts and resumes</h1>'
             '<p class="lede">Plain-English how-tos for formatting documents in Word, written to be useful whether or not you ever use the tool.</p></section>']
    for key, heading, blurb in GROUPS:
        items = [p for p in pages if p["engine"] == key]
        if not items:
            continue
        parts.append(f'<section class="group"><h2>{e(heading)}</h2><p class="blurb">{e(blurb)}</p><div class="cards">')
        for p in sorted(items, key=lambda x: x["title"].lower()):
            parts.append(f'<a class="card" href="/guides/{e(p["slug"])}/"><strong>{e(p["title"])}</strong><span>{e(p["description"])}</span></a>')
        parts.append("</div></section>")
    if not pages:
        parts.append("<p>Guides are on their way.</p>")
    return _layout(cfg, f"{cfg['product_name']} Guides: formatting papers, AI drafts and resumes",
                   "Practical guides to formatting papers, cleaning up AI-generated drafts and laying out resumes in Word.",
                   cfg["site_url"] + "/", "\n".join(parts))


def build_404(cfg):
    body = ('<h1>Page not found</h1><p>That page does not exist or has moved.</p>'
            '<p><a class="btn" href="/">See all guides</a></p>')
    return _layout(cfg, "Page not found | " + cfg["product_name"], "Page not found.", cfg["site_url"] + "/", body,
                   '<meta name="robots" content="noindex">\n')


def build_all(content_dir, out_dir, cfg):
    content_dir, out_dir = Path(content_dir), Path(out_dir)
    pages, problems = [], []
    for path in sorted(content_dir.glob("*.md")):
        try:
            meta, body = parse_front_matter(path.read_text(encoding="utf-8"))
            for k in ("title", "description", "slug", "engine", "date"):
                if not meta.get(k):
                    raise ValueError(f"missing {k}")
            if not re.fullmatch(r"[a-z0-9]+(-[a-z0-9]+)*", meta["slug"]):
                raise ValueError("bad slug")
            if meta["slug"] != path.stem:
                raise ValueError(f"slug must match the file name ({path.stem})")
            pages.append((meta, body))
        except ValueError as ex:
            problems.append(f"{path.name}: {ex}")
    if problems:
        raise SystemExit("Not building. Fix these pages first:\n  " + "\n  ".join(problems))

    for sub in ("guides",):
        target = out_dir / sub
        if target.exists():
            for f in sorted(target.rglob("*"), reverse=True):
                f.unlink() if f.is_file() else f.rmdir()
    out_dir.mkdir(parents=True, exist_ok=True)
    meta_list = [m for m, _ in pages]
    (out_dir / "index.html").write_text(build_index(cfg, meta_list), encoding="utf-8")
    (out_dir / "404.html").write_text(build_404(cfg), encoding="utf-8")
    for meta, body in pages:
        d = out_dir / "guides" / meta["slug"]
        d.mkdir(parents=True, exist_ok=True)
        (d / "index.html").write_text(build_page(cfg, meta, body), encoding="utf-8")
    urls = [cfg["site_url"] + "/"] + [f"{cfg['site_url']}/guides/{m['slug']}/" for m in meta_list]
    dates = [None] + [m["date"] for m in meta_list]
    sm = ['<?xml version="1.0" encoding="UTF-8"?>', '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">']
    for u, d in zip(urls, dates):
        sm.append(f"<url><loc>{html.escape(u)}</loc>" + (f"<lastmod>{d}</lastmod>" if d else "") + "</url>")
    sm.append("</urlset>")
    (out_dir / "sitemap.xml").write_text("\n".join(sm) + "\n", encoding="utf-8")
    (out_dir / "robots.txt").write_text(f"User-agent: *\nAllow: /\n\nSitemap: {cfg['site_url']}/sitemap.xml\n", encoding="utf-8")
    return len(pages)


def load_config():
    return json.loads((ROOT / "site.json").read_text(encoding="utf-8"))


if __name__ == "__main__":
    count = build_all(os.environ.get("CONTENT_DIR", ROOT / "content"), os.environ.get("OUT_DIR", ROOT / "dist"), load_config())
    print(f"Built {count} guide page(s) into dist/")
