import http.client
import json
import sys
import tempfile
import threading
import unittest
from functools import partial
from http.server import ThreadingHTTPServer
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
import build_site as B  # noqa: E402

CFG = B.load_config()
PAGE = """---
title: A test page
description: A description that is long enough for a test page to be valid here.
slug: a-test-page
engine: academic
date: 2026-11-14
---
Intro with **bold**, *italic*, `code` and a [link](https://formatforge.help/x).

## Section one

- one
- two

1. first
2. second

> Illustrative example (made-up text)

| a | b |
|---|---|
| 1 | 2 |
"""


def build(pages):
    tmp = Path(tempfile.mkdtemp())
    (tmp / "content").mkdir()
    for name, text in pages.items():
        (tmp / "content" / f"{name}.md").write_text(text, encoding="utf-8")
    n = B.build_all(tmp / "content", tmp / "dist", CFG)
    return tmp / "dist", n


class Rendering(unittest.TestCase):
    def test_basic_markdown(self):
        h = B.render_markdown(PAGE.split("---\n", 2)[2], {"formatforge.help"})
        for needle in ("<strong>bold</strong>", "<em>italic</em>", "<code>code</code>", "<h2", "<ul>", "<ol>", "<blockquote>", "<table>"):
            self.assertIn(needle, h)

    def test_html_and_script_are_escaped(self):
        h = B.render_markdown("<script>alert(1)</script> and <img src=x onerror=alert(1)>", set())
        self.assertNotIn("<script", h)
        self.assertNotIn("<img", h)
        self.assertIn("&lt;script&gt;", h)

    def test_dangerous_link_schemes_are_dropped(self):
        for bad in ("javascript:alert(1)", "data:text/html,x", "vbscript:x", "JaVaScRiPt:alert(1)"):
            h = B.render_markdown(f"[click]({bad})", set())
            self.assertNotIn("href", h, bad)

    def test_external_links_are_marked_but_own_links_are_not(self):
        h = B.render_markdown("[a](https://example.org) [b](https://formatforge.help/x)", {"formatforge.help"})
        self.assertIn('rel="nofollow noopener noreferrer"', h.split("[b]")[0])
        self.assertEqual(h.count("nofollow"), 1)

    def test_headings_never_become_second_h1(self):
        self.assertNotIn("<h1", B.render_markdown("# Title", set()))

    def test_code_block_is_literal(self):
        h = B.render_markdown("```\n<b>x</b> **y**\n```", set())
        self.assertIn("&lt;b&gt;x&lt;/b&gt; **y**", h)


class Building(unittest.TestCase):
    def test_site_files(self):
        dist, n = build({"a-test-page": PAGE})
        self.assertEqual(n, 1)
        page = (dist / "guides/a-test-page/index.html").read_text()
        for needle in ('<link rel="canonical" href="https://guides.formatforge.help/guides/a-test-page/">',
                       '<meta name="description"', "application/ld+json", "utm_campaign=a-test-page", "tool=academic",
                       CFG["required_disclosure"], CFG["contact_email"], 'name="viewport"'):
            self.assertIn(needle, page)
        self.assertEqual(page.count("<h1"), 1)
        sm = (dist / "sitemap.xml").read_text()
        self.assertIn("https://guides.formatforge.help/guides/a-test-page/", sm)
        self.assertIn("Sitemap: https://guides.formatforge.help/sitemap.xml", (dist / "robots.txt").read_text())
        self.assertIn("noindex", (dist / "404.html").read_text())
        index = (dist / "index.html").read_text()
        self.assertIn("A test page", index)
        self.assertIn('href="/guides/a-test-page/"', index)

    def test_no_scripts_trackers_or_external_resources(self):
        dist, _ = build({"a-test-page": PAGE})
        for f in dist.rglob("*.html"):
            t = f.read_text()
            self.assertNotRegex(t, r"<script(?![^>]*ld\+json)", f.name)
            for bad in ("googletagmanager", "google-analytics", "facebook", "<iframe", "<link rel=\"stylesheet\""):
                self.assertNotIn(bad, t, f.name)

    def test_title_with_markup_is_escaped(self):
        dist, _ = build({"a-test-page": PAGE.replace("A test page", "<b>Hi</b> & co")})
        page = (dist / "guides/a-test-page/index.html").read_text()
        self.assertNotIn("<b>Hi</b>", page)
        self.assertIn("&lt;b&gt;Hi&lt;/b&gt; &amp; co", page)

    def test_broken_page_stops_the_build_and_names_it(self):
        with self.assertRaises(SystemExit) as cm:
            build({"a-test-page": PAGE, "bad": "no front matter here"})
        self.assertIn("bad.md", str(cm.exception))
        with self.assertRaises(SystemExit) as cm:
            build({"wrong-name": PAGE})
        self.assertIn("must match the file name", str(cm.exception))

    def test_removed_page_disappears_from_the_site(self):
        tmp = Path(tempfile.mkdtemp()); (tmp / "content").mkdir()
        (tmp / "content/a-test-page.md").write_text(PAGE)
        B.build_all(tmp / "content", tmp / "dist", CFG)
        (tmp / "content/a-test-page.md").unlink()
        B.build_all(tmp / "content", tmp / "dist", CFG)
        self.assertFalse((tmp / "dist/guides/a-test-page").exists())
        self.assertIn("on their way", (tmp / "dist/index.html").read_text())


class Consistency(unittest.TestCase):
    CLAIMS = ROOT.parent / "marketing-agents" / "claims.json"

    @unittest.skipUnless(CLAIMS.exists(), "marketing-agents not next to this folder")
    def test_site_json_matches_claims(self):
        c = json.loads(self.CLAIMS.read_text())
        for k, v in CFG.items():
            self.assertEqual(c[k], v, k)

    @unittest.skipUnless(CLAIMS.exists(), "marketing-agents not next to this folder")
    def test_every_content_page_passes_the_agents_own_checks(self):
        sys.path.insert(0, str(ROOT.parent / "marketing-agents"))
        from checks import run_checks
        c = json.loads(self.CLAIMS.read_text())
        pages = list((ROOT / "content").glob("*.md"))
        self.assertTrue(pages)
        for p in pages:
            meta, _ = B.parse_front_matter(p.read_text(encoding="utf-8"))
            issues = run_checks(p.read_text(encoding="utf-8"), c, {"slug": meta["slug"], "engine": meta["engine"]})
            self.assertEqual([str(i) for i in issues], [], p.name)


class Server(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        import server as S
        cls.dist, _ = build({"a-test-page": PAGE})
        S.DIST = cls.dist
        cls.httpd = ThreadingHTTPServer(("127.0.0.1", 0), partial(S.Handler, directory=str(cls.dist)))
        cls.port = cls.httpd.server_address[1]
        threading.Thread(target=cls.httpd.serve_forever, daemon=True).start()

    @classmethod
    def tearDownClass(cls):
        cls.httpd.shutdown()

    def get(self, path, method="GET"):
        c = http.client.HTTPConnection("127.0.0.1", self.port, timeout=5)
        c.request(method, path)
        r = c.getresponse()
        return r.status, dict(r.getheaders()), r.read()

    def test_pages_and_headers(self):
        status, h, body = self.get("/guides/a-test-page/")
        self.assertEqual(status, 200)
        self.assertIn(b"A test page", body)
        self.assertEqual(h["X-Content-Type-Options"], "nosniff")
        self.assertIn("default-src 'none'", h["Content-Security-Policy"])
        self.assertEqual(self.get("/")[0], 200)
        self.assertEqual(self.get("/robots.txt")[0], 200)
        self.assertEqual(self.get("/sitemap.xml")[0], 200)

    def test_slashless_url_redirects_and_404_is_a_real_404(self):
        self.assertIn(self.get("/guides/a-test-page")[0], (301, 302))
        status, _, body = self.get("/nope")
        self.assertEqual(status, 404)
        self.assertIn(b"Page not found", body)
        self.assertEqual(self.get("/nope", "HEAD")[0], 404)

    def test_cannot_read_outside_the_site_folder(self):
        for path in ("/../server.py", "/..%2fserver.py", "/%2e%2e/build_site.py", "/guides/../../site.json"):
            status, _, body = self.get(path)
            self.assertNotIn(b"def main", body, path)
            self.assertNotIn(b"product_name", body, path)


if __name__ == "__main__":
    unittest.main()
