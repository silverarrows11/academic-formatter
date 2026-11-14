# FormatForge guides website

A small static website (no frameworks, no trackers, no scripts for visitors) that publishes the Markdown guides in
`content/`. The marketing agents add pages there through pull requests; merging a pull request publishes the page.

- `python build_site.py` turns `content/*.md` into `dist/` (pages, home, sitemap.xml, robots.txt, 404).
- `python server.py` serves `dist/` on `$PORT`.
- Railway: new service from this repo, **root directory `marketing-site`**, watch path `/marketing-site/**`.
  `railway.json` already sets the start command (`python build_site.py && python server.py`) and health check.
  Add the custom domain `guides.formatforge.help` to this service.
- `site.json` holds the site name, addresses, contact email and disclosure line. Keep it equal to `claims.json`.
- A page with broken front matter stops the build and names the file, so a bad page can never take the site down
  silently: Railway keeps serving the last good version.

Run the tests: `python -I -m unittest discover -s tests`
