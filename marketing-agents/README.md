# FormatForge marketing agents

Drafts honest SEO pages for the marketing site and opens them as **draft pull requests**. It never merges, deploys,
posts, or spends money on ads. You approve every page.

## How a page gets made

```
topics.json -> Worker drafts -> code checks -> Reviewer (different model) -> draft PR -> you merge or close
                  ^______________ feedback, up to 3 rounds ______________|
```

Safety layers, in order:
1. **claims.json** is the only source of product facts (features, price, privacy wording). Workers and reviewers see it.
2. **Code checks** (`checks.py`) block: AI-detector or plagiarism talk, grade/job/ATS promises, superlatives, fake
   reviews, invented stats, urgency tactics, privacy/security wording other than your approved sentence, refunds
   (until you approve a policy), any price except $1.99, links off formatforge.help, unlabelled sample text.
3. **LLM reviewer** (default `gpt-4o`, writer is `gpt-4o-mini`) audits against claims.json. Unreadable output fails closed.
4. **You**: the PR has a 2-minute checklist. Drafts only.
5. **Supervisor** (`orchestrator.py`): per-run and monthly budget caps, failure counts, GitHub issue alerts.

## Where things live in your GitHub repo

```
app.py, requirements.txt      the FormatForge app           (Railway service 1, live at formatforge.help)
marketing-site/               the guides website            (Railway service 2, live at guides.formatforge.help)
marketing-agents/             these agents                  (Railway service 3, runs on a weekly schedule)
```

**Important:** in Railway, give each service a *Watch Path* so a change in one folder never redeploys the others.
Service 1: `/app.py` and `/requirements.txt`. Service 2: `/marketing-site/**`. Service 3: `/marketing-agents/**`.
Without this, merging a guide would restart your live app.

## Setup (one time)

1. Read `claims.json` line by line. Fix anything inaccurate. Set `contact_email`. Add a `refund_policy` string only if
   you decide to offer refunds. Then set `"owner_confirmed": true`. The orchestrator refuses to run until you do.
2. In the OpenAI dashboard set a **hard monthly spend limit** as a backstop to the in-code caps.
3. Create a GitHub fine-grained token for this repo only: Contents, Pull requests, Issues = read and write.
4. Add a Railway service from this folder (root directory `marketing-agents`): start command `python orchestrator.py`,
   cron `0 7 * * 1` (Mondays, UTC). Variables: `OPENAI_API_KEY`, `GITHUB_TOKEN`, `GITHUB_REPO` (owner/name), `DRY_RUN=1`.
5. Run it with `DRY_RUN=1` first. Drafts land in `./out`. Read a few. When you trust them set `DRY_RUN=0`.

## Your weekly routine (about 5 minutes)

Each Monday the agents open up to two draft pull requests. For each one: read the checklist in the description, skim the
page, then **Merge** (it goes live in a minute or two) or **Close** (it is dropped and retried or skipped). Nothing is
ever published without your merge.

Optional variables: `WORKER_MODEL`, `REVIEWER_MODEL`, `MAX_ROUNDS` (3), `MAX_PAGES_PER_RUN` (2),
`MAX_RUN_USD` (2.0), `MONTHLY_BUDGET_USD` (10.0), `CONTENT_DIR` (`marketing-site/content`), `BASE_BRANCH` (`main`).
Prices in `config.py` are estimates: check them against OpenAI's pricing page.

## Run the tests

```
python -I -m unittest discover -s tests
```

## Deliberately not included

- **Automated posting or messaging** on Reddit, forums, social media, comment sections or email. Those platforms
  restrict self-promotion and automated accounts, bans would damage the brand, and a live reply cannot be checked
  against `claims.json` before it is sent. If you ever want help there, the agents can draft text for you to read,
  but a person should post it.
- **Auto-merge.** Add only after you have reviewed 10-15 PRs and trust the output.
- **Paid ads.** They need your ad account, a budget and conversion tracking.
- Analytics agent (needs Search Console credentials).

## Notes

- App links look like `/?tool=<engine>&utm_source=seo&utm_campaign=<slug>`. The app reads `tool`
  (academic, resume, humanizer, or the alias cleaner) to open the right engine, and ignores the utm parameters.
- `claims.json` and `marketing-site/site.json` must agree (a test checks this when both folders are present).
- Keep `claims.json` in step with the app: if you ever change a limit, price or privacy wording in app.py, change it here too.
- The app only promises a refund if `REFUND_DAYS` and `CONTACT_EMAIL` are set in Railway. If you turn that on, also
  write the policy into `refund_policy` in `claims.json`; until then the agents are blocked from mentioning refunds.
- The pass is never described as "unlimited". The checks reject that word.
