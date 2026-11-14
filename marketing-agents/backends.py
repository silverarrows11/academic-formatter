"""Where state lives and where finished drafts go. Dry-run (local files) or GitHub (PRs + issues)."""
import json
from pathlib import Path

from github_api import GitHub

LEDGER = "ledger.json"


def pr_body(result, claims) -> str:
    t = result.topic
    notes = "\n".join(f"- {n}" for n in result.notes) or "- none"
    return f"""**Page:** `{t['slug']}` | **Tool:** {claims['engines'][t['engine']]['name']} | **Target query:** {t['query']}

Passed deterministic checks and the LLM reviewer in {result.rounds} round(s). Model spend: ${result.cost_usd:.3f}.
Reviewer summary: {result.summary or 'n/a'}

**Non-blocking reviewer notes**
{notes}

**Before you merge (about 2 minutes)**
- [ ] Every statement about the product matches `claims.json` (features, price, privacy)
- [ ] Any APA/MLA/Chicago/IEEE or resume-convention fact is one you'd stand behind
- [ ] Sample text is clearly labelled as made-up
- [ ] You'd be comfortable if a customer read this next to their receipt

Close the PR to reject it. Nothing here merges or deploys automatically."""


# ---------------- state ----------------
class LocalStore:
    def __init__(self, path="state/ledger.json"):
        self.path = Path(path)

    def load(self) -> dict:
        return json.loads(self.path.read_text()) if self.path.exists() else {}

    def save(self, ledger: dict):
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.path.write_text(json.dumps(ledger, indent=2))


class GitHubStore:
    """Railway's disk is ephemeral, so the ledger lives on a separate branch of your repo."""

    def __init__(self, gh: GitHub, branch: str):
        self.gh, self.branch, self.sha = gh, branch, None

    def load(self) -> dict:
        if not self.gh.branch_exists(self.branch):
            self.gh.create_branch(self.branch)
        text, self.sha = self.gh.get_file(LEDGER, self.branch)
        return json.loads(text) if text else {}

    def save(self, ledger: dict):
        self.gh.put_file(LEDGER, json.dumps(ledger, indent=2), self.branch, "Update marketing agent ledger", self.sha)


# ---------------- output ----------------
class DryRunPublisher:
    def __init__(self, out_dir="out"):
        self.out = Path(out_dir)

    def publish(self, result, claims, today) -> str:
        self.out.mkdir(parents=True, exist_ok=True)
        slug = result.topic["slug"]
        (self.out / f"{slug}.md").write_text(result.draft)
        (self.out / f"{slug}.review.md").write_text(pr_body(result, claims))
        return str(self.out / f"{slug}.md")

    def alert(self, title, body):
        print(f"ALERT: {title}\n{body}")


class GitHubPublisher:
    def __init__(self, gh: GitHub, content_dir: str):
        self.gh, self.content_dir = gh, content_dir

    def publish(self, result, claims, today) -> str:
        slug = result.topic["slug"]
        path = f"{self.content_dir}/{slug}.md"
        existing, _ = self.gh.get_file(path, self.gh.base)
        if existing is not None:
            raise RuntimeError(f"{path} already exists on {self.gh.base}; refusing to overwrite")
        branch = f"marketing/{slug}-{today}"
        self.gh.create_branch(branch)
        self.gh.put_file(path, result.draft, branch, f"Add page: {slug}")
        return self.gh.open_pr(f"[Marketing agent] {slug}", branch, pr_body(result, claims), draft=True)

    def alert(self, title, body):
        self.gh.open_issue(f"[Marketing agent] {title}", body)
