"""Settings, read from environment variables (set them in Railway)."""
import os
from dataclasses import dataclass
from pathlib import Path

ROOT = Path(__file__).parent

# USD per 1M tokens (input, output). Verify against OpenAI's pricing page; unknown models
# are billed at the highest rate here so the budget cap errs on the safe side.
PRICES = {
    "gpt-4o-mini": (0.15, 0.60),
    "gpt-4o": (2.50, 10.00),
}


def _f(name, default):
    return float(os.environ.get(name, default))


def _i(name, default):
    return int(os.environ.get(name, default))


@dataclass
class Config:
    worker_model: str = "gpt-4o-mini"
    reviewer_model: str = "gpt-4o"  # a different, stronger model reduces shared blind spots
    max_rounds: int = 3  # worker/reviewer rounds per page
    max_pages_per_run: int = 2
    max_run_usd: float = 2.0
    monthly_budget_usd: float = 10.0
    max_failed_runs_per_topic: int = 2
    dry_run: bool = True  # safe default: write drafts to ./out and open nothing on GitHub
    github_repo: str = ""  # owner/name
    github_token: str = ""
    base_branch: str = "main"
    content_dir: str = "marketing-site/content"
    state_branch: str = "agent-state"

    @classmethod
    def from_env(cls):
        return cls(
            worker_model=os.environ.get("WORKER_MODEL", "gpt-4o-mini"),
            reviewer_model=os.environ.get("REVIEWER_MODEL", "gpt-4o"),
            max_rounds=_i("MAX_ROUNDS", 3),
            max_pages_per_run=_i("MAX_PAGES_PER_RUN", 2),
            max_run_usd=_f("MAX_RUN_USD", 2.0),
            monthly_budget_usd=_f("MONTHLY_BUDGET_USD", 10.0),
            max_failed_runs_per_topic=_i("MAX_FAILED_RUNS_PER_TOPIC", 2),
            dry_run=os.environ.get("DRY_RUN", "1") != "0",
            github_repo=os.environ.get("GITHUB_REPO", ""),
            github_token=os.environ.get("GITHUB_TOKEN", ""),
            base_branch=os.environ.get("BASE_BRANCH", "main"),
            content_dir=os.environ.get("CONTENT_DIR", "marketing-site/content"),
            state_branch=os.environ.get("STATE_BRANCH", "agent-state"),
        )
