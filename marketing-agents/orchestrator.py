"""Supervisor: picks topics, runs the worker/reviewer loop, enforces budgets, publishes drafts for approval.

Run weekly as a Railway cron service:   python orchestrator.py
"""
import datetime
import json
import sys

from agents import produce_page
from backends import DryRunPublisher, GitHubPublisher, GitHubStore, LocalStore
from config import ROOT, Config
from github_api import GitHub
from llm import Budget, BudgetExceeded, OpenAILLM


def run(cfg, claims, topics, llm, publisher, store, today=None, log=print):
    today = today or datetime.date.today()
    ledger = store.load()
    month = today.strftime("%Y-%m")
    if ledger.get("month") != month:
        ledger["month"], ledger["spend_usd"] = month, 0.0
    published = ledger.setdefault("published", [])
    failed = ledger.setdefault("failed_runs", {})

    budget = Budget(cfg.max_run_usd, cfg.monthly_budget_usd - ledger["spend_usd"])
    pending = [t for t in topics
               if t["slug"] not in published and failed.get(t["slug"], 0) < cfg.max_failed_runs_per_topic]
    lines = []
    try:
        for topic in pending[:cfg.max_pages_per_run]:
            try:
                res = produce_page(topic, claims, cfg, llm, budget, today.isoformat())
            except BudgetExceeded as e:
                publisher.alert("Budget cap hit, run stopped", str(e))
                lines.append(f"STOPPED: {e}")
                break
            if res.passed:
                where = publisher.publish(res, claims, today.isoformat())
                published.append(topic["slug"])
                lines.append(f"READY  {topic['slug']} ({res.rounds} round(s), ${res.cost_usd:.3f}) -> {where}")
            else:
                failed[topic["slug"]] = failed.get(topic["slug"], 0) + 1
                detail = "\n".join(f"- {i}" for i in res.issues)
                lines.append(f"REJECTED {topic['slug']} after {res.rounds} rounds")
                if failed[topic["slug"]] >= cfg.max_failed_runs_per_topic:
                    publisher.alert(f"Topic keeps failing review: {topic['slug']}",
                                    f"The agents could not produce a compliant page after {failed[topic['slug']]} runs.\n"
                                    f"Outstanding issues:\n{detail}\n\nThe topic is now skipped. Adjust claims.json or topics.json.")
    finally:
        ledger["spend_usd"] = round(ledger["spend_usd"] + budget.spent, 4)
        store.save(ledger)
    if not pending:
        lines.append("No pending topics. Add more to topics.json.")
    lines.append(f"Spend this run: ${budget.spent:.3f} | month to date: ${ledger['spend_usd']:.2f} of ${cfg.monthly_budget_usd:.2f}")
    for line in lines:
        log(line)
    return lines


def main():
    cfg = Config.from_env()
    claims = json.loads((ROOT / "claims.json").read_text())
    topics = json.loads((ROOT / "topics.json").read_text())

    if not claims.get("owner_confirmed"):
        sys.exit("claims.json: review the facts, set contact_email, then set owner_confirmed to true. Nothing runs until then.")
    if "SET_ME" in claims.get("contact_email", "SET_ME"):
        sys.exit("claims.json: contact_email is still the placeholder.")

    llm = OpenAILLM()
    if cfg.dry_run:
        store, publisher = LocalStore(), DryRunPublisher()
        print("DRY RUN: drafts are written to ./out, nothing is opened on GitHub.")
    else:
        if not (cfg.github_repo and cfg.github_token):
            sys.exit("Set GITHUB_REPO and GITHUB_TOKEN, or run with DRY_RUN=1.")
        gh = GitHub(cfg.github_repo, cfg.github_token, cfg.base_branch)
        store, publisher = GitHubStore(gh, cfg.state_branch), GitHubPublisher(gh, cfg.content_dir)
    run(cfg, claims, topics, llm, publisher, store)


if __name__ == "__main__":
    main()
