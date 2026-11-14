"""Worker (writes), Reviewer (audits), and the loop that connects them."""
import json
import re
from dataclasses import dataclass, field

from checks import Issue, run_checks
from llm import Budget


def _facts(claims: dict) -> str:
    keep = {k: v for k, v in claims.items() if not k.startswith("_") and k != "owner_confirmed"}
    return json.dumps(keep, indent=2)


WORKER_SYSTEM = """You write honest, useful SEO landing pages for {product}, a small paid web tool.

GROUND TRUTH (the only product facts you may state):
{facts}

RULES
- State product facts only if they appear in the ground truth above. Never invent features, numbers, user counts,
  reviews, testimonials, awards, comparisons with named competitors, speed claims, or guarantees.
- Be catchy and specific, but describe what the tool does, not what it will do for the reader's grades or career.
- Be upfront about limits. Include a short "What {product} does not do" section drawn from does_not.
- {style_fact_policy}
- Do not mention AI detectors or plagiarism tools at all.
- Include the privacy_statement and required_disclosure from the ground truth word for word, each in its own
  paragraph, and mention the price exactly as given in pricing (only $1.99 may appear).
- Link to the app at least once using this exact URL: {app_url}/?tool={engine}&utm_source=seo&utm_campaign={slug}
  Do not link anywhere else.
- If you include sample text, introduce it with the exact words "Illustrative example (made-up text)". Sample text
  must be obviously fictional and must not look like a real source or a real person's resume.
- Genuinely help the reader first (explain the problem and how to fix it by hand), then show where the tool fits.

OUTPUT: Markdown only, no code fences around the whole thing, starting with this front matter:
---
title: <max 70 chars>
description: <50-160 chars>
slug: {slug}
engine: {engine}
date: {date}
---
Then 500-1600 words with at least three '## ' headings.
"""

REVIEWER_SYSTEM = """You are a strict compliance and fact reviewer for marketing copy about {product}.
Your job is to stop misleading, overclaiming, or inaccurate content before it is published. If unsure, fail it.

GROUND TRUTH:
{facts}

Check the draft for:
1. Product claims not supported by the ground truth (features, limits, pricing, privacy, what each tool does).
2. Promises or implications about grades, interviews, jobs, ATS results, detection tools, or time saved.
3. Invented statistics, testimonials, user counts, endorsements, or named competitors.
4. Style-guide facts (APA/MLA/Chicago/IEEE, resume conventions) that are wrong or stated with false certainty.
5. Sample text that could be mistaken for a real source, quote, person, or result.
6. Tone: catchy is fine; pressure, hype, or anything a customer could later call misleading is not.
7. Does it genuinely help a reader who never buys? If it is mostly a sales pitch, fail it.

Return JSON only:
{{"verdict": "pass" | "fail",
  "issues": [{{"severity": "block" | "note", "quote": "<exact text>", "problem": "<why>", "fix": "<how>"}}],
  "summary": "<two sentences>"}}
Use severity "block" for anything that must change before publishing. Verdict is "pass" only if there are no block issues."""


def _strip_fences(text: str) -> str:
    t = text.strip()
    t = re.sub(r"\A```(?:markdown|md)?\s*\n", "", t)
    t = re.sub(r"\n```\s*\Z", "", t)
    return t.strip() + "\n"


@dataclass
class ReviewResult:
    passed: bool
    issues: list = field(default_factory=list)  # list[Issue]
    notes: list = field(default_factory=list)  # list[str] non-blocking
    summary: str = ""


class Worker:
    def __init__(self, llm, model, claims, budget: Budget):
        self.llm, self.model, self.claims, self.budget = llm, model, claims, budget

    def write(self, topic, today, previous=None, issues=None) -> str:
        system = WORKER_SYSTEM.format(
            product=self.claims["product_name"], facts=_facts(self.claims),
            style_fact_policy=self.claims["style_fact_policy"], app_url=self.claims["app_url"],
            slug=topic["slug"], engine=topic["engine"], date=today)
        user = (f"Topic: {topic['slug']}\nSearch query to serve: {topic['query']}\n"
                f"Reader intent: {topic['intent']}\nTool to feature: {self.claims['engines'][topic['engine']]['name']}\n")
        if previous and issues:
            fixes = "\n".join(f"- {i}" for i in issues)
            user += f"\nYour previous draft was rejected. Fix every problem below and return the full corrected page.\n{fixes}\n\nPrevious draft:\n{previous}"
        self.budget.check()
        text, usage = self.llm.chat(self.model, system, user, json_mode=False, temperature=0.5)
        self.budget.charge(self.model, usage)
        return _strip_fences(text)


class Reviewer:
    def __init__(self, llm, model, claims, budget: Budget):
        self.llm, self.model, self.claims, self.budget = llm, model, claims, budget

    def review(self, draft: str, topic: dict) -> ReviewResult:
        system = REVIEWER_SYSTEM.format(product=self.claims["product_name"], facts=_facts(self.claims))
        user = f"Topic: {topic['slug']} (tool: {topic['engine']})\n\nDRAFT:\n{draft}"
        self.budget.check()
        text, usage = self.llm.chat(self.model, system, user, json_mode=True, temperature=0.0)
        self.budget.charge(self.model, usage)
        try:
            data = json.loads(text)
            verdict = data["verdict"]
            raw = data.get("issues", [])
            assert verdict in ("pass", "fail") and isinstance(raw, list)
        except Exception:
            # Fail closed: an unreadable review never counts as a pass.
            return ReviewResult(False, [Issue("reviewer_unreadable", "Reviewer output was not valid JSON.")])
        blocks, notes = [], []
        for r in raw:
            if not isinstance(r, dict):
                continue
            msg = f"{r.get('problem', '')} Fix: {r.get('fix', '')}".strip()
            if r.get("severity") == "block":
                blocks.append(Issue("reviewer", msg, str(r.get("quote", ""))[:120]))
            else:
                notes.append(f'"{str(r.get("quote", ""))[:80]}": {msg}')
        passed = verdict == "pass" and not blocks
        if verdict == "fail" and not blocks:
            blocks.append(Issue("reviewer", "Reviewer failed the draft without a specific issue: " + data.get("summary", "")))
        return ReviewResult(passed, blocks, notes, str(data.get("summary", "")))


@dataclass
class PageResult:
    topic: dict
    passed: bool
    draft: str
    rounds: int
    cost_usd: float
    issues: list  # outstanding issues if failed
    notes: list  # non-blocking reviewer notes
    summary: str = ""


def produce_page(topic, claims, cfg, llm, budget: Budget, today: str) -> PageResult:
    """Worker drafts -> deterministic checks -> LLM reviewer -> back to worker on any failure.

    Publishes only if the final draft clears BOTH layers. Otherwise returns passed=False.
    """
    worker = Worker(llm, cfg.worker_model, claims, budget)
    reviewer = Reviewer(llm, cfg.reviewer_model, claims, budget)
    start = budget.spent
    draft, issues, notes, summary = None, [], [], ""
    for rnd in range(1, cfg.max_rounds + 1):
        draft = worker.write(topic, today, previous=draft, issues=issues)
        issues = run_checks(draft, claims, topic)
        if issues:  # cheap code checks first; skip the paid reviewer call
            continue
        review = reviewer.review(draft, topic)
        notes, summary = review.notes, review.summary
        if review.passed:
            return PageResult(topic, True, draft, rnd, budget.spent - start, [], notes, summary)
        issues = review.issues
    return PageResult(topic, False, draft or "", cfg.max_rounds, budget.spent - start, issues, notes, summary)
