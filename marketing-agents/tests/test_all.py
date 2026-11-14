import datetime
import json
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from agents import produce_page  # noqa: E402
from backends import DryRunPublisher, LocalStore  # noqa: E402
from checks import run_checks  # noqa: E402
from config import Config  # noqa: E402
from llm import Budget, BudgetExceeded, Usage  # noqa: E402
from orchestrator import run  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
CLAIMS = json.loads((ROOT / "claims.json").read_text())
CLAIMS["owner_confirmed"] = True
TOPIC = {"slug": "apa-hanging-indent-in-word", "engine": "academic",
         "query": "how to do a hanging indent in Word APA", "intent": "x"}
TOPIC2 = {"slug": "one-page-resume-format-guide", "engine": "resume", "query": "q", "intent": "x"}


def good_draft(topic=TOPIC, extra=""):
    filler = " ".join(["Paste your text, review each change, and download the file when you are happy with it."] * 45)
    return f"""---
title: A clear guide to formatting your document in Word
description: A practical walkthrough of the manual steps, and where FormatForge fits if you would rather not do them by hand.
slug: {topic['slug']}
engine: {topic['engine']}
date: 2026-11-11
---
## The problem

{filler}

## Doing it by hand

{filler}

## Where FormatForge fits

Try it at {CLAIMS['app_url']}/?tool={topic['engine']}&utm_source=seo&utm_campaign={topic['slug']} and see a free preview first. A full result costs a one-time $1.99 pass.

{extra}

## What FormatForge does not do

It does not guarantee any outcome. It does not verify that a source exists.

## Privacy

{CLAIMS['privacy_statement']}

{CLAIMS['required_disclosure']}
"""


class CheckTests(unittest.TestCase):
    def codes(self, draft, topic=TOPIC):
        return {i.code for i in run_checks(draft, CLAIMS, topic)}

    def test_good_draft_is_clean(self):
        self.assertEqual(run_checks(good_draft(), CLAIMS, TOPIC), [])

    def test_banned_wording(self):
        cases = {
            "detector_evasion": "This makes your text undetectable.",
            "outcome_promise": "Use it to land an interview fast.",
            "superlative": "It is the best tool around.",
            "unlimited_claim": "Get unlimited re-runs for one price.",
            "fake_social_proof": "Students love it and it is trusted by many.",
            "invented_statistic": "It fixes 98% of errors.",
            "fake_urgency_or_free": "Limited time offer, act now.",
            "false_billing_terms": "Cancel anytime from your subscription.",
            "privacy_overclaim": "Your files are encrypted and fully secure.",
            "refund": "Get a refund if you are unhappy.",
            "guarantee": "We guarantee a great result.",
            "price": "Only $4.99 for everything.",
            "external_link": "See https://example.com/page for more.",
            "placeholder": "TODO add a screenshot here.",
        }
        for code, sentence in cases.items():
            with self.subTest(code=code):
                self.assertIn(code, self.codes(good_draft(extra=sentence)))

    def test_link_check_uses_exact_hosts(self):
        base = CLAIMS["app_url"]
        for bad in (base + ".evil.com/x", "https://evil.com/?u=" + base, "http://" + base.split("//")[1] + "@evil.com/"):
            with self.subTest(bad=bad):
                self.assertIn("external_link", self.codes(good_draft(extra=f"See {bad} now.")))
        self.assertNotIn("external_link", self.codes(good_draft(extra=f"See {CLAIMS['site_url']}/guides/x/ and {base}.")))

    def test_preview_shows_is_not_flagged_as_a_review_claim(self):
        self.assertNotIn("fake_social_proof", self.codes(good_draft(extra="The free preview shows the opening of your result.")))
        self.assertIn("fake_social_proof", self.codes(good_draft(extra="Reviews show people love it.")))

    def test_negated_guarantee_is_allowed(self):
        self.assertNotIn("guarantee", self.codes(good_draft(extra="We do not guarantee any grade.")))

    def test_structure_rules(self):
        self.assertIn("front_matter", self.codes("no front matter here"))
        self.assertIn("front_matter", self.codes(good_draft().replace(TOPIC["slug"], "wrong-slug", 1)))
        front = good_draft().split("## The problem")[0]
        short = front + "## A\n\nShort.\n\n## B\n\nShort.\n\n## C\n\nShort.\n\n" + CLAIMS["privacy_statement"] + "\n\n" + CLAIMS["required_disclosure"] + "\n\n" + CLAIMS["app_url"] + "\n"
        self.assertIn("length", self.codes(short))
        self.assertIn("privacy_statement", self.codes(good_draft().replace(CLAIMS["privacy_statement"], "We keep nothing.")))
        self.assertIn("required_disclosure", self.codes(good_draft().replace(CLAIMS["required_disclosure"], "")))
        self.assertIn("cta", self.codes(good_draft().replace(CLAIMS["app_url"], "")))

    def test_samples_must_be_labelled(self):
        self.assertIn("unlabelled_sample", self.codes(good_draft(extra="> before: ### Heading **bold**")))
        self.assertNotIn("unlabelled_sample",
                         self.codes(good_draft(extra="Illustrative example (made-up text)\n\n> before: ### Heading")))


class FakeLLM:
    def __init__(self, drafts, reviews, tokens=1000):
        self.drafts, self.reviews, self.tokens, self.calls = list(drafts), list(reviews), tokens, 0

    def chat(self, model, system, user, json_mode=False, temperature=0.4):
        self.calls += 1
        queue = self.reviews if json_mode else self.drafts
        return queue.pop(0), Usage(self.tokens, self.tokens)


PASS = json.dumps({"verdict": "pass", "issues": [{"severity": "note", "quote": "q", "problem": "tiny", "fix": "ok"}], "summary": "Fine."})
FAIL = json.dumps({"verdict": "fail", "issues": [{"severity": "block", "quote": "x", "problem": "overclaim", "fix": "remove"}], "summary": "No."})


class LoopTests(unittest.TestCase):
    cfg = Config(max_rounds=3, dry_run=True)

    def page(self, llm, run_usd=5.0):
        return produce_page(TOPIC, CLAIMS, self.cfg, llm, Budget(run_usd, 100), "2026-11-11")

    def test_passes_first_time(self):
        res = self.page(FakeLLM([good_draft()], [PASS]))
        self.assertTrue(res.passed)
        self.assertEqual(res.rounds, 1)
        self.assertEqual(len(res.notes), 1)

    def test_code_check_failure_skips_paid_reviewer_then_recovers(self):
        bad = good_draft(extra="It is the best tool.")
        llm = FakeLLM([bad, good_draft()], [PASS])
        res = self.page(llm)
        self.assertTrue(res.passed)
        self.assertEqual(res.rounds, 2)
        self.assertEqual(llm.calls, 3)  # worker, worker, reviewer (reviewer never saw the bad draft)

    def test_reviewer_block_forces_rewrite(self):
        res = self.page(FakeLLM([good_draft(), good_draft()], [FAIL, PASS]))
        self.assertTrue(res.passed)
        self.assertEqual(res.rounds, 2)

    def test_never_passes_when_reviewer_keeps_failing(self):
        res = self.page(FakeLLM([good_draft()] * 3, [FAIL] * 3))
        self.assertFalse(res.passed)
        self.assertTrue(res.issues)

    def test_unreadable_reviewer_fails_closed(self):
        res = self.page(FakeLLM([good_draft()] * 3, ["not json"] * 3))
        self.assertFalse(res.passed)

    def test_pass_verdict_with_block_issue_is_not_a_pass(self):
        sneaky = json.dumps({"verdict": "pass", "issues": [{"severity": "block", "quote": "x", "problem": "p", "fix": "f"}]})
        res = self.page(FakeLLM([good_draft()] * 3, [sneaky] * 3))
        self.assertFalse(res.passed)

    def test_budget_cap_stops_run(self):
        llm = FakeLLM([good_draft()] * 3, [PASS] * 3, tokens=2_000_000)  # each call costs several dollars
        with self.assertRaises(BudgetExceeded):
            self.page(llm, run_usd=0.50)


class OrchestratorTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.store = LocalStore(Path(self.tmp.name) / "ledger.json")
        self.pub = DryRunPublisher(Path(self.tmp.name) / "out")
        self.alerts = []
        self.pub.alert = lambda t, b: self.alerts.append(t)
        self.day = datetime.date(2026, 11, 11)

    def go(self, llm, cfg=None, topics=(TOPIC, TOPIC2)):
        return run(cfg or Config(max_pages_per_run=2), CLAIMS, list(topics), llm, self.pub, self.store, self.day, log=lambda *_: None)

    def test_publishes_and_records_and_skips_done_topics(self):
        self.go(FakeLLM([good_draft(TOPIC), good_draft(TOPIC2)], [PASS, PASS]))
        ledger = self.store.load()
        self.assertEqual(ledger["published"], [TOPIC["slug"], TOPIC2["slug"]])
        self.assertTrue((Path(self.tmp.name) / "out" / f"{TOPIC['slug']}.md").exists())
        llm2 = FakeLLM([], [])
        self.go(llm2)
        self.assertEqual(llm2.calls, 0)  # nothing pending

    def test_failed_topic_is_not_published_and_alerts_after_repeat(self):
        cfg = Config(max_pages_per_run=1, max_failed_runs_per_topic=2)
        for _ in range(2):
            self.go(FakeLLM([good_draft()] * 3, [FAIL] * 3), cfg, topics=(TOPIC,))
        self.assertEqual(self.store.load()["published"], [])
        self.assertEqual(self.store.load()["failed_runs"][TOPIC["slug"]], 2)
        self.assertEqual(len(self.alerts), 1)
        self.assertFalse((Path(self.tmp.name) / "out").exists())

    def test_budget_hit_alerts_and_saves_spend(self):
        cfg = Config(max_pages_per_run=2, max_run_usd=0.50)
        self.go(FakeLLM([good_draft()] * 3, [PASS] * 3, tokens=2_000_000), cfg, topics=(TOPIC,))
        self.assertIn("Budget cap hit, run stopped", self.alerts)
        self.assertGreater(self.store.load()["spend_usd"], 0)

    def test_monthly_budget_carries_and_resets(self):
        self.store.save({"month": "2026-11", "spend_usd": 9.99, "published": [], "failed_runs": {}})
        cfg = Config(monthly_budget_usd=10.0)
        self.go(FakeLLM([good_draft()] * 3, [PASS] * 3, tokens=2_000_000), cfg, topics=(TOPIC,))
        self.assertIn("Budget cap hit, run stopped", self.alerts)
        self.store.save({"month": "2026-10", "spend_usd": 9.99, "published": [], "failed_runs": {}})
        self.alerts.clear()
        self.go(FakeLLM([good_draft()], [PASS]), cfg, topics=(TOPIC,))
        self.assertEqual(self.alerts, [])


if __name__ == "__main__":
    unittest.main()
