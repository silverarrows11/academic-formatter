"""Deterministic guardrails. Plain code, no LLM: these catch what a model reviewing a model can miss.

A draft is only eligible for the LLM reviewer if this returns no issues.
"""
import re
from dataclasses import dataclass
from urllib.parse import urlparse


@dataclass
class Issue:
    code: str
    message: str
    quote: str = ""

    def __str__(self):
        q = f' (text: "{self.quote}")' if self.quote else ""
        return f"[{self.code}] {self.message}{q}"


def norm(s: str) -> str:
    return " ".join(s.split())


def parse_front_matter(text: str):
    m = re.match(r"\A---[ \t]*\n(.*?)\n---[ \t]*\n(.*)\Z", text.strip() + "\n", re.S)
    if not m:
        return None, text
    meta = {}
    for line in m.group(1).splitlines():
        if ":" in line:
            k, v = line.split(":", 1)
            meta[k.strip()] = v.strip().strip("\"'")
    return meta, m.group(2)


# (code, regex, message). All case-insensitive. Any match blocks the draft.
BLOCKED = [
    ("detector_evasion",
     r"undetect\w*|bypass\w*|evad\w+|humani[sz]\w*|turnitin|gptzero|originality\.ai|copyleaks"
     r"|ai[- ]?detect\w*|plagiari\w+",
     "Do not mention AI detectors, plagiarism tools, or evading them. FormatForge makes no claim about them."),
    ("outcome_promise",
     r"(get|land|win|score)\s+(you\s+|yourself\s+)?(hired|the job|that job|an interview|interviews|a job|an offer)"
     r"|(higher|better|top|good)\s+grades?|\b(improve|boost|raise|increase)\s+your\s+(grade|gpa|score|chances|odds)"
     r"|\bpass(es)?\s+(any\s+|every\s+)?(ats|screening)|beat\s+(the\s+)?ats|get\s+an\s+a\b",
     "Do not promise grades, interviews, jobs or ATS outcomes."),
    ("superlative",
     r"\bbest(?!\s+practices?)\b|#1\b|\bnumber one\b|\bleading\b|\btop[- ]rated\b|\bworld'?s\b|\bultimate\b"
     r"|\bperfect(ly)?\b|\bflawless\w*|\binstant(ly)?\b|\bmagic(al)?\b|\bguaranteed\b",
     "Unprovable superlative or absolute claim."),
    ("unlimited_claim",
     r"\bunlimited\b|\bno\s+limits?\b|\bwithout\s+limits?\b|\buse\s+it\s+as\s+much\s+as\s+you\s+(want|like)\b",
     "The pass is not unlimited: it is subject to fair use. Say 're-runs and edits included (fair use)' instead."),
    ("fake_social_proof",
     r"testimonial|(customers?|users?|students?)\s+(say|love|report)|loved by|trusted by|as seen|featured in"
     r"|★|⭐|\b5[- ]star|\bfive[- ]star|\brated\b|\breviews?\s+(say|show)",
     "No testimonials, ratings, or social-proof claims. There are no real ones yet."),
    ("invented_statistic",
     r"\d\s?%|\bpercent\b|\b\d+x\b|\b\d+\s?times\s+(faster|better|more)"
     r"|\b\d[\d,.]*\+?\s*(k|m)?\s*(users|customers|students|people|documents|downloads|resumes|papers)\b"
     r"|in\s+(just\s+|only\s+)?(\d+|a few|under \d+)\s+(seconds|minutes)",
     "No statistics, counts, percentages or speed claims."),
    ("fake_urgency_or_free",
     r"limited[- ]time|act now|hurry|only \d+ left|discount|\bsale\b|free forever|always free|completely free"
     r"|100% free|no cost|\bcheap(est)?\b",
     "No urgency tactics or unsupported free/cheap claims."),
    ("false_billing_terms",
     r"subscri\w+|recurring|monthly (fee|plan)|cancel anytime",
     "The pass is a one-time payment, not a subscription."),
    ("privacy_overclaim",
     r"secur\w+|encrypt\w*|gdpr|ccpa|hipaa|soc\s?2|iso\s?27001|zero[- ]retention|private\b|confidential\w*"
     r"|anonym\w+|never (stored|saved|shared|sent|seen)|deleted?\s+(immediately|after)|\bsafe\b",
     "Only the approved privacy statement may be used. No other privacy, security or compliance wording."),
    ("placeholder",
     r"\bTODO\b|\bTBD\b|lorem ipsum|\[insert|\[your|\bxxx\b",
     "Placeholder text left in draft."),
]


def _strip_approved(text: str, claims: dict) -> str:
    t = norm(text)
    for key in ("privacy_statement", "required_disclosure"):
        s = claims.get(key)
        if s:
            t = t.replace(norm(s), " ")
    return t


def run_checks(draft: str, claims: dict, topic: dict) -> list[Issue]:
    issues: list[Issue] = []
    meta, body = parse_front_matter(draft)

    # ---- structure ----
    if meta is None:
        return [Issue("front_matter", "Draft must start with a front matter block between --- lines.")]
    for key in ("title", "description", "slug", "engine", "date"):
        if not meta.get(key):
            issues.append(Issue("front_matter", f"Missing front matter key: {key}."))
    if meta.get("slug") and meta["slug"] != topic["slug"]:
        issues.append(Issue("front_matter", f'slug must be exactly "{topic["slug"]}".'))
    if meta.get("engine") and meta["engine"] != topic["engine"]:
        issues.append(Issue("front_matter", f'engine must be exactly "{topic["engine"]}".'))
    if meta.get("date") and not re.fullmatch(r"\d{4}-\d{2}-\d{2}", meta["date"]):
        issues.append(Issue("front_matter", "date must look like YYYY-MM-DD."))
    if len(meta.get("title", "")) > 70:
        issues.append(Issue("front_matter", "title must be 70 characters or fewer."))
    d = len(meta.get("description", ""))
    if meta.get("description") and not 50 <= d <= 160:
        issues.append(Issue("front_matter", f"description must be 50-160 characters (is {d})."))

    words = len(body.split())
    if not 500 <= words <= 1600:
        issues.append(Issue("length", f"Body must be 500-1600 words (is {words})."))
    if len(re.findall(r"^##\s", body, re.M)) < 3:
        issues.append(Issue("structure", "Use at least three '## ' section headings."))

    # ---- mandatory statements ----
    ntext = norm(draft)
    for key, label in (("privacy_statement", "approved privacy statement"),
                       ("required_disclosure", "required disclosure line")):
        s = claims.get(key)
        if s and norm(s) not in ntext:
            issues.append(Issue(key, f"The {label} must appear word for word: {s}"))

    # ---- links ----
    allowed_hosts = {urlparse(claims["site_url"]).hostname, urlparse(claims["app_url"]).hostname}
    app_host = urlparse(claims["app_url"]).hostname
    urls = [u.rstrip(".,;:") for u in re.findall(r"https?://[^\s)>\]\"']+", draft)]
    for u in urls:
        if urlparse(u).hostname not in allowed_hosts:   # exact host match: formatforge.help.evil.com must not pass
            issues.append(Issue("external_link", "Only links to formatforge.help are allowed.", u))
    if not any(urlparse(u).hostname == app_host for u in urls):
        issues.append(Issue("cta", f"Include at least one link to the app: {claims['app_url']}"))

    # ---- illustrative samples must be labelled ----
    if "```" in body or re.search(r"^>\s", body, re.M):
        if "Illustrative example (made-up text)" not in body:
            issues.append(Issue(
                "unlabelled_sample",
                'Any sample text must be introduced with the exact words "Illustrative example (made-up text)".'))

    # ---- banned wording ----
    scan = _strip_approved(draft, claims)
    scan = re.sub(r"\bnot a subscription\b", " ", scan, flags=re.I)
    for code, pattern, message in BLOCKED:
        for m in re.finditer(pattern, scan, re.I):
            issues.append(Issue(code, message, m.group(0)))
            break  # one finding per rule keeps feedback short

    if not claims.get("refund_policy"):
        m = re.search(r"refund\w*|money[- ]back|satisfaction", scan, re.I)
        if m:
            issues.append(Issue("refund", "No refund policy is approved yet. Do not mention refunds.", m.group(0)))

    # guarantee only allowed in a negated sentence ("does not guarantee ...")
    for s in re.split(r"(?<=[.!?])\s+", scan):
        if re.search(r"guarante", s, re.I) and not re.search(r"\b(not|no|never|cannot)\b|n't|can't", s, re.I):
            issues.append(Issue("guarantee", "Never promise or guarantee outcomes.", s[:80]))

    # ---- prices: only the approved one ----
    for m in re.finditer(r"\$\s?\d[\d,.]*", scan):
        if m.group(0).replace(" ", "") != "$1.99":
            issues.append(Issue("price", "The only approved price is $1.99.", m.group(0)))

    return issues
