"""Minimal GitHub REST client (standard library only)."""
import base64
import json
import urllib.error
import urllib.request


class GitHubError(RuntimeError):
    def __init__(self, status, message):
        super().__init__(message)
        self.status = status


class GitHub:
    def __init__(self, repo: str, token: str, base_branch: str = "main"):
        self.repo, self.token, self.base = repo, token, base_branch

    def _req(self, method, path, body=None):
        url = f"https://api.github.com/repos/{self.repo}{path}"
        data = json.dumps(body).encode() if body is not None else None
        req = urllib.request.Request(url, data=data, method=method, headers={
            "Authorization": f"Bearer {self.token}",
            "Accept": "application/vnd.github+json",
            "X-GitHub-Api-Version": "2022-11-28",
            "User-Agent": "formatforge-marketing-agents",
        })
        try:
            with urllib.request.urlopen(req, timeout=30) as r:
                raw = r.read()
                return json.loads(raw) if raw else {}
        except urllib.error.HTTPError as e:
            raise GitHubError(e.code, f"{method} {path} -> {e.code}: {e.read()[:300]!r}") from e

    def branch_sha(self, branch):
        return self._req("GET", f"/git/ref/heads/{branch}")["object"]["sha"]

    def branch_exists(self, branch):
        try:
            self.branch_sha(branch)
            return True
        except GitHubError as e:
            if e.status == 404:
                return False
            raise

    def create_branch(self, name, from_branch=None):
        sha = self.branch_sha(from_branch or self.base)
        self._req("POST", "/git/refs", {"ref": f"refs/heads/{name}", "sha": sha})

    def get_file(self, path, ref):
        """Return (text, sha) or (None, None) if the file does not exist."""
        try:
            r = self._req("GET", f"/contents/{path}?ref={ref}")
        except GitHubError as e:
            if e.status == 404:
                return None, None
            raise
        return base64.b64decode(r["content"]).decode(), r["sha"]

    def put_file(self, path, text, branch, message, sha=None):
        body = {"message": message, "content": base64.b64encode(text.encode()).decode(), "branch": branch}
        if sha:
            body["sha"] = sha
        self._req("PUT", f"/contents/{path}", body)

    def open_pr(self, title, head, body, draft=True):
        r = self._req("POST", "/pulls", {"title": title, "head": head, "base": self.base, "body": body, "draft": draft})
        return r["html_url"]

    def open_issue(self, title, body):
        r = self._req("POST", "/issues", {"title": title, "body": body})
        return r["html_url"]
