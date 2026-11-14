"""Thin OpenAI wrapper plus cost accounting and the budget cap."""
from dataclasses import dataclass

from config import PRICES


@dataclass
class Usage:
    prompt_tokens: int = 0
    completion_tokens: int = 0

    def cost(self, model: str) -> float:
        pin, pout = PRICES.get(model, max(PRICES.values()))
        return (self.prompt_tokens * pin + self.completion_tokens * pout) / 1_000_000


class BudgetExceeded(RuntimeError):
    pass


class Budget:
    """Hard stop. Checked before every model call."""

    def __init__(self, max_run_usd: float, month_remaining_usd: float):
        self.max_run = max_run_usd
        self.month_remaining = month_remaining_usd
        self.spent = 0.0

    def check(self):
        if self.spent >= self.max_run:
            raise BudgetExceeded(f"per-run cap reached (${self.spent:.2f} of ${self.max_run:.2f})")
        if self.spent >= self.month_remaining:
            raise BudgetExceeded("monthly budget reached")

    def charge(self, model: str, usage: Usage):
        self.spent += usage.cost(model)


class OpenAILLM:
    def __init__(self, api_key: str | None = None):
        from openai import OpenAI  # imported lazily so tests run without the package

        self.client = OpenAI(api_key=api_key)

    def chat(self, model, system, user, json_mode=False, temperature=0.4):
        kwargs = dict(
            model=model,
            temperature=temperature,
            messages=[{"role": "system", "content": system}, {"role": "user", "content": user}],
        )
        if json_mode:
            kwargs["response_format"] = {"type": "json_object"}
        r = self.client.chat.completions.create(**kwargs)
        u = r.usage
        return r.choices[0].message.content or "", Usage(u.prompt_tokens, u.completion_tokens)
