"""Shared prompts, parsing and provider calls for utility billing.

This module has no Gmail, spreadsheet, invoice or filesystem write operations.
Provider SDKs are imported only when a client is requested. Prompts never contain
an answer key, and neither prompts nor model response bodies are logged here.
"""

from dataclasses import dataclass
from datetime import date
import html as html_lib
import os
import re
import time

PROVIDERS = ("azure_openai", "claude_foundry")
EXTRACTION_MAX_TOKENS = 50
EXTRACTION_TEMPERATURE = 0
COMMENTARY_MAX_TOKENS = 400
COMMENTARY_TEMPERATURE = 0.3


@dataclass(frozen=True)
class Completion:
    text: str
    latency_seconds: float
    input_tokens: int | None
    output_tokens: int | None
    provider: str
    deployment: str
    cache_read_input_tokens: int | None = None
    cache_creation_input_tokens: int | None = None
    finish_reason: str | None = None
    response_model: str | None = None


def get_provider(config: dict) -> str:
    provider = config.get("azure_openai", {}).get("provider", "azure_openai")
    if provider not in PROVIDERS:
        raise ValueError("azure_openai.provider must be azure_openai or claude_foundry")
    return provider


def get_client(config: dict, provider: str | None = None, *,
               max_retries: int | None = None, timeout: float | None = None):
    """Return (client, deployment); unavailable credentials preserve regex fallback.

    An explicit provider is used by comparison without modifying the live config.
    No SDK request is made during client construction.
    """
    provider = get_provider(config) if provider is None else provider
    if provider not in PROVIDERS:
        raise ValueError("Unknown LLM provider")
    options = {}
    if max_retries is not None:
        options["max_retries"] = max_retries
    if timeout is not None:
        options["timeout"] = timeout
    if provider == "azure_openai":
        az = config.get("azure_openai", {})
        endpoint = os.getenv("AZURE_OPENAI_ENDPOINT") or az.get("endpoint", "")
        api_key = os.getenv("AZURE_OPENAI_API_KEY") or az.get("api_key", "")
        deployment = az.get("deployment", "gpt-4o-mini")
        if not endpoint or not api_key or "YOUR_" in endpoint or "YOUR_" in api_key:
            return None, None
        try:
            from openai import AzureOpenAI
        except ImportError:
            print("  ⚠  openai package not installed. Run: pip3 install openai")
            return None, None
        return AzureOpenAI(
            azure_endpoint=endpoint,
            api_key=api_key,
            api_version=az.get("api_version", "2024-08-01-preview"),
            **options,
        ), deployment

    # Foundry credentials belong in the environment, never in the public config.
    resource = os.getenv("ANTHROPIC_FOUNDRY_RESOURCE", "")
    api_key = os.getenv("ANTHROPIC_FOUNDRY_API_KEY", "")
    deployment = config.get("claude_foundry", {}).get("deployment", "claude-haiku-4-5")
    if (not resource or not api_key or not deployment
            or resource.startswith("YOUR_") or api_key.startswith("YOUR_")
            or resource == "your-foundry-resource" or api_key == "your-key-here"):
        return None, None
    try:
        from anthropic import AnthropicFoundry
    except ImportError:
        print("  ⚠  anthropic package not installed. Run: pip3 install -r requirements.txt")
        return None, None
    return AnthropicFoundry(api_key=api_key, resource=resource, **options), deployment


def complete(client, provider: str, deployment: str, prompt: str, *,
             max_tokens: int, temperature: float) -> Completion:
    """Execute exactly one logical request and expose usage without logging content.

    Exceptions propagate to the caller so comparison can distinguish API errors
    from a successful NULL answer. SDK retries are controlled by the caller.
    """
    started = time.perf_counter()
    if provider == "azure_openai":
        response = client.chat.completions.create(
            model=deployment,
            messages=[{"role": "user", "content": prompt}],
            max_completion_tokens=max_tokens,
            temperature=temperature,
        )
        text = response.choices[0].message.content or ""
        usage = getattr(response, "usage", None)
        input_tokens = getattr(usage, "prompt_tokens", None)
        output_tokens = getattr(usage, "completion_tokens", None)
        details = getattr(usage, "prompt_tokens_details", None)
        cache_read = getattr(details, "cached_tokens", None)
        cache_creation = None
        finish_reason = getattr(response.choices[0], "finish_reason", None)
        response_model = getattr(response, "model", None)
    elif provider == "claude_foundry":
        message = client.messages.create(
            model=deployment,
            max_tokens=max_tokens,
            temperature=temperature,
            messages=[{"role": "user", "content": prompt}],
        )
        # Haiku 4.5 without thinking returns text first.
        # A different block type is an API/compatibility error, not a NULL answer.
        text = message.content[0].text
        usage = getattr(message, "usage", None)
        input_tokens = getattr(usage, "input_tokens", None)
        output_tokens = getattr(usage, "output_tokens", None)
        cache_read = getattr(usage, "cache_read_input_tokens", None)
        cache_creation = getattr(usage, "cache_creation_input_tokens", None)
        finish_reason = getattr(message, "stop_reason", None)
        response_model = getattr(message, "model", None)
    else:
        raise ValueError("Unknown LLM provider")
    return Completion(text, time.perf_counter() - started, input_tokens,
                      output_tokens, provider, deployment, cache_read,
                      cache_creation, finish_reason, response_model)


def parse_amount(raw: str | None) -> float | None:
    """Preserve the existing Azure numeric/NULL parsing behavior verbatim."""
    raw = (raw or "").strip()
    if not raw or raw.upper() == "NULL":
        return None
    raw = raw.replace("$", "").replace(",", "").strip().rstrip(".")
    m = re.search(r"-?\d+(?:\.\d+)?", raw)
    if not m:
        return None
    return float(m.group(0))


def extraction_prompt(body: str, label: str, snippet: str = "") -> str:
    # Strip <style> and <script> blocks first, they hog truncation budget
    # without containing any actual bill content.
    no_style = re.sub(r"<style\b[^>]*>.*?</style>", " ", body, flags=re.DOTALL | re.IGNORECASE)
    no_style = re.sub(r"<script\b[^>]*>.*?</script>", " ", no_style, flags=re.DOTALL | re.IGNORECASE)
    # Strip <head> entirely, it's all metadata, no bill content
    no_head = re.sub(r"<head\b[^>]*>.*?</head>", " ", no_style, flags=re.DOTALL | re.IGNORECASE)
    # Decode HTML entities before tag-stripping so &#36;251.60 → $251.60
    clean = html_lib.unescape(no_head)
    # Replace tags with space (not empty) so adjacent <td>$111.40</td> doesn't merge
    clean = re.sub(r"<[^>]+>", " ", clean)
    # Drop zero-width joiners and other invisible padding chars Quantum loves
    clean = re.sub(r"[​‌‍⁠͏]+", "", clean)
    clean = re.sub(r"\s+", " ", clean).strip()[:8000]
    snippet_clean = html_lib.unescape(snippet).strip() if snippet else ""

    # Gmail's snippet is server-generated and usually contains exactly the key
    # text (e.g., "balance due $251.60"). Putting it up front gives Azure a
    # high-signal anchor even when the body is mostly layout markup.
    body_block = (
        f"GMAIL SNIPPET (high-signal preview): {snippet_clean}\n\n" if snippet_clean else ""
    ) + f"FULL BODY TEXT:\n{clean}"

    return (
        f"You are reading a utility bill email for {label}.\n\n"
        f"Return the SINGLE NUMERIC AMOUNT the customer actually owes/paid for this billing period.\n\n"
        f"Rules:\n"
        f"- Pick the bill total, balance due, payment amount, or amount due, NOT retail price, NOT discount/savings, NOT $0.00 placeholders, NOT 'last month' references.\n"
        f"- If there are multiple amounts (retail, discount, total, etc.), pick the one labeled as 'total', 'amount due', 'balance due', 'amount billed', or 'amount charged'.\n"
        f"- Reply with ONLY the numeric value: e.g., 57.18 (no dollar sign, no commas, no explanation).\n"
        f"- If you genuinely cannot determine the amount, reply exactly: NULL\n\n"
        f"{body_block}"
    )


def commentary_prompt(analysis: dict, billing_month: date) -> str:
    # Build a compact, structured digest for the model
    lines = []
    lines.append(f"Billing month: {billing_month.strftime('%B %Y')}")
    lines.append(f"Tenant total this month: ${analysis['tenant_total_now']:.2f}")
    if analysis["tenant_total_prior"] is not None:
        lines.append(f"Tenant total prior month: ${analysis['tenant_total_prior']:.2f}")
    lines.append("")
    lines.append("Per-utility raw amounts and changes:")
    for u in analysis["per_utility"]:
        pct = f"{u['pct_change']:+.1f}%" if u["pct_change"] is not None else "new"
        z = f"z={u['zscore']:+.2f}" if u["zscore"] is not None else "z=n/a"
        lines.append(f"  - {u['label']}: ${u['current']:.2f} (prior ${u['prior']:.2f}, {pct}, {z})")
    if analysis["flags"]:
        lines.append("")
        lines.append("Auto-flagged anomalies:")
        for f in analysis["flags"]:
            lines.append(f"  - [{f['kind']}] {f['detail']}")

    digest = "\n".join(lines)
    season_hint = {
        12: "Winter, high heating gas use expected.",
        1: "Winter, high heating gas use expected.",
        2: "Winter, high heating gas use expected.",
        3: "End of heating season, gas should begin dropping.",
        4: "Spring, gas should drop sharply, water/electric stable.",
        5: "Spring, heating off, possible AC ramp.",
        6: "Summer, AC ramps up; gas minimal.",
        7: "Summer peak, high electricity expected for AC.",
        8: "Summer peak, high electricity expected for AC.",
        9: "Early fall, electricity drops, gas still off.",
        10: "Fall, gas may start rising.",
        11: "Fall/early winter, gas climbs.",
    }.get(billing_month.month, "")

    return f"""You are reviewing a landlord's monthly utility bills for a rental property.

Seasonal context: {season_hint}

Bill data:
{digest}

Write a SHORT (4-6 sentences) candid analysis. Cover:
1. Are these bills reasonable for the month? Call out anything that looks off.
2. Possible causes for any flagged anomalies (leak? appliance? weather? service change?).
3. Whether the landlord should act on anything before sending the invoice to tenants.

Use plain English. No bullet points. No fluff. Be direct."""
