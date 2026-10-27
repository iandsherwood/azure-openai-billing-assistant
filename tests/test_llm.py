"""Offline regressions for provider parity and billing fallback boundaries."""

import base64
from contextlib import redirect_stdout
from datetime import date
import hashlib
import importlib.util
import io
import json
from pathlib import Path
import sys
from types import SimpleNamespace as NS
import unittest
from unittest.mock import Mock, patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from billing_automation import llm, run_billing


def azure_response(text="123.45", usage=True):
    return NS(choices=[NS(message=NS(content=text), finish_reason="stop")],
              usage=NS(prompt_tokens=100, completion_tokens=5,
                       prompt_tokens_details=NS(cached_tokens=0)) if usage else None,
              model="azure-model-version")


def claude_response(text="123.45"):
    return NS(content=[NS(text=text)],
              usage=NS(input_tokens=110, output_tokens=5,
                       cache_read_input_tokens=0, cache_creation_input_tokens=0),
              stop_reason="end_turn", model="claude-model-version")


class PromptRegressionTests(unittest.TestCase):
    def test_extraction_prompt_matches_pre_change_capture(self):
        # Hash captured from the pre-change function with a fake Azure client.
        prompt = llm.extraction_prompt(
            "<head>ignored</head><style>ignored</style><p>Balance &#36;123.45</p>",
            "Electricity", "Pay $123.45")
        self.assertEqual(hashlib.sha256(prompt.encode()).hexdigest(),
                         "b55140e35818d39afe5b2064b7bb3f7571e7656675b169ad416c9c3b5054ba25")

    def test_commentary_prompt_matches_pre_change_capture(self):
        analysis = {
            "tenant_total_now": 100, "tenant_total_prior": 90,
            "per_utility": [{"label": "Electricity", "current": 40, "prior": 30,
                             "pct_change": 33.3333, "zscore": 2.1}],
            "flags": [{"kind": "MoM", "detail": "Electricity changed"}],
        }
        prompt = llm.commentary_prompt(analysis, date(2026, 4, 1))
        self.assertEqual(hashlib.sha256(prompt.encode()).hexdigest(),
                         "b573edbc4cc27f0c07f1ed638953c1c7030eff370275fa91ba02d0cf27e4f567")

    def test_legacy_null_and_numeric_parsing(self):
        cases = [(None, None), ("", None), (" NULL ", None), ("null", None),
                 ("no amount", None), ("$1,234.56.", 1234.56),
                 ("**57.18**", 57.18), ("Amount: -12.10 credit", -12.10),
                 ("0", 0.0)]
        for raw, expected in cases:
            with self.subTest(raw=raw):
                self.assertEqual(llm.parse_amount(raw), expected)


class ProviderTests(unittest.TestCase):
    def test_default_and_invalid_provider(self):
        self.assertEqual(llm.get_provider({}), "azure_openai")
        with self.assertRaises(ValueError):
            llm.get_provider({"azure_openai": {"provider": "typo"}})

    @patch.dict("os.environ", {}, clear=True)
    def test_azure_constructor_is_unchanged_by_default(self):
        constructor = Mock(return_value=object())
        config = {"azure_openai": {"endpoint": "https://example.openai.azure.com/",
                                   "api_key": "synthetic-key", "deployment": "baseline",
                                   "api_version": "2024-12-01-preview"}}
        with patch.dict(sys.modules, {"openai": NS(AzureOpenAI=constructor)}):
            client, deployment = llm.get_client(config)
        self.assertIs(client, constructor.return_value)
        self.assertEqual(deployment, "baseline")
        constructor.assert_called_once_with(
            azure_endpoint="https://example.openai.azure.com/", api_key="synthetic-key",
            api_version="2024-12-01-preview")

    @patch.dict("os.environ", {"AZURE_OPENAI_ENDPOINT": "https://env.openai.azure.com/",
                               "AZURE_OPENAI_API_KEY": "synthetic-env-key"}, clear=True)
    def test_azure_env_precedence_and_comparison_retry_control(self):
        constructor = Mock()
        with patch.dict(sys.modules, {"openai": NS(AzureOpenAI=constructor)}):
            llm.get_client({"azure_openai": {"endpoint": "https://config.invalid",
                                            "api_key": "synthetic-config-key"}},
                           max_retries=0, timeout=30)
        self.assertEqual(constructor.call_args.kwargs["azure_endpoint"],
                         "https://env.openai.azure.com/")
        self.assertEqual(constructor.call_args.kwargs["max_retries"], 0)
        self.assertEqual(constructor.call_args.kwargs["timeout"], 30)

    @patch.dict("os.environ", {"ANTHROPIC_FOUNDRY_RESOURCE": "synthetic-resource",
                               "ANTHROPIC_FOUNDRY_API_KEY": "synthetic-key"}, clear=True)
    def test_foundry_constructor_uses_environment_and_named_deployment(self):
        constructor = Mock()
        config = {"azure_openai": {"provider": "claude_foundry"},
                  "claude_foundry": {"deployment": "approved-haiku-deployment",
                                     "api_key": "ignored-config-value"}}
        with patch.dict(sys.modules, {"anthropic": NS(AnthropicFoundry=constructor)}):
            _, deployment = llm.get_client(config, max_retries=0)
        self.assertEqual(deployment, "approved-haiku-deployment")
        constructor.assert_called_once_with(api_key="synthetic-key",
                                            resource="synthetic-resource", max_retries=0)

    @patch.dict("os.environ", {}, clear=True)
    def test_missing_credentials_do_not_construct_clients(self):
        self.assertEqual(llm.get_client({}), (None, None))
        self.assertEqual(llm.get_client({"azure_openai": {"provider": "claude_foundry"},
                                        "claude_foundry": {"api_key": "ignored"}}),
                         (None, None))

    def test_both_providers_receive_same_prompt_limits_and_temperatures(self):
        azure = Mock()
        azure.chat.completions.create.return_value = azure_response()
        claude = Mock()
        claude.messages.create.return_value = claude_response()
        for tokens, temperature in [(50, 0), (400, 0.3)]:
            with self.subTest(tokens=tokens):
                a = llm.complete(azure, "azure_openai", "azure-deployment", "frozen prompt",
                                 max_tokens=tokens, temperature=temperature)
                b = llm.complete(claude, "claude_foundry", "claude-deployment", "frozen prompt",
                                 max_tokens=tokens, temperature=temperature)
                az = azure.chat.completions.create.call_args.kwargs
                cl = claude.messages.create.call_args.kwargs
                self.assertEqual(az["messages"], cl["messages"])
                self.assertEqual(az["max_completion_tokens"], cl["max_tokens"])
                self.assertEqual(az["temperature"], cl["temperature"])
                self.assertEqual((a.text, b.text), ("123.45", "123.45"))
                self.assertEqual((a.input_tokens, b.input_tokens), (100, 110))
                self.assertEqual((a.output_tokens, b.output_tokens), (5, 5))
                self.assertGreaterEqual(a.latency_seconds, 0)

    def test_missing_usage_stays_unknown_and_api_errors_propagate(self):
        client = Mock()
        client.chat.completions.create.return_value = azure_response(usage=False)
        result = llm.complete(client, "azure_openai", "sample", "prompt",
                              max_tokens=50, temperature=0)
        self.assertIsNone(result.input_tokens)
        self.assertIsNone(result.output_tokens)
        client.chat.completions.create.side_effect = RuntimeError("synthetic response body")
        with self.assertRaises(RuntimeError):
            llm.complete(client, "azure_openai", "sample", "prompt",
                         max_tokens=50, temperature=0)


class BillingIntegrationTests(unittest.TestCase):
    def fake_service(self):
        service = Mock()
        messages = service.users.return_value.messages.return_value
        messages.list.return_value.execute.return_value = {"messages": [{"id": "synthetic"}]}
        messages.get.return_value.execute.return_value = {
            "snippet": "Amount due $12.34",
            "payload": {"mimeType": "text/plain", "body": {
                "data": base64.urlsafe_b64encode(b"Due $12.34").decode()},
                "headers": [{"name": "Subject", "value": "Synthetic utility bill"}]},
        }
        return service

    def config(self, provider="azure_openai"):
        return {"azure_openai": {"provider": provider},
                "gmail_searches": {"electricity": {"sender": "test@example.invalid"}},
                "utility_labels": {"electricity": "Electricity"}}

    def test_no_client_uses_existing_regex_fallback_for_both_providers(self):
        for provider in llm.PROVIDERS:
            with self.subTest(provider=provider), \
                 patch.object(run_billing, "_gmail_service", return_value=self.fake_service()), \
                 patch.object(run_billing, "_get_llm_client", return_value=(None, None)), \
                 redirect_stdout(io.StringIO()):
                bills = run_billing.read_bills_from_gmail(date(2026, 4, 1), self.config(provider))
                self.assertEqual(bills, {"electricity": 12.34})

    def test_successful_null_does_not_get_concealed_by_regex(self):
        for provider in llm.PROVIDERS:
            with self.subTest(provider=provider), \
                 patch.object(run_billing, "_gmail_service", return_value=self.fake_service()), \
                 patch.object(run_billing, "_get_llm_client", return_value=(object(), "sample")), \
                 patch.object(llm, "complete", return_value=llm.Completion("NULL", 0.1, 10, 1, provider, "sample")), \
                 patch.object(run_billing, "_extract_amount_regex") as fallback, \
                 redirect_stdout(io.StringIO()):
                bills = run_billing.read_bills_from_gmail(date(2026, 4, 1), self.config(provider))
                self.assertEqual(bills, {"electricity": None})
                fallback.assert_not_called()

    def test_commentary_switch_preserves_shared_settings(self):
        analysis = {"tenant_total_now": 100, "tenant_total_prior": None,
                    "per_utility": [], "flags": []}
        for provider in llm.PROVIDERS:
            with self.subTest(provider=provider), \
                 patch.object(run_billing, "_get_llm_client", return_value=(object(), "sample")), \
                 patch.object(llm, "complete", return_value=llm.Completion(" Narrative. ", 0.1, 10, 2, provider, "sample")) as call:
                text = run_billing.azure_commentary(analysis, date(2026, 4, 1), self.config(provider))
                self.assertEqual(text, "Narrative.")
                self.assertEqual(call.call_args.args[1], provider)
                self.assertEqual(call.call_args.kwargs, {"max_tokens": 400, "temperature": 0.3})

    def test_error_diagnostics_never_echo_sdk_response_body(self):
        with patch.object(llm, "complete", side_effect=RuntimeError("private synthetic body")):
            output = io.StringIO()
            with redirect_stdout(output):
                result = run_billing._extract_amount_azure("body", "Utility", object(), "sample")
            self.assertIsNone(result)
            self.assertIn("RuntimeError", output.getvalue())
            self.assertNotIn("private synthetic body", output.getvalue())


@unittest.skipUnless(importlib.util.find_spec("anthropic"), "Install requirements for SDK wire test")
class FoundrySDKTests(unittest.TestCase):
    @patch.dict("os.environ", {}, clear=True)
    def test_pinned_sdk_wire_request_with_no_network(self):
        import httpx
        import anthropic

        requests = []

        def respond(request):
            requests.append(request)
            return httpx.Response(200, json={
                "id": "synthetic", "type": "message", "role": "assistant",
                "model": "synthetic-haiku", "content": [{"type": "text", "text": "12.34"}],
                "stop_reason": "end_turn", "stop_sequence": None,
                "usage": {"input_tokens": 15, "output_tokens": 4},
            }, request=request)

        self.assertEqual(anthropic.__version__, "0.125.0")
        with httpx.Client(transport=httpx.MockTransport(respond)) as transport:
            with anthropic.AnthropicFoundry(
                api_key="synthetic-key", resource="synthetic-resource",
                http_client=transport, max_retries=0,
            ) as client:
                result = llm.complete(client, "claude_foundry", "test-deployment",
                                      "Synthetic utility amount.", max_tokens=50, temperature=0)
        self.assertEqual(len(requests), 1)
        self.assertEqual(str(requests[0].url),
                         "https://synthetic-resource.services.ai.azure.com/anthropic/v1/messages")
        self.assertEqual(requests[0].headers["anthropic-version"], "2023-06-01")
        payload = json.loads(requests[0].content)
        self.assertEqual(payload["model"], "test-deployment")
        self.assertEqual(payload["temperature"], 0)
        self.assertEqual(payload["max_tokens"], 50)
        self.assertEqual((result.text, result.input_tokens, result.output_tokens), ("12.34", 15, 4))


if __name__ == "__main__":
    unittest.main()
