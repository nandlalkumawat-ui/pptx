import json
import time
from typing import Any

import ollama

from config import OLLAMA_HOST, OLLAMA_KEEP_ALIVE, OLLAMA_MODEL, OLLAMA_NUM_CTX
from utils.perf import PERF


class OllamaClient:

    def __init__(
        self,
        model: str = OLLAMA_MODEL,
    ):
        self.model = model

        self.client = ollama.Client(
            host=OLLAMA_HOST,
        )

    def _options(self, extra: dict | None = None) -> dict:
        options = {"num_ctx": OLLAMA_NUM_CTX}
        options.update(extra or {})
        return options

    @staticmethod
    def _log_metrics(response) -> None:
        """Print Ollama's own breakdown: model load, prompt eval, generation."""

        ns = 1e9
        load = (response.get("load_duration") or 0) / ns
        prompt_eval = (response.get("prompt_eval_duration") or 0) / ns
        generation = (response.get("eval_duration") or 0) / ns
        prompt_tokens = response.get("prompt_eval_count") or 0
        output_tokens = response.get("eval_count") or 0

        PERF.add_time("ollama model load", load)
        PERF.add_time("ollama prompt processing", prompt_eval)
        PERF.add_time("ollama token generation", generation)
        PERF.incr("ollama prompt tokens", prompt_tokens)
        PERF.incr("ollama output tokens", output_tokens)

        rate = output_tokens / generation if generation else 0
        print(
            f"[timing] Ollama load {load:.2f}s | prompt {prompt_tokens} tok "
            f"in {prompt_eval:.2f}s | output {output_tokens} tok in "
            f"{generation:.2f}s ({rate:.1f} tok/s)"
        )
        if prompt_tokens >= OLLAMA_NUM_CTX - 16:
            print("[warning] Prompt filled the context window; it may have been truncated.")

    def warm_up(self, prefix: str = "") -> None:
        """
        Load the model into memory (no-op if it is already loaded).

        With ``prefix``, also evaluate that text once so Ollama keeps its
        KV cache; a later prompt starting with the same text skips
        re-processing it.
        """

        started = time.perf_counter()
        try:
            if prefix:
                self.client.chat(
                    model=self.model,
                    messages=[{"role": "user", "content": prefix}],
                    keep_alive=OLLAMA_KEEP_ALIVE,
                    options=self._options({"num_predict": 1}),
                )
            else:
                self.client.generate(
                    model=self.model,
                    prompt="",
                    keep_alive=OLLAMA_KEEP_ALIVE,
                    options=self._options(),
                )
            print(f"[timing] Ollama warm-up: {time.perf_counter() - started:.2f}s")
        except Exception as exc:
            print(f"Ollama warm-up failed (continuing): {exc}")

    def generate(
        self,
        prompt: str,
    ) -> str:

        print(
            f"\nGenerating with Ollama: {self.model}"
        )
        started = time.perf_counter()

        response = self.client.chat(
            model=self.model,
            messages=[
                {
                    "role": "user",
                    "content": prompt,
                }
            ],
            keep_alive=OLLAMA_KEEP_ALIVE,
            options=self._options(),
        )

        elapsed = time.perf_counter() - started
        print(f"[timing] Ollama inference: {elapsed:.2f}s")
        self._log_metrics(response)
        return response["message"]["content"]

    def generate_json(
        self,
        prompt: str,
        schema: dict[str, Any] | None = None,
        options: dict | None = None,
    ):
        """
        Generate JSON. When ``schema`` is given, Ollama constrains decoding
        to that JSON schema (structured output), otherwise plain JSON mode.
        """

        print(
            f"\nGenerating JSON with Ollama: "
            f"{self.model}"
        )
        started = time.perf_counter()

        response = self.client.chat(
            model=self.model,
            messages=[
                {
                    "role": "user",
                    "content": prompt,
                }
            ],
            format=schema or "json",
            keep_alive=OLLAMA_KEEP_ALIVE,
            options=self._options(options),
        )

        elapsed = time.perf_counter() - started
        print(f"[timing] Ollama JSON inference: {elapsed:.2f}s")
        self._log_metrics(response)

        content = response["message"]["content"]

        return json.loads(content)
