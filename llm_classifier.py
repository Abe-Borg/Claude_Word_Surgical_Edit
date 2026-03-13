"""
LLM classifier module for automated CSI paragraph classification.

Calls the Anthropic API with the master prompt + slim bundle to produce
classification instructions (same schema as instructions.json).

Design constraint: pure module with no CLI — imported by gui.py.
"""
from __future__ import annotations

import json
import re
import time
from typing import Any, Dict, List, Optional, Tuple


def estimate_tokens(text: str) -> int:
    """Rough token estimate (1 token ≈ 4 chars)."""
    return len(text) // 4


def _strip_code_fences(text: str) -> str:
    """Remove markdown code fences if present."""
    text = text.strip()
    if text.startswith("```"):
        # Remove opening fence (```json or ```)
        text = re.sub(r"^```\w*\s*\n?", "", text)
        # Remove closing fence
        text = re.sub(r"\n?```\s*$", "", text)
    return text.strip()


def _call_api(
    client: Any,
    system: str,
    user_message: str,
    model: str,
    max_tokens: int = 128000,
) -> str:
    """Single API call with retry logic. Returns raw response text."""
    import anthropic

    last_error: Optional[Exception] = None
    for attempt in range(3):  # initial + 2 retries
        try:
            with client.messages.stream(
                model=model,
                max_tokens=max_tokens,
                temperature=1,
                thinking={"type": "adaptive"},
                output_config={"effort":"max"},
                system=system,
                messages=[{"role": "user", "content": user_message}],
            ) as stream:
                return stream.get_final_text()
        except (anthropic.APIError, anthropic.APIConnectionError, anthropic.RateLimitError) as e:
            last_error = e
            if attempt < 2:
                time.sleep(2 ** (attempt + 1))
    raise last_error  # type: ignore[misc]


def _parse_response(raw: str) -> dict:
    """Parse JSON from LLM response, handling code fences."""
    cleaned = _strip_code_fences(raw)
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError as e:
        raise ValueError(
            f"LLM response is not valid JSON: {e}\n\nRaw response (first 2000 chars):\n{raw[:2000]}"
        ) from e


def classify_document(
    slim_bundle: dict,
    master_prompt: str,
    run_instruction: str,
    api_key: str,
    model: str = "claude-opus-4-6",
    min_coverage: float = 0.90,
) -> dict:
    """
    Classify all paragraphs in a slim bundle using the Anthropic API.

    Args:
        slim_bundle: The slim bundle dict from build_slim_bundle().
        master_prompt: Content of master_prompt.txt (system prompt).
        run_instruction: Content of run_instruction_prompt.txt (task prompt).
        api_key: Anthropic API key.
        model: Model ID to use.
        min_coverage: Minimum coverage fraction (0.0–1.0). Raises ValueError
            if the fraction of classifiable paragraphs with style assignments
            falls below this threshold. Default 0.90.

    Returns:
        Parsed instructions dict (same schema as instructions.json).

    Raises:
        ValueError: If the LLM response is not valid JSON, fails validation,
            or coverage is below min_coverage.
    """
    import anthropic
    from docx_decomposer import validate_instructions

    client = anthropic.Anthropic(api_key=api_key)

    bundle_json = json.dumps(slim_bundle, indent=2)
    user_message = f"{run_instruction}\n\nSlim bundle:\n{bundle_json}"

    # Check if chunking is needed
    total_text = master_prompt + user_message
    token_est = estimate_tokens(total_text)

    n_paragraphs = len(slim_bundle.get("paragraphs", []))
    if token_est > 80_000 or n_paragraphs > 300:
        instructions = _classify_chunked(
            slim_bundle, master_prompt, run_instruction, client, model
        )
    else:
        raw = _call_api(client, master_prompt, user_message, model)
        instructions = _parse_response(raw)

    validate_instructions(instructions, slim_bundle=slim_bundle)

    # Coverage enforcement
    coverage, styled, classifiable = compute_coverage(slim_bundle, instructions)
    if coverage < min_coverage:
        raise ValueError(
            f"Coverage {coverage:.1%} ({styled}/{classifiable}) is below the "
            f"required minimum of {min_coverage:.0%}. The LLM did not classify "
            f"enough paragraphs."
        )

    return instructions


def _classify_chunked(
    slim_bundle: dict,
    master_prompt: str,
    run_instruction: str,
    client: Any,
    model: str,
    chunk_size: int = 200,
    overlap: int = 20,
) -> dict:
    """
    Chunked classification for large documents.

    First chunk returns full instructions (create_styles + roles + apply_pStyle).
    Subsequent chunks receive the already-determined styles/roles as context
    and return only apply_pStyle for their paragraph range.
    """
    from docx_decomposer import validate_instructions

    paragraphs = slim_bundle.get("paragraphs", [])
    total = len(paragraphs)

    # Build chunk boundaries
    chunks: List[tuple] = []  # (start, end) indices into paragraphs list
    start = 0
    while start < total:
        end = min(start + chunk_size, total)
        chunks.append((start, end))
        start = end - overlap if end < total else end

    # First chunk: full classification
    first_bundle = dict(slim_bundle)
    first_bundle["paragraphs"] = paragraphs[chunks[0][0]:chunks[0][1]]
    bundle_json = json.dumps(first_bundle, indent=2)
    user_msg = f"{run_instruction}\n\nSlim bundle:\n{bundle_json}"

    raw = _call_api(client, master_prompt, user_msg, model)
    merged = _parse_response(raw)
    validate_instructions(merged)

    if len(chunks) <= 1:
        return merged

    # Subsequent chunks: only apply_pStyle
    context_info = json.dumps({
        "create_styles": merged.get("create_styles", []),
        "roles": merged.get("roles", {}),
    }, indent=2)

    # Track styleId per paragraph_index for collision detection
    idx_to_style: Dict[int, str] = {
        item["paragraph_index"]: item["styleId"]
        for item in merged.get("apply_pStyle", [])
    }

    for chunk_start, chunk_end in chunks[1:]:
        chunk_bundle = dict(slim_bundle)
        chunk_bundle["paragraphs"] = paragraphs[chunk_start:chunk_end]
        chunk_json = json.dumps(chunk_bundle, indent=2)

        chunk_prompt = (
            f"{run_instruction}\n\n"
            f"The following styles and roles have already been determined:\n{context_info}\n\n"
            f"You MUST use these exact styles. Return ONLY the apply_pStyle array for the "
            f"paragraphs in this chunk (paragraph indices {chunk_start} to {chunk_end - 1}).\n"
            f"Output format: {{\"apply_pStyle\": [...]}}\n\n"
            f"Slim bundle chunk:\n{chunk_json}"
        )

        raw = _call_api(client, master_prompt, chunk_prompt, model)
        chunk_result = _parse_response(raw)
        chunk_apply = chunk_result.get("apply_pStyle", [])

        # Merge with collision detection in overlap regions
        for item in chunk_apply:
            idx = item["paragraph_index"]
            sid = item["styleId"]
            if idx in idx_to_style:
                if idx_to_style[idx] != sid:
                    raise ValueError(
                        f"Chunk merge collision at paragraph {idx}: "
                        f"previous chunk assigned '{idx_to_style[idx]}', "
                        f"current chunk assigned '{sid}'"
                    )
                # Same assignment in overlap — skip duplicate
            else:
                merged.setdefault("apply_pStyle", []).append(item)
                idx_to_style[idx] = sid

    # Sort apply_pStyle by paragraph_index
    merged["apply_pStyle"] = sorted(
        merged.get("apply_pStyle", []),
        key=lambda x: x["paragraph_index"],
    )

    # Full validation with slim_bundle for coverage + range checks
    validate_instructions(merged, slim_bundle=slim_bundle)
    return merged


def compute_coverage(slim_bundle: dict, instructions: dict) -> Tuple[float, int, int]:
    """
    Compute what percentage of classifiable paragraphs received a style.

    Returns:
        (coverage_fraction, styled_count, classifiable_count)
    """
    paragraphs = slim_bundle.get("paragraphs", [])
    classifiable = 0
    classifiable_indices = set()

    for p in paragraphs:
        text = (p.get("text") or "").strip()
        if not text:
            continue
        if p.get("contains_sectPr", False):
            continue
        if text.upper() == "END OF SECTION":
            continue
        # Editor/specifier notes in brackets
        if text.startswith("[") and text.endswith("]"):
            continue
        classifiable += 1
        classifiable_indices.add(p["paragraph_index"])

    styled_indices = {item["paragraph_index"] for item in instructions.get("apply_pStyle", [])}
    styled_count = len(styled_indices & classifiable_indices)

    coverage = styled_count / classifiable if classifiable > 0 else 1.0
    return coverage, styled_count, classifiable
