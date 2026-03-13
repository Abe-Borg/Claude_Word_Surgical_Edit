"""Unit tests for llm_classifier — pure functions that don't require API calls."""

from __future__ import annotations

import pytest

from llm_classifier import _strip_code_fences, _parse_response, compute_coverage


# ---------------------------------------------------------------------------
# _strip_code_fences
# ---------------------------------------------------------------------------

class TestStripCodeFences:

    def test_no_fences(self):
        assert _strip_code_fences('{"key": "value"}') == '{"key": "value"}'

    def test_json_fences(self):
        raw = '```json\n{"key": "value"}\n```'
        assert _strip_code_fences(raw) == '{"key": "value"}'

    def test_bare_fences(self):
        raw = '```\n{"key": "value"}\n```'
        assert _strip_code_fences(raw) == '{"key": "value"}'

    def test_whitespace_around_fences(self):
        raw = '  ```json\n{"key": "value"}\n```  '
        assert _strip_code_fences(raw) == '{"key": "value"}'

    def test_no_fences_preserved(self):
        raw = '  {"key": "value"}  '
        assert _strip_code_fences(raw) == '{"key": "value"}'


# ---------------------------------------------------------------------------
# _parse_response
# ---------------------------------------------------------------------------

class TestParseResponse:

    def test_valid_json(self):
        result = _parse_response('{"create_styles": []}')
        assert result == {"create_styles": []}

    def test_json_with_code_fences(self):
        result = _parse_response('```json\n{"create_styles": []}\n```')
        assert result == {"create_styles": []}

    def test_invalid_json_raises(self):
        with pytest.raises(ValueError, match="not valid JSON"):
            _parse_response("this is not json at all")

    def test_empty_string_raises(self):
        with pytest.raises(ValueError, match="not valid JSON"):
            _parse_response("")

    def test_truncated_json_raises(self):
        with pytest.raises(ValueError, match="not valid JSON"):
            _parse_response('{"create_styles": [')


# ---------------------------------------------------------------------------
# compute_coverage
# ---------------------------------------------------------------------------

class TestComputeCoverage:

    def _bundle(self, paragraphs):
        return {"paragraphs": paragraphs}

    def _para(self, index, text, sectPr=False):
        return {"paragraph_index": index, "text": text, "contains_sectPr": sectPr}

    def test_full_coverage(self):
        bundle = self._bundle([
            self._para(0, "Content A"),
            self._para(1, "Content B"),
        ])
        instructions = {"apply_pStyle": [
            {"paragraph_index": 0, "styleId": "s1"},
            {"paragraph_index": 1, "styleId": "s2"},
        ]}
        cov, styled, classifiable = compute_coverage(bundle, instructions)
        assert cov == 1.0
        assert styled == 2
        assert classifiable == 2

    def test_partial_coverage(self):
        bundle = self._bundle([
            self._para(0, "Content A"),
            self._para(1, "Content B"),
        ])
        instructions = {"apply_pStyle": [
            {"paragraph_index": 0, "styleId": "s1"},
        ]}
        cov, styled, classifiable = compute_coverage(bundle, instructions)
        assert cov == 0.5
        assert styled == 1
        assert classifiable == 2

    def test_skip_empty(self):
        bundle = self._bundle([
            self._para(0, "Content"),
            self._para(1, ""),
        ])
        instructions = {"apply_pStyle": [
            {"paragraph_index": 0, "styleId": "s1"},
        ]}
        cov, styled, classifiable = compute_coverage(bundle, instructions)
        assert cov == 1.0
        assert classifiable == 1

    def test_skip_sectpr(self):
        bundle = self._bundle([
            self._para(0, "Content"),
            self._para(1, "Section break", sectPr=True),
        ])
        instructions = {"apply_pStyle": [
            {"paragraph_index": 0, "styleId": "s1"},
        ]}
        cov, styled, classifiable = compute_coverage(bundle, instructions)
        assert cov == 1.0

    def test_skip_end_of_section(self):
        bundle = self._bundle([
            self._para(0, "Content"),
            self._para(1, "END OF SECTION"),
        ])
        instructions = {"apply_pStyle": [
            {"paragraph_index": 0, "styleId": "s1"},
        ]}
        cov, styled, classifiable = compute_coverage(bundle, instructions)
        assert cov == 1.0

    def test_skip_editor_notes(self):
        bundle = self._bundle([
            self._para(0, "Content"),
            self._para(1, "[Editor note: remove before final]"),
        ])
        instructions = {"apply_pStyle": [
            {"paragraph_index": 0, "styleId": "s1"},
        ]}
        cov, styled, classifiable = compute_coverage(bundle, instructions)
        assert cov == 1.0

    def test_empty_document(self):
        bundle = self._bundle([])
        instructions = {"apply_pStyle": []}
        cov, styled, classifiable = compute_coverage(bundle, instructions)
        assert cov == 1.0
        assert classifiable == 0
