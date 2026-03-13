"""Contract tests for Phase 1 pipeline — validates the hardened instruction
validation, prompt packaging, and schema integrity."""

from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any, Dict, List

import pytest

from docx_decomposer import validate_instructions

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

REPO_ROOT = Path(__file__).resolve().parent.parent


def _slim_bundle(paragraphs: List[Dict[str, Any]], style_catalog: List[Dict[str, Any]] | None = None) -> Dict[str, Any]:
    """Build a minimal slim_bundle for testing."""
    bundle: Dict[str, Any] = {"paragraphs": paragraphs}
    if style_catalog is not None:
        bundle["style_catalog"] = style_catalog
    return bundle


def _make_paragraph(index: int, text: str, contains_sectPr: bool = False) -> Dict[str, Any]:
    return {"paragraph_index": index, "text": text, "contains_sectPr": contains_sectPr}


def _valid_instructions(apply_indices: List[int] | None = None) -> Dict[str, Any]:
    """Return valid instructions covering paragraph indices 0-6 by default."""
    if apply_indices is None:
        apply_indices = [0, 1, 2, 3, 4, 5, 6]
    return {
        "create_styles": [
            {"styleId": "CSI_SectionTitle__ARCH", "name": "Section Title", "type": "paragraph", "derive_from_paragraph_index": 0},
            {"styleId": "CSI_Part__ARCH", "name": "Part", "type": "paragraph", "derive_from_paragraph_index": 1},
            {"styleId": "CSI_Article__ARCH", "name": "Article", "type": "paragraph", "derive_from_paragraph_index": 2},
            {"styleId": "CSI_Paragraph__ARCH", "name": "Paragraph", "type": "paragraph", "derive_from_paragraph_index": 3},
            {"styleId": "CSI_Subparagraph__ARCH", "name": "Subparagraph", "type": "paragraph", "derive_from_paragraph_index": 4},
            {"styleId": "CSI_Subsubparagraph__ARCH", "name": "Subsubparagraph", "type": "paragraph", "derive_from_paragraph_index": 6},
        ],
        "apply_pStyle": [
            {"paragraph_index": i, "styleId": "CSI_SectionTitle__ARCH"} if i == 0
            else {"paragraph_index": i, "styleId": "CSI_Part__ARCH"} if i == 1
            else {"paragraph_index": i, "styleId": "CSI_Article__ARCH"} if i == 2
            else {"paragraph_index": i, "styleId": "CSI_Paragraph__ARCH"} if i == 3
            else {"paragraph_index": i, "styleId": "CSI_Subparagraph__ARCH"} if i == 4
            else {"paragraph_index": i, "styleId": "CSI_Subparagraph__ARCH"} if i == 5
            else {"paragraph_index": i, "styleId": "CSI_Subsubparagraph__ARCH"}
            for i in apply_indices
        ],
        "roles": {
            "SectionTitle": {"styleId": "CSI_SectionTitle__ARCH", "exemplar_paragraph_index": 0},
            "PART": {"styleId": "CSI_Part__ARCH", "exemplar_paragraph_index": 1},
            "ARTICLE": {"styleId": "CSI_Article__ARCH", "exemplar_paragraph_index": 2},
            "PARAGRAPH": {"styleId": "CSI_Paragraph__ARCH", "exemplar_paragraph_index": 3},
            "SUBPARAGRAPH": {"styleId": "CSI_Subparagraph__ARCH", "exemplar_paragraph_index": 4},
            "SUBSUBPARAGRAPH": {"styleId": "CSI_Subsubparagraph__ARCH", "exemplar_paragraph_index": 6},
        },
        "notes": ["test"],
    }


def _default_paragraphs() -> List[Dict[str, Any]]:
    """7 classifiable paragraphs + 2 skippable ones (empty + END OF SECTION)."""
    return [
        _make_paragraph(0, "SECTION TITLE"),
        _make_paragraph(1, "PART 1 - GENERAL"),
        _make_paragraph(2, "1.01 SUMMARY"),
        _make_paragraph(3, "A. Section includes"),
        _make_paragraph(4, "1. Supply ductwork"),
        _make_paragraph(5, "2. Return ductwork"),
        _make_paragraph(6, "a. Per specification"),
        _make_paragraph(7, ""),              # empty — skip
        _make_paragraph(8, "END OF SECTION"),  # skip
    ]


# ---------------------------------------------------------------------------
# Shape-only validation (no slim_bundle)
# ---------------------------------------------------------------------------

class TestValidateInstructionsShapeOnly:

    def test_valid_passes(self):
        validate_instructions(_valid_instructions())

    def test_duplicate_paragraph_indices(self):
        instr = _valid_instructions()
        instr["apply_pStyle"].append({"paragraph_index": 0, "styleId": "CSI_Part__ARCH"})
        with pytest.raises(ValueError, match="Duplicate paragraph_index"):
            validate_instructions(instr)

    def test_invalid_role_name(self):
        instr = _valid_instructions()
        instr["roles"]["BOGUS"] = {"styleId": "CSI_Part__ARCH", "exemplar_paragraph_index": 1}
        with pytest.raises(ValueError, match="Unknown role.*BOGUS"):
            validate_instructions(instr)

    def test_exemplar_mismatch(self):
        instr = _valid_instructions()
        instr["roles"]["PART"]["exemplar_paragraph_index"] = 99
        with pytest.raises(ValueError, match="must equal derive_from_paragraph_index"):
            validate_instructions(instr)

    def test_forbidden_formatting_fields(self):
        instr = _valid_instructions()
        instr["create_styles"][0]["pPr"] = "<w:pPr/>"
        with pytest.raises(ValueError, match="LLM formatting fields are forbidden"):
            validate_instructions(instr)

    def test_extra_top_level_keys(self):
        instr = _valid_instructions()
        instr["extra_key"] = "bad"
        with pytest.raises(ValueError, match="Invalid instruction keys"):
            validate_instructions(instr)

    def test_missing_roles(self):
        instr = _valid_instructions()
        del instr["roles"]
        with pytest.raises(ValueError, match="Missing/invalid required key: roles"):
            validate_instructions(instr)

    def test_disallowed_style_id(self):
        instr = _valid_instructions()
        instr["create_styles"].append({
            "styleId": "CSI_Custom__ARCH",
            "name": "Custom",
            "type": "paragraph",
            "derive_from_paragraph_index": 0,
        })
        with pytest.raises(ValueError, match="styleId is not allowed"):
            validate_instructions(instr)


# ---------------------------------------------------------------------------
# Document-aware validation (with slim_bundle)
# ---------------------------------------------------------------------------

class TestValidateInstructionsWithBundle:

    def test_valid_with_bundle(self):
        bundle = _slim_bundle(_default_paragraphs())
        validate_instructions(_valid_instructions(), slim_bundle=bundle)

    def test_out_of_range_paragraph_index(self):
        bundle = _slim_bundle(_default_paragraphs())  # 9 paragraphs (0-8)
        instr = _valid_instructions()
        instr["apply_pStyle"].append({"paragraph_index": 100, "styleId": "CSI_Part__ARCH"})
        with pytest.raises(ValueError, match="out of range"):
            validate_instructions(instr, slim_bundle=bundle)

    def test_out_of_range_derive_from(self):
        bundle = _slim_bundle(_default_paragraphs())
        instr = _valid_instructions()
        instr["create_styles"][0]["derive_from_paragraph_index"] = 100
        instr["roles"]["SectionTitle"]["exemplar_paragraph_index"] = 100
        with pytest.raises(ValueError, match="out of range"):
            validate_instructions(instr, slim_bundle=bundle)

    def test_missing_coverage(self):
        """Instructions missing a classifiable paragraph should fail."""
        bundle = _slim_bundle(_default_paragraphs())
        # Remove paragraph 5 from apply_pStyle
        instr = _valid_instructions(apply_indices=[0, 1, 2, 3, 4, 6])
        with pytest.raises(ValueError, match="Incomplete apply_pStyle coverage"):
            validate_instructions(instr, slim_bundle=bundle)

    def test_exemplar_is_empty(self):
        paragraphs = _default_paragraphs()
        # Make paragraph 0 empty
        paragraphs[0]["text"] = ""
        bundle = _slim_bundle(paragraphs)
        instr = _valid_instructions()
        with pytest.raises(ValueError, match="exemplar paragraph 0 is empty"):
            validate_instructions(instr, slim_bundle=bundle)

    def test_exemplar_is_sectpr(self):
        paragraphs = _default_paragraphs()
        paragraphs[0]["contains_sectPr"] = True
        bundle = _slim_bundle(paragraphs)
        instr = _valid_instructions()
        with pytest.raises(ValueError, match="exemplar paragraph 0 contains sectPr"):
            validate_instructions(instr, slim_bundle=bundle)

    def test_exemplar_is_editor_note(self):
        paragraphs = _default_paragraphs()
        paragraphs[0]["text"] = "[Editor note: placeholder]"
        bundle = _slim_bundle(paragraphs)
        instr = _valid_instructions()
        with pytest.raises(ValueError, match="exemplar paragraph 0 is an editor note"):
            validate_instructions(instr, slim_bundle=bundle)

    def test_style_reference_consistency(self):
        """apply_pStyle references a style not in create_styles or catalog."""
        bundle = _slim_bundle(_default_paragraphs())
        instr = _valid_instructions()
        instr["apply_pStyle"][0]["styleId"] = "NonexistentStyle"
        with pytest.raises(ValueError, match="neither.*create_styles.*style catalog"):
            validate_instructions(instr, slim_bundle=bundle)

    def test_role_style_reference_consistency(self):
        """Role references a style not in create_styles or catalog."""
        bundle = _slim_bundle(_default_paragraphs())
        instr = _valid_instructions()
        # Remove the corresponding create_styles entry and change role styleId
        instr["create_styles"] = [s for s in instr["create_styles"] if s["styleId"] != "CSI_SectionTitle__ARCH"]
        instr["roles"]["SectionTitle"]["styleId"] = "NonexistentRole"
        with pytest.raises(ValueError, match="neither.*create_styles.*style catalog"):
            validate_instructions(instr, slim_bundle=bundle)

    def test_existing_style_in_catalog_passes(self):
        """apply_pStyle can reference an existing style from the catalog."""
        bundle = _slim_bundle(
            _default_paragraphs(),
            style_catalog=[{"styleId": "ExistingCSIStyle"}],
        )
        instr = _valid_instructions()
        instr["apply_pStyle"][5]["styleId"] = "ExistingCSIStyle"
        validate_instructions(instr, slim_bundle=bundle)

    def test_skippable_paragraphs_not_required(self):
        """Empty paragraphs, sectPr, END OF SECTION, editor notes should be skip-ok."""
        paragraphs = [
            _make_paragraph(0, "SECTION TITLE"),
            _make_paragraph(1, "PART 1"),
            _make_paragraph(2, "1.01 SUMMARY"),
            _make_paragraph(3, "A. Content"),
            _make_paragraph(4, "1. Sub"),
            _make_paragraph(5, "a. Subsub"),
            _make_paragraph(6, ""),                      # empty
            _make_paragraph(7, "END OF SECTION"),         # END OF SECTION
            _make_paragraph(8, "[Editor note]"),           # editor note
            _make_paragraph(9, "sectPr para", contains_sectPr=True),  # sectPr
        ]
        bundle = _slim_bundle(paragraphs)
        instr = _valid_instructions(apply_indices=[0, 1, 2, 3, 4, 5])
        # Need to fix derive_from to match smaller paragraph count
        instr["create_styles"][-1]["derive_from_paragraph_index"] = 5
        instr["roles"]["SUBSUBPARAGRAPH"]["exemplar_paragraph_index"] = 5
        instr["apply_pStyle"][-1] = {"paragraph_index": 5, "styleId": "CSI_Subsubparagraph__ARCH"}
        validate_instructions(instr, slim_bundle=bundle)


# ---------------------------------------------------------------------------
# Packaging checks
# ---------------------------------------------------------------------------

class TestPackaging:

    def test_prompt_files_exist(self):
        assert (REPO_ROOT / "master_prompt.txt").exists(), "master_prompt.txt missing from repo"
        assert (REPO_ROOT / "run_instruction_prompt.txt").exists(), "run_instruction_prompt.txt missing from repo"

    def test_schemas_exist(self):
        assert (REPO_ROOT / "schemas" / "phase1_instructions.schema.json").exists(), \
            "phase1_instructions.schema.json missing"
        assert (REPO_ROOT / "schemas" / "arch_style_registry.v1.schema.json").exists(), \
            "arch_style_registry.v1.schema.json missing"

    def test_instructions_schema_is_valid_json(self):
        schema_path = REPO_ROOT / "schemas" / "phase1_instructions.schema.json"
        data = json.loads(schema_path.read_text(encoding="utf-8"))
        assert data.get("$schema"), "Schema file missing $schema key"
        assert "create_styles" in str(data), "Schema should reference create_styles"

    def test_style_registry_schema_is_valid_json(self):
        schema_path = REPO_ROOT / "schemas" / "arch_style_registry.v1.schema.json"
        data = json.loads(schema_path.read_text(encoding="utf-8"))
        assert data.get("$schema"), "Schema file missing $schema key"
        assert data["properties"]["roles"].get("additionalProperties") is False, \
            "roles should have additionalProperties: false"

    def test_example_instructions_valid_json(self):
        instr_path = REPO_ROOT / "instructions.json"
        data = json.loads(instr_path.read_text(encoding="utf-8"))
        assert "create_styles" in data
        assert "apply_pStyle" in data
        assert "roles" in data


# ---------------------------------------------------------------------------
# Prompt loading helper
# ---------------------------------------------------------------------------

class TestLoadPrompt:

    @pytest.fixture(autouse=True)
    def _skip_no_tkinter(self):
        pytest.importorskip("tkinter")

    def test_missing_file_raises_clear_error(self, tmp_path: Path):
        from gui import _load_prompt
        with pytest.raises(FileNotFoundError, match="Required prompt file not found"):
            _load_prompt(tmp_path, "nonexistent_prompt.txt")

    def test_existing_file_loads(self, tmp_path: Path):
        from gui import _load_prompt
        prompt_file = tmp_path / "test_prompt.txt"
        prompt_file.write_text("Hello prompt", encoding="utf-8")
        result = _load_prompt(tmp_path, "test_prompt.txt")
        assert result == "Hello prompt"
