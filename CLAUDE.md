# CLAUDE.md — AI Assistant Guide for DOCX CSI Normalizer

## Project Overview

This is **Phase 1** of a two-phase DOCX specification automation pipeline. It takes an architect's Word specification template (.docx) and produces two formal contract artifacts:

1. **`arch_style_registry.json`** — Maps CSI (Construction Specifications Institute) structural roles to Word paragraph styles
2. **`arch_template_registry.json`** — Captures the complete formatting environment ("rendering VM snapshot")

Phase 2 (separate codebase) uses these artifacts to apply architect formatting to MEP (Mechanical/Electrical/Plumbing) consultant specs.

**The architect's template is sacred.** The output document must be pixel-identical to the input — only `<w:pStyle>` tags are inserted.

## Repository Structure

```
.
├── docx_decomposer.py          # Library module — extraction, slim bundle, style application
├── llm_classifier.py           # LLM automation — calls Anthropic API, chunking, coverage check
├── gui.py                      # Tkinter GUI wrapper (thin — no business logic)
├── arch_env_extractor.py       # Environment capture — produces arch_template_registry.json (has CLI)
├── phase1_validator.py         # Contract validation — validates both registries before writing
├── phase1_smoke_test.py        # Integration smoke test (extract → apply → validate)
├── master_prompt.txt           # System prompt for LLM CSI classification
├── run_instruction_prompt.txt  # Task prompt for LLM
├── instructions.json           # Example LLM output (style instructions)
├── schemas/
│   ├── arch_style_registry.v1.schema.json   # Formal JSON Schema for style registry
│   └── arch_template_registry.json          # Example/template for environment registry
├── tests/
│   ├── __init__.py
│   ├── test_arch_env_extractor.py           # Unit/integration tests for env extractor
│   ├── test_arch_template_registry_validation.py  # XML fragment parseability tests
│   └── test_phase1_validator.py             # Unit tests for contract validation
├── requirements.txt            # Runtime dependencies (anthropic)
├── requirements-build.txt      # PyInstaller build dependencies
├── *.docx                      # Sample architect specification templates
├── *_extracted/                 # DOCX extraction working directories (generated)
├── README.md
└── .gitignore
```

## Technology Stack

- **Language:** Python 3.8+
- **External API:** Anthropic (Claude) — for semantic CSI structure classification
- **Key stdlib modules:** `zipfile`, `re`, `json`, `xml.etree.ElementTree`, `hashlib`, `pathlib`, `tkinter`
- **Runtime dependency:** `anthropic` (for API calls)

## Architecture and Data Flow

### GUI Pipeline (gui.py)

```
DOCX (.docx file)
  │
  └─ [Run Phase 1] ──► extract ZIP
                          │
                          ├── build_slim_bundle() ──► slim_bundle.json
                          │                                │
                          │                   classify_document() (Anthropic API)
                          │                                │
                          │                                ▼
                          │                       instructions.json (saved for audit)
                          │                                │
                          ├── validate_instructions()      │
                          ├── apply_instructions()  ◄──────┘
                          │     ├── derive styles from exemplar paragraphs
                          │     ├── insert <w:pStyle> tags only
                          │     └── verify_stability() (hash checks)
                          │
                          ├──► arch_style_registry.json  ──► (copied to output folder)
                          ├──► arch_template_registry.json ──► (copied to output folder)
                          └──► coverage metric (% paragraphs classified)
```

## Critical Design Invariants

**These are hard rules. Violating them will break the pipeline or corrupt documents.**

1. **Never full-XML-parse `document.xml`** — Uses regex (`iter_paragraph_xml_blocks()`) to preserve paragraph indices and raw XML structure. ElementTree is only used for `styles.xml` name lookups and catalog building.

2. **Surgical XML insertion only** — The only modification to `document.xml` is inserting/replacing `<w:pStyle>` elements. Nothing else may change.

3. **Exemplar-based formatting** — New CSI styles are derived from actual paragraphs in the template (`derive_from_paragraph_index`). The LLM is forbidden from specifying any formatting (pPr, rPr, fonts, spacing, alignment, etc.).

4. **Stability snapshots** — `StabilitySnapshot` (dataclass) records SHA-256 hashes of headers, footers, section properties, and document.xml.rels before any modifications. `verify_stability()` enforces these haven't changed after processing.

5. **No sectPr paragraphs** — Paragraphs containing `<w:sectPr>` are never styled and never used as exemplars.

6. **No DOCX reconstruction** — Phase 1 intentionally does NOT produce a .docx output file. It works on the extracted folder only.

## Key Source Files

### `docx_decomposer.py` — Library Module

| Function | Purpose |
|---|---|
| `sha256_bytes()` / `sha256_text()` | Hash utilities for bytes and strings |
| `snapshot_headers_footers()` | Captures header/footer file hashes for stability checks |
| `snapshot_doc_rels_hash()` | Captures document.xml.rels hash for stability checks |
| `extract_sectpr_block()` | Extracts sectPr XML blocks from document.xml |
| `snapshot_stability()` / `verify_stability()` | Hash-based invariant enforcement |
| `extract_docx()` | Unzips .docx into workspace directory (with OneDrive lock retry) |
| `iter_paragraph_xml_blocks()` | Regex iterator over `<w:p>` blocks — preserves indices |
| `paragraph_text_from_block()` | Extracts visible text from paragraph XML |
| `paragraph_contains_sectpr()` | Checks if paragraph contains `<w:sectPr>` |
| `paragraph_pstyle_from_block()` | Extracts existing `<w:pStyle>` value from paragraph |
| `paragraph_numpr_from_block()` | Extracts numbering properties (numId, ilvl) |
| `paragraph_ppr_hints_from_block()` | Extracts lightweight pPr hints (alignment, indent, spacing) |
| `build_style_catalog()` | Builds style name/type catalog from `styles.xml` with basedOn chain |
| `build_numbering_catalog()` | Builds numbering definition catalog from `numbering.xml` |
| `build_slim_bundle()` | Creates minimal JSON (text + numbering hints) for LLM input |
| `strip_pstyle_from_paragraph()` | Removes `<w:pStyle>` tags from paragraph XML |
| `ppr_without_pstyle()` | Extracts pPr without pStyle |
| `xml_escape()` | XML entity encoding for strings |
| `extract_paragraph_ppr_inner()` | Extracts pPr inner XML from exemplar paragraph |
| `extract_paragraph_rpr_inner()` | Extracts representative rPr inner XML (handles multi-run) |
| `derive_style_def_from_paragraph()` | Extracts pPr/rPr from exemplar paragraph to build style definition |
| `build_style_xml_block()` | Generates `<w:style>` XML for insertion into `styles.xml` |
| `insert_styles_into_styles_xml()` | Inserts style blocks into `styles.xml` |
| `validate_instructions()` | Strict validation of LLM output before application |
| `apply_instructions()` | Main apply logic: create styles, insert pStyle, verify stability |
| `apply_pstyle_to_paragraph_block()` | Surgically inserts `<w:pStyle>` into a single paragraph |
| `build_style_registry_dict()` | Builds arch_style_registry payload dict (without writing to disk) |
| `emit_arch_style_registry()` | Writes the final `arch_style_registry.json` contract |

### `llm_classifier.py` — LLM Automation

Pure module (no CLI) — called by `gui.py`.

| Function | Purpose |
|---|---|
| `classify_document()` | Main entry: calls Anthropic API with slim bundle, returns instructions dict. Default model: `claude-opus-4-6` |
| `compute_coverage()` | Computes % of classifiable paragraphs that received a style |
| `estimate_tokens()` | Rough token count for chunking decisions |

**Design constraints:** No CLI of its own. Retry logic (up to 2 retries) for transient API failures. Chunking activates automatically when token estimate > 80K or paragraphs > 300.

### `gui.py` — Tkinter GUI (Primary Entry Point)

Thin wrapper over the pipeline functions — no business logic.

| Class | Purpose |
|---|---|
| `App` | Main window: template picker, API key field, output folder picker, Run button, log area, status |
| `PipelineThread` | Background thread that runs the full pipeline |
| `LogRedirector` | Thread-safe stdout redirector for log display |

### `arch_env_extractor.py` — Environment Capture (has CLI)

Has a full CLI entry point (`python arch_env_extractor.py`) in addition to being importable as a library.

| Function | Purpose |
|---|---|
| `extract_arch_template_registry()` | Main orchestrator — builds complete registry |
| `extract_doc_defaults()` | Extracts `<w:docDefaults>` (baseline rPr/pPr) |
| `extract_style_defs()` | All style definitions with raw XML blocks |
| `extract_latent_styles()` | Extracts latent style exception settings |
| `extract_table_styles()` | Extracts table-specific style definitions |
| `extract_styles_section()` | Composite: doc_defaults + style_defs + latent + table styles |
| `extract_theme()` | Theme fonts and colors from `theme1.xml` |
| `extract_settings()` | Compatibility flags from `settings.xml` |
| `extract_page_layout()` | Section properties, margins, columns |
| `extract_headers_footers()` | Complete header/footer XML |
| `extract_numbering()` | Numbering definitions from `numbering.xml` |
| `extract_fonts()` | Font table declarations |
| `extract_relationships()` | Document relationship entries |
| `extract_package_inventory()` | Inventories which OOXML parts are present |
| `extract_docx_to_dir()` | Extracts .docx ZIP to a directory |

### `phase1_validator.py` — Contract Validation

Standalone validation module — validates both registries before writing to disk. Imported by `phase1_smoke_test.py` and usable independently.

| Function | Purpose |
|---|---|
| `validate_template_registry()` | Validates `arch_template_registry.json` structure and XML fragment well-formedness |
| `validate_style_registry()` | Validates `arch_style_registry.json` structure, version, required roles |
| `validate_cross_registry()` | Verifies every role `style_id` in the style registry exists in the template registry's `style_defs` |
| `validate_phase1_contracts()` | Top-level entry: runs all three validations in sequence |

### `phase1_smoke_test.py` — Integration Smoke Test

Calls `extract_docx()`, `build_slim_bundle()`, `apply_instructions()`, `build_style_registry_dict()`, and `extract_arch_template_registry()` directly. Delegates registry validation to `phase1_validator.validate_phase1_contracts()`. The test fails if either registry is missing, malformed, or contains truncated XML fragments.

### `tests/` — Unit and Integration Tests

| File | What it tests |
|---|---|
| `test_phase1_validator.py` | Unit tests for all four `phase1_validator` functions (required/optional roles, cross-registry consistency, XML well-formedness) |
| `test_arch_env_extractor.py` | Unit and integration tests for extraction functions (block extraction, XML fragment completeness) |
| `test_arch_template_registry_validation.py` | XML fragment parseability validation for generated registry output |

## Commands

### GUI (primary entry point)
```bash
python gui.py
```

The GUI provides:
- **Template** picker — select the architect .docx specification
- **API Key** field — Anthropic API key (pre-populated from `ANTHROPIC_API_KEY` env var)
- **Output Folder** picker — where `arch_style_registry.json` and `arch_template_registry.json` are written (defaults to same directory as the input .docx)
- **Run Phase 1** button — runs the full pipeline
- Post-completion buttons to open the output folder or view the style registry

### Standalone Environment Extraction
```bash
# From a .docx file
python arch_env_extractor.py TEMPLATE.docx

# From an already-extracted folder
python arch_env_extractor.py --extract-dir TEMPLATE_extracted

# Custom output path
python arch_env_extractor.py TEMPLATE.docx --output /path/to/output.json
```

### Smoke Test
```bash
python phase1_smoke_test.py TEMPLATE.docx instructions.json
```

### Unit Tests
```bash
python -m pytest tests/
```

## CSI Role Hierarchy and Allowed Style IDs

The pipeline recognizes these CSI structural roles (from schema):

| Role | Style ID | Required? |
|---|---|---|
| `SectionID` | `CSI_SectionID__ARCH` | Optional |
| `SectionTitle` | `CSI_SectionTitle__ARCH` or `CSI_SectionName__ARCH` | Required |
| `PART` | `CSI_Part__ARCH` | Required |
| `ARTICLE` | `CSI_Article__ARCH` | Required |
| `PARAGRAPH` | `CSI_Paragraph__ARCH` | Required |
| `SUBPARAGRAPH` | `CSI_Subparagraph__ARCH` | Required |
| `SUBSUBPARAGRAPH` | `CSI_Subsubparagraph__ARCH` | Optional |

All created style IDs must match the pattern `CSI_*__ARCH`.

## Output Artifacts

### `arch_style_registry.json`
```json
{
  "version": 1,
  "source_docx": "TEMPLATE.docx",
  "roles": {
    "PART": { "style_id": "CSI_Part__ARCH", "exemplar_paragraph_index": 4, "style_name": "..." },
    ...
  }
}
```
Validated against `schemas/arch_style_registry.v1.schema.json`.

### `arch_template_registry.json`
Complete formatting environment with sections: `meta`, `package_inventory`, `doc_defaults`, `styles`, `theme`, `settings`, `page_layout`, `headers_footers`, `numbering`, `fonts`, `custom_xml`, `capture_policy`.

### Coverage Metric
After classification, the pipeline reports what percentage of non-empty, non-sectPr, non-editor-note paragraphs received a style. Coverage below 90% triggers a warning.

## Development Conventions

### Code Style
- Python 3.8+ compatible (uses `from __future__ import annotations`)
- Type hints throughout (`Dict`, `List`, `Optional`, `Tuple`, `Set`, `Any` from `typing`)
- Frozen dataclasses for immutable state (`StabilitySnapshot`)
- Functions are well-documented with inline comments explaining "why"

### XML Handling
- **Regex-first for `document.xml`** — preserves byte-level structure and paragraph indices
- **ElementTree only for read-only lookups** on `styles.xml`, `numbering.xml`
- Raw XML blocks are stored as strings in JSON registries (not parsed/re-serialized)
- `_canonicalize()` strips rsids and proofing marks for cleaner output

### Error Handling
- Hard `ValueError` raises for all invariant violations
- No silent failures — every validation check is explicit
- Descriptive error messages with context (paragraph index, style ID, etc.)

### Testing
- Unit/integration tests in `tests/` directory (pytest-compatible)
- `phase1_smoke_test.py` for end-to-end integration with real DOCX files
- Stability verification is built into the apply pipeline itself
- Smoke test creates timestamped extraction directories to avoid collisions

## Common Pitfalls When Modifying This Code

1. **Do not switch `document.xml` parsing to ElementTree** — it will reformat XML and break paragraph index mapping.

2. **Do not add formatting fields to the LLM instruction schema** — the LLM must never specify pPr/rPr. Only `derive_from_paragraph_index` is allowed.

3. **Do not modify paragraphs containing `<w:sectPr>`** — these are section break containers and styling them can corrupt the document.

4. **Do not remove stability checks** — they are the primary safety mechanism ensuring the template isn't corrupted.

5. **`requirements.txt` is for runtime dependencies** (`anthropic`). Build/packaging dependencies are in `requirements-build.txt`.

6. **The `.docx` files and `*_extracted/` directories in the repo are test data** — they are architect specification templates used for development and testing.

7. **`llm_classifier.py` must remain a pure module** — no CLI of its own. It is imported only by `gui.py`.

8. **`gui.py` must remain a thin wrapper** — no pipeline logic. It imports and calls library functions from `docx_decomposer.py`.

9. **`docx_decomposer.py` is a library module only** — no CLI entry point. All user interaction goes through `gui.py`.

10. **`phase1_validator.py` is the single source of truth for registry validation** — all validation logic (template registry structure, style registry roles, cross-registry consistency) lives here. Imported by `phase1_smoke_test.py`, `gui.py`, and `arch_env_extractor.py`.

## Environment Setup

```bash
pip install -r requirements.txt
export ANTHROPIC_API_KEY='your-key-here'
```

Runtime: Python 3.8+ on Windows or Linux.

## License

Copyright 2025 Abraham Borg. All Rights Reserved. Proprietary software — no license to use, copy, modify, or distribute without written permission.
