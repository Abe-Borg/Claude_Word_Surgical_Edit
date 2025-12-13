# DOCX CSI Normalizer

A tool that adds CSI structural tagging to Word specification documents without changing how they look.

## What it does

Takes an architect's Word spec template and:
1. Identifies CSI structural elements (PART, Article, Paragraph, etc.)
2. Learns the formatting from the document itself
3. Creates Word paragraph styles based on actual formatting
4. Applies those styles to matching paragraphs
5. Guarantees zero visual change to the document

The output looks identical but now has proper paragraph styles you can work with programmatically.

## Installation

```bash
pip install anthropic  # Required for LLM API calls
```

Set your Anthropic API key:
```bash
export ANTHROPIC_API_KEY='your-key-here'
```

## Usage

### Basic workflow

```bash
# Extract the DOCX to see what you're working with
python docx_decomposer.py extract input.docx

# Run the normalizer (creates styles and applies them)
python docx_decomposer.py normalize input.docx output.docx

# Compare before/after to verify no visual drift
python docx_decomposer.py compare input.docx output.docx
```

### Available commands

**`extract <docx_file> [output_dir]`**  
Unzips the DOCX into a folder for inspection.

**`analyze <docx_file>`**  
Generates a detailed markdown report of the document structure.

**`slim-bundle <docx_file>`**  
Creates a minimal JSON representation (text, numbering hints, styles) for LLM analysis.

**`normalize <input.docx> <output.docx> [--custom-prompt prompt.txt]`**  
Main command. Adds CSI paragraph styles without changing appearance.
- Uses Claude API to classify structure
- Derives formatting from exemplar paragraphs
- Applies styles in-place
- Verifies zero drift

**`compare <before.docx> <after.docx>`**  
Visual diff showing what changed (should be nothing except style tags).

**`reconstruct <extracted_dir> <output.docx>`**  
Rebuilds a DOCX from an extracted folder.

### Optional: Custom prompts

By default, `normalize` uses the built-in prompt. To customize:

```bash
python docx_decomposer.py normalize input.docx output.docx --custom-prompt my_instructions.txt
```

## How it works

1. **Extract**: Unzips DOCX, records hashes of headers/footers/section properties
2. **Analyze**: LLM sees slim bundle (text + numbering context only, no formatting)
3. **Classify**: LLM returns JSON with CSI role assignments and exemplar paragraph indices
4. **Derive**: Script extracts formatting from exemplar paragraphs chosen by LLM
5. **Apply**: Inserts `<w:pStyle>` tags into paragraphs by index
6. **Verify**: Fails if anything changed except `<w:pStyle>` additions

## What gets created

After running `normalize`, the document will have these paragraph styles:
- `CSI_SectionTitle__ARCH`
- `CSI_Part__ARCH`
- `CSI_Article__ARCH`
- `CSI_Paragraph__ARCH`
- `CSI_Subparagraph__ARCH`
- `CSI_Subsubparagraph__ARCH`

Each style captures the exact formatting from representative paragraphs in the original document.

## What it doesn't do

- Change visual appearance
- Modify numbering definitions
- Touch headers, footers, or section breaks
- Normalize spacing, indents, or alignment
- Generate new content

These are intentional safeguards. The architect's template is sacred.

## Safety features

- Hard fails if headers/footers change
- Hard fails if section properties change
- Hard fails if relationships change
- Hard fails if paragraph properties drift beyond `<w:pStyle>` insertion
- LLM is forbidden from specifying formatting (only structure classification)

## Example output

```json
{
  "create_styles": [
    {
      "styleId": "CSI_Part__ARCH",
      "name": "CSI Part (Architect Template)",
      "type": "paragraph",
      "derive_from_paragraph_index": 12
    }
  ],
  "apply_pStyle": [
    {"paragraph_index": 12, "styleId": "CSI_Part__ARCH"},
    {"paragraph_index": 45, "styleId": "CSI_Part__ARCH"},
    {"paragraph_index": 78, "styleId": "CSI_Part__ARCH"}
  ],
  "notes": ["Applied CSI_Part__ARCH to 3 PART headings"]
}
```

## Requirements

- Python 3.8+
- Anthropic API key (Claude Sonnet 4)
- Windows or Linux (tested on both)

## Troubleshooting

**"Paragraph drift detected"**  
The script changed something it shouldn't have. This is a bug, not expected behavior.

**"derive_from_paragraph_index out of range"**  
LLM referenced a paragraph that doesn't exist. Try re-running or check your custom prompt.

**"LLM formatting fields are forbidden"**  
LLM tried to specify formatting directly instead of referencing an exemplar. This violates the contract.

## License

Do whatever you want with it.