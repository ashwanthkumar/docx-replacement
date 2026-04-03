# docx-replacement

Replace `${VAR_NAME}` placeholders in `.docx` files with values from a JSON file.

## Setup

Requires Python 3.12+ and [uv](https://docs.astral.sh/uv/).

```bash
uv sync
```

## Usage

```bash
uv run python main.py <input.docx> <variables.json> [-o output.docx]
```

- `input.docx` — the template document containing `${VAR_NAME}` placeholders
- `variables.json` — a JSON file mapping variable names to replacement values
- `-o` / `--output` — optional output path (defaults to `<input>_filled.docx`)

### Example

Given a `variables.json`:

```json
{
  "NAME": "Ashwanth Kumar",
  "DATE": "2026-04-03",
  "COMPANY": "Acme Corp"
}
```

Run:

```bash
uv run python main.py template.docx variables.json -o filled.docx
```

All occurrences of `${NAME}`, `${DATE}`, `${COMPANY}` in the document (paragraphs and table cells) will be replaced. Unmatched variables are left as-is.

See [`DOCX Replacement - Sample.docx`](DOCX%20Replacement%20-%20Sample.docx) for an example template.
