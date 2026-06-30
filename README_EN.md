English | [中文](README.md)

# 🔄 Smart Table Format Converter

Upload source data + specify target format, AI auto-generates column mapping rules, preview and convert with one click.

## Features

- **3 target format modes**: Natural language instructions / Upload template / Saved presets
- **AI auto-mapping**: Call DeepSeek API to auto-identify source-to-target column correspondence and transform rules
- **Preview before export**: Mapping comparison table + converted data preview, confirm before downloading
- **Rich transformations**: Column name mapping, date format conversion, number formatting, computed columns, row filtering, sorting
- **Preset management**: Save frequently used mapping rules for one-click reuse
- **Multi-format output**: Excel (with formatting) / CSV
- **Template-based output**: Upload a target template to preserve its original formatting

## Installation

```bash
pip install -r requirements.txt
```

## Quick Start

```bash
streamlit run app.py
```

## Usage

1. **Upload source data** — Excel or CSV file
2. **Specify target format** — 3 options:
   - 📝 **Natural language instructions** — Describe target columns and format requirements in plain text
   - 📎 **Upload template** — Auto-read template column headers
   - 💾 **Use preset** — Load previously saved mapping rules
3. **AI generates mapping** — Auto-match source columns → target columns + transform rules
4. **Preview & confirm** — View mapping comparison table and converted data
5. **Export** — Download Excel or CSV

## AI Configuration (Optional)

Enter your DeepSeek API Key in the sidebar to enable AI auto-mapping. Without it, you'll need to manually configure JSON mapping rules.

Also supports replacing with any OpenAI-compatible API endpoint.

## Supported Transform Types

| Transform | Description | Example |
|-----------|-------------|---------|
| `none` | Copy as-is | — |
| `text` | Convert to text | `100` → `"100"` |
| `strip` | Trim whitespace | `" hello "` → `"hello"` |
| `round:N` | Round to N decimals | `3.14159` → `3.14` |
| `date:SRC→DST` | Date format conversion | `date:YYYYMMDD→YYYY-MM-DD` |
| `prefix:X` | Add prefix | `"hello"` → `"XX_hello"` |
| `suffix:X` | Add suffix | `"hello"` → `"hello_END"` |
| `upper` / `lower` | Case conversion | `"abc"` → `"ABC"` |
| `regex:P:R` | Regex replacement | `regex:\s:_` → replace spaces with underscores |

## Mapping Rule JSON Format

```json
{
  "mappings": [
    {"source": "source_column", "target": "target_column", "transform": "none"}
  ],
  "computed": [
    {"target": "new_column", "formula": "concat(col_a, '-', col_b)"}
  ],
  "filters": [
    {"column": "status", "op": "eq", "value": "active"}
  ],
  "sort": {"column": "date", "order": "asc"}
}
```

## Presets

Save your mapping rules as named presets in the sidebar for quick reuse next time. Presets are stored locally in the `presets/` directory.
