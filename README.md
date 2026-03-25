# fast-ods

A fast, streaming parser for OpenDocument Spreadsheet (`.ods`) file.

Designed for efficiency, performance and low memory usage, making it an ideal choice for large files and ETL workflows.

> Requires Python version 3.10 and above.

> This project is still in alpha (pre-1.0). API and behavior might change drastically until the full release.

## Features

- ⚡ Streaming parsing (suitable for handling large files with constant memory usage)
- 📈 Supports repeated rows and columns (.ods specific)
- 🧩 Simple, minimal API
- 🌐 Zero external dependencies
- 🔧 Flexible configuration settings

## Scope

This library focuses on **fast data extraction**, not full ODS spec compliance. It itentionally does not handle features such as:

- Styling
- Formulas
- Metadata

## Installation

Install with the Python default package manager:

```bash
pip install fast-ods
```

## Usage

```python
from fast_ods import ODSParser

parser = ODSParser()

for row in parser.parse("file.ods"):
    print(row)
```

## Options

```python
from fast_ods import ODSParser, ODSParserOptions

options = ODSParserOptions(
    table=0,                # Choose a table by index or name
    convert_values=False,   # Convert the cell values (float, date, etc.)
    skip_n_rows=None,       # Skip first N rows
    take_n_rows=None,       # Limit number of rows
    verify_zip=True         # Emits a warning if the .zip archive is corrupted 
)

parser = ODSParser(options)
```