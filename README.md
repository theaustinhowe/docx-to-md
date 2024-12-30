# Docx to Markdown Converter

This repository contains a Python script for converting `.docx` files into Markdown format. It supports:

- **Headings**: Converts `Heading 1` to `Heading 6` styles into corresponding Markdown `#` syntax.
- **Text Formatting**: Preserves bold, italic, and bold-italic formatting.
- **Tables**: Converts tables into Markdown tables, including headers and rows.
- **Images**: Adds image references if present in the document.

## Features
- Easy-to-use script with minimal dependencies.
- Comprehensive Markdown support for headings, formatted text, tables, and images.
- Compatible with Python 3.

## Requirements
- Python 3.7+
- `python-docx` library

## Installation
1. Clone this repository:
   ```bash
   git clone https://github.com/theaustinhowe/docx-to-md.git
   cd docx-to-md
   ```

2. Install dependencies:
   ```bash
   pip install python-docx
   ```

## Usage
1. Place your `.docx` file in the same directory as the script.
2. Update the `input_docx` and `output_md` variables in the script:
   ```python
   input_docx = "your_file.docx"  # Replace with your .docx file path
   output_md = "output.md"         # Replace with your desired output file path
   ```
3. Run the script:
   ```bash
   python docx_to_markdown.py
   ```
4. The Markdown file will be saved to the specified `output_md` path.

## Example Output
### Input (`.docx`):
```
# Heading 1
## Heading 2
- **Bold text**
- *Italic text*
- ***Bold and Italic text***
| Header 1 | Header 2 |
|----------|----------|
| Row 1    | Data     |
```

### Output (`.md`):
```markdown
# Heading 1

## Heading 2

- **Bold text**
- *Italic text*
- ***Bold and Italic text***

| Header 1 | Header 2 |
|----------|----------|
| Row 1    | Data     |
```

## Contributing
Contributions are welcome! Feel free to submit a pull request or open an issue.

## License
This project is licensed under the MIT License. See the `LICENSE` file for details.

## Author
Created by [Austin Howe](https://github.com/theaustinhowe).
