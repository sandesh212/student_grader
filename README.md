# Student Assessment Document Parser

A Python tool that intelligently extracts and classifies questions, instructions, headings, and answers from Microsoft Word (.docx) assessment marking sheets.

## Features

- Smart question detection (numbering, directive verbs, question marks)
- Instruction extraction (admin/assessment metadata)
- Answer extraction (red-colored text)
- Heading recognition
- Automatic deduplication
- Bullet-point intelligence

## Installation

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

## Usage

```bash
python3 grader.py
```

## Requirements

- Python 3.8+
- python-docx

## License

MIT
