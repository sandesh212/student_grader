# Student Assessment Document Parser

A Python tool that intelligently extracts and classifies questions, instructions, headings, and answers from Microsoft Word (.docx) assessment marking sheets, and automatically identifies which training.gov.au unit each question belongs to based on content similarity.

## Features

- Smart question detection (numbering, directive verbs, question marks)
- **NEW: Automatic unit classification** - matches questions to training.gov.au units based on content
- Instruction extraction (admin/assessment metadata)
- Answer extraction (red-colored text)
- Heading recognition
- Automatic deduplication
- Bullet-point intelligence
- TF-IDF based text similarity matching

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
- openpyxl
- scikit-learn

## How Unit Classification Works

The tool uses TF-IDF (Term Frequency-Inverse Document Frequency) vectorization and cosine similarity to match questions to training units:

1. Loads unit codes and descriptions from `Units.xlsx`
2. Converts both questions and unit descriptions into TF-IDF vectors
3. Calculates cosine similarity between each question and all units
4. Returns the best matching unit with a confidence score

This allows the system to learn from the content of questions and automatically identify which training.gov.au unit they relate to, even without explicit mentions of unit codes.

## License

MIT
