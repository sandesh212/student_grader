import os
import re
from docx import Document
import openpyxl
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
import numpy as np

# ---------- heuristics (structure first, minimal keywords) ----------

def load_units_from_excel(excel_path='Units.xlsx'):
    """
    Load unit codes and descriptions from Units.xlsx.
    Returns a list of tuples: [(unit_code, description), ...]
    """
    units = []
    try:
        wb = openpyxl.load_workbook(excel_path)
        ws = wb.active
        for row in ws.iter_rows(values_only=True):
            if row[0] and row[1]:  # Both unit code and description must exist
                units.append((str(row[0]).strip(), str(row[1]).strip()))
        wb.close()
    except Exception as e:
        print(f"Warning: Could not load units from {excel_path}: {e}")
    return units

def classify_question_to_unit(question_text, units, threshold=0.1):
    """
    Match a question to the most relevant unit based on text similarity.
    
    Args:
        question_text: The question to classify
        units: List of (unit_code, description) tuples
        threshold: Minimum similarity score (0-1) to return a match
    
    Returns:
        Tuple of (unit_code, description, similarity_score) or None if no match
    """
    if not units or not question_text:
        return None
    
    # Prepare texts for comparison
    unit_descriptions = [desc for code, desc in units]
    all_texts = unit_descriptions + [question_text]
    
    # Use TF-IDF vectorization for text similarity
    try:
        vectorizer = TfidfVectorizer(stop_words='english', lowercase=True)
        tfidf_matrix = vectorizer.fit_transform(all_texts)
        
        # Calculate cosine similarity between question and all units
        question_vec = tfidf_matrix[-1]
        unit_vecs = tfidf_matrix[:-1]
        similarities = cosine_similarity(question_vec, unit_vecs)[0]
        
        # Find the best match
        best_idx = np.argmax(similarities)
        best_score = similarities[best_idx]
        
        if best_score >= threshold:
            unit_code, unit_desc = units[best_idx]
            return (unit_code, unit_desc, best_score)
    except Exception as e:
        print(f"Warning: Error in unit classification: {e}")
    
    return None

# ---------- heuristics (structure first, minimal keywords) ----------

def is_heading(block: str) -> bool:
    text = block.strip()
    if not text:
        return False
    low = text.lower()
    # Labeled section headers
    if re.match(r'^(part|section)\s+[a-z0-9]+$', low):
        return True
    # Short ALL CAPS line
    if text.isupper() and 5 <= len(text) <= 60:
        return True
    return False

def is_question_block(block: str) -> bool:
    """
    Structural signals only:
    - Starts with numbering (1., 1.1., 3), a), i) …)
    - Ends with a question mark
    - Contains fill-in patterns (____, …)
    """
    t = block.strip()
    if not t:
        return False

    first = t.splitlines()[0].strip()

    # Numbering patterns
    if re.match(r'^(\d+(?:\.\d+)*[.)]|\d+\.)\s+', first):
        return True
    if re.match(r'^([a-z]\)|\([a-z]\)|[ivxlcdm]+\))\s+', first, re.IGNORECASE):
        return True
    
    # Question mark anywhere
    if '?' in t:
        return True
    
    # Fill-in blanks
    if re.search(r'_{3,}|\.\.\.|_____', t):
        return True

    # ADDED: Directive verbs that indicate questions
    if re.search(r'\b(list|describe|name|state|explain|identify|give|provide|calculate)\b', first.lower()):
        return True

    return False

def is_instruction(block: str) -> bool:
    """
    Administrative/meta text about the assessment process.
    """
    t = block.strip()
    if not t:
        return False
    low = t.lower()

    # IMPORTANT: Don't classify questions as instructions
    if is_question_block(t):
        return False

    # Strong admin indicators (expanded list)
    admin = [
        r'\bassessor', r'\bparticipant', r'\bmarking', r'\bcriteria',
        r'\bmodel answer', r'\bsatisfactory response', r'\blog\b',
        r'\bdocument the\b', r'\brefer to\b', r'\bindicate whether\b',
        r'\bcourse participant', r'\bmarking sheet\b', r'\btechnically correct\b',
        r'\bacceptable responses\b', r'\blevel of detail\b', r'\bminimum required\b',
        r'\bbasic level of knowledge\b', r'\bmust be completed\b', r'\binitial next to it\b',
        r'\banswers within this', r'\blisted below are', r'\bmust refer to these\b'
    ]
    if any(re.search(p, low) for p in admin):
        return True

    # Check for bullet point structure (multiple lines starting with bullets)
    lines = [ln.strip() for ln in t.splitlines() if ln.strip()]
    
    # Count lines that start with bullet markers
    bullet_count = sum(1 for ln in lines if re.match(r'^[•\-\*]\s+', ln))
    
    # If 3+ bullet points, it's likely an instruction list
    if bullet_count >= 3:
        return True
    
    # If 2+ bullets AND contains admin keywords
    if bullet_count >= 2 and any(re.search(p, low) for p in admin[:5]):
        return True

    # Header-like instruction (but only if NOT a question)
    if lines:
        first = lines[0]
        if first.endswith(':') and not re.search(r'\b(list|describe|name|state|explain)\b', first.lower()):
            return True

    return False

# ---------- text extraction helpers ----------

def extract_red_and_black_text(cell):
    """
    Return (red_fragments, black_text_with_newlines)
    """
    red = []
    black_parts = []

    for p in cell.paragraphs:
        para_parts = []
        for run in p.runs:
            txt = run.text or ''
            if not txt:
                continue
            is_red = False
            try:
                color = getattr(run.font, "color", None)
                rgb = getattr(color, "rgb", None)
                if rgb is not None:
                    # Accept any representation containing FF0000
                    if "FF0000" in str(rgb).upper():
                        is_red = True
            except Exception:
                pass

            if is_red:
                if txt.strip():
                    red.append(txt.strip())
            else:
                para_parts.append(txt)
        # preserve paragraph boundaries
        if para_parts:
            black_parts.append("".join(para_parts))
        # even if no non-red in this paragraph, keep a newline to avoid merging
        black_parts.append("\n")

    black_text = "".join(black_parts).replace("\r\n", "\n")
    # Collapse multiple blank lines
    black_text = re.sub(r'\n{3,}', '\n\n', black_text).strip()
    return red, black_text

# ---------- splitting logic (prevents merging) ----------

def split_into_question_blocks(text: str):
    """
    Normalize and split into atomic question blocks.
    DON'T split bullet-pointed instruction blocks.
    """
    if not text:
        return []

    # First check: if this looks like a multi-line bullet instruction block, keep it whole
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    bullet_count = sum(1 for ln in lines if re.match(r'^[•\-\*]\s+', ln))
    
    # If 3+ bullets in the whole text, treat as single instruction block
    if bullet_count >= 3:
        return [text.strip()]

    s = text

    # Force newlines before markers that might appear mid-line
    s = re.sub(r'(?i)(?<!\n)(\b\d+(?:\.\d+)*[.)]|\b\d+\.)\s+', r'\n\1 ', s)
    s = re.sub(r'(?i)(?<!\n)(\b[a-z]\)|\([a-z]\)|\b[ivxlcdm]+\))\s+', r'\n\1 ', s)

    # Split after question marks if there is more content following
    s = re.sub(r'\?\s*(?=\S)', '?\n', s)

    # Split by lines, then group into blocks that start with a marker or contain a '?'
    lines = [ln.strip() for ln in s.split('\n') if ln.strip()]

    def is_start_marker(line: str) -> bool:
        return bool(
            re.match(r'^(\d+(?:\.\d+)*[.)]|\d+\.)\s+', line) or
            re.match(r'^([a-z]\)|\([a-z]\)|[ivxlcdm]+\))\s+', line, re.IGNORECASE)
        )

    blocks = []
    cur = []

    for ln in lines:
        if is_start_marker(ln) or ln.endswith('?'):
            # start a new block
            if cur:
                blocks.append("\n".join(cur).strip())
            cur = [ln]
        else:
            cur.append(ln)

    if cur:
        blocks.append("\n".join(cur).strip())

    # Post-process: if a block contains multiple questions (multiple '?'),
    # split them further to keep one question per block.
    final_blocks = []
    for b in blocks:
        parts = re.split(r'(\?)', b)
        if parts.count('?') <= 1:
            final_blocks.append(b)
        else:
            # Recombine text so each piece ending with '?' is its own block
            acc = ""
            for seg in parts:
                acc += seg
                if seg == '?':
                    final_blocks.append(acc.strip())
                    acc = ""
            if acc.strip():
                final_blocks.append(acc.strip())

    # Remove tiny fragments that are clearly not questions/headers
    final_blocks = [blk for blk in final_blocks if len(blk) >= 2]

    return final_blocks

# ---------- classification pipeline ----------

def normalize_block_key(block: str) -> str:
    return re.sub(r'[\s,.;:!?\-]+', '', block.lower())

def extract_and_classify_blocks(docx_path):
    doc = Document(docx_path)

    headings, questions, instructions, red_texts = [], [], [], []
    seen_h, seen_q, seen_i = set(), set(), set()

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                reds, black = extract_red_and_black_text(cell)
                red_texts.extend(reds)

                if not black:
                    continue

                # Split cell text into atomic blocks BEFORE classifying
                for block in split_into_question_blocks(black):
                    key = normalize_block_key(block)
                    if not key:
                        continue

                    # FIXED ORDER: heading -> question -> instruction
                    # (questions must be checked BEFORE instructions!)
                    if is_heading(block):
                        if key not in seen_h:
                            headings.append(block)
                            seen_h.add(key)
                        continue

                    if is_question_block(block):
                        if key not in seen_q:
                            questions.append(block)
                            seen_q.add(key)
                        continue

                    if is_instruction(block):
                        if key not in seen_i:
                            instructions.append(block)
                            seen_i.add(key)
                        continue

                    # If none matched, ignore

    return headings, questions, instructions, red_texts

# ---------- CLI ----------

def main():
    doc_filename = input("Enter DOCX filename (just file, no path): ").strip()
    if os.path.isfile(doc_filename):
        path = doc_filename
    elif os.path.isfile(os.path.join("samples", doc_filename)):
        path = os.path.join("samples", doc_filename)
    else:
        print(f"ERROR: File '{doc_filename}' not found in CWD or samples/.")
        return

    print(f"Processing: {path}\n")
    
    # Load units from Excel file
    print("Loading training units from Units.xlsx...")
    units = load_units_from_excel()
    if units:
        print(f"Loaded {len(units)} training units.\n")
    else:
        print("Warning: No units loaded. Unit classification will be skipped.\n")
    
    heads, qs, instr, reds = extract_and_classify_blocks(path)

    # Classify questions to units
    question_classifications = []
    if units and qs:
        print("Classifying questions to training units...\n")
        for q in qs:
            classification = classify_question_to_unit(q, units)
            question_classifications.append(classification)

    print("=" * 60)
    print("IDENTIFIED QUESTIONS WITH UNIT CLASSIFICATION")
    print("=" * 60)
    for i, q in enumerate(qs, 1):
        print(f"{i}. {q}")
        if i <= len(question_classifications) and question_classifications[i-1]:
            unit_code, unit_desc, score = question_classifications[i-1]
            print(f"   → Unit: {unit_code} - {unit_desc}")
            print(f"   → Confidence: {score:.2%}\n")
        else:
            print(f"   → Unit: Not classified\n")

    print("\n" + "=" * 60)
    print("INSTRUCTIONS")
    print("=" * 60)
    for i, ins in enumerate(instr, 1):
        print(f"{i}. {ins}\n")

    print("\n" + "=" * 60)
    print("HEADINGS")
    print("=" * 60)
    for i, h in enumerate(heads, 1):
        print(f"{i}. {h}\n")

    print("\n" + "=" * 60)
    print("RED TEXT (ANSWERS)")
    print("=" * 60)
    for i, r in enumerate(reds, 1):
        print(f"{i}. {r}")

if __name__ == "__main__":
    main()
