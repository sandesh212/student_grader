import os
import re
import json
from docx import Document
try:
    from openai import OpenAI
except ImportError:
    print("OpenAI library not found. Please run 'pip install openai'")
    exit()

# --- AI Configuration ---
# This script is configured to use a local Ollama model.
# Make sure the Ollama application is running on your Mac.
client = OpenAI(
    base_url='http://localhost:11434/v1',
    api_key='ollama'  # required, but unused by Ollama
)
AI_MODEL = "phi3:mini" # Using a lightweight model for speed.

# ---------- RULE-BASED HEURISTICS (for speed) ----------

def is_question_block(block: str) -> bool:
    """
    A quick check to see if a block might be a question.
    This is used to prevent instructions from being misclassified as questions.
    """
    t = block.strip()
    if not t: return False
    first = t.splitlines()[0].strip()
    # Numbering patterns
    if re.match(r'^(\d+(?:\.\d+)*[.)]|\d+\.)\s+', first): return True
    if re.match(r'^([a-z]\)|\([a-z]\)|[ivxlcdm]+\))\s+', first, re.IGNORECASE): return True
    # Question mark or directive verbs
    if '?' in t: return True
    if re.search(r'\b(list|describe|name|state|explain|identify|give|provide|calculate|fill in|match)\b', first.lower()): return True
    return False

def is_instruction(block: str) -> bool:
    """
    Fast, rule-based check for instruction text.
    """
    t = block.strip()
    if not t or is_question_block(t):
        return False
    
    low = t.lower()
    # Keywords that strongly indicate instructions
    admin_keywords = [
        r'\bassessor', r'\bparticipant', r'\bmarking', r'\bcriteria',
        r'\bmodel answer', r'\bsatisfactory response', r'\btechnically correct\b',
        r'\bacceptable responses\b', r'\bmust be completed\b', r'trainer.*assessor.*instruction'
    ]
    if any(re.search(p, low) for p in admin_keywords):
        return True
    
    return False

# ---------- AI CLASSIFICATION (for accuracy) ----------

def classify_remaining_with_ai(text_block: str):
    """
    Uses AI to classify text that is NOT an instruction.
    """
    system_prompt = """
    You are an expert document analysis assistant. Your task is to classify a block of text that has already been determined NOT to be an instruction.
    Classify it and extract its contents into a structured JSON format.

    The possible classifications for the 'type' field are:
    1.  'heading': A major section title (e.g., "Part A – Emergency preparedness").
    2.  'question_with_answers': A block containing BOTH a question and its corresponding answers.
    3.  'other': Any text that doesn't fit the above categories (e.g., page footers).

    You MUST return a single JSON object with two keys: 'type' and 'content'.
    - For 'heading' or 'other', the 'content' should be a single string.
    - For 'question_with_answers', the 'content' MUST be a JSON object with two keys: 'question' (a string) and 'answers' (an array of strings).
    """

    try:
        print(f"\nAI_INFO: Analyzing block: \"{text_block[:85].replace(os.linesep, ' ')}...\"")
        response = client.chat.completions.create(
            model=AI_MODEL,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": text_block}
            ],
            response_format={"type": "json_object"},
            temperature=0.0,
        )
        result_str = response.choices[0].message.content
        result_json = json.loads(result_str)
        
        if 'type' in result_json and 'content' in result_json:
            print(f"AI_INFO: Classified as '{result_json['type']}'.")
            return result_json
        else:
            print("AI_WARNING: AI response was not in the expected format.")
            return None
    except Exception as e:
        print(f"AI_ERROR: An error occurred during AI analysis: {e}")
        print("AI_HINT: Is the Ollama application running on your Mac?")
        return None

# ---------- DOCUMENT PROCESSING (Hybrid Logic) ----------

def create_logical_blocks(doc):
    """
    Groups paragraphs and table content into logical blocks, preserving document order.
    """
    blocks = []
    current_block_lines = []

    def flush_current_block():
        """Helper to add the current block to the list and reset."""
        if current_block_lines:
            blocks.append("\n".join(current_block_lines))
            current_block_lines.clear()

    # Create mappings from the underlying XML element to the python-docx object.
    # This is an efficient way to look up the wrapper object.
    para_map = {p._p: p for p in doc.paragraphs}
    table_map = {t._tbl: t for t in doc.tables}

    # Iterate through the direct children of the document's body element.
    # This is the crucial step to preserve the order of paragraphs and tables.
    for body_child_element in doc.element.body:
        # Check if the element is a paragraph
        if body_child_element.tag.endswith('p'):
            para = para_map.get(body_child_element)
            if para:
                text = para.text.strip()
                if not text:
                    continue
                
                # Heuristic: Start a new block for major sections or new numbered items.
                is_new_block_start = re.match(r'^(Part\s+[A-Z]|Section\s+[A-Z]|\d+\.\s)', text)
                if is_new_block_start and current_block_lines:
                    flush_current_block()
                
                current_block_lines.append(text)

        # Check if the element is a table
        elif body_child_element.tag.endswith('tbl'):
            table = table_map.get(body_child_element)
            if table:
                # A table is a distinct logical unit. Flush any preceding paragraph block.
                flush_current_block()
                
                # Extract all text from the table as a single, self-contained block.
                table_lines = []
                for row in table.rows:
                    for cell in row.cells:
                        for para_in_cell in cell.paragraphs:
                            cell_text = para_in_cell.text.strip()
                            if cell_text:
                                table_lines.append(cell_text)
                if table_lines:
                    blocks.append("\n".join(table_lines))

    # Add any remaining lines from the very last block in the document.
    flush_current_block()
    return blocks

def process_document_hybrid(docx_path):
    """
    Processes the document using a hybrid rule-based and AI approach.
    """
    try:
        doc = Document(docx_path)
    except Exception as e:
        print(f"ERROR: Could not open or read DOCX file: {e}")
        return [], [], []

    headings, instructions, questions = [], [], []
    logical_blocks = create_logical_blocks(doc)

    for block_text in logical_blocks:
        # Fast path: Use rules to identify instructions first.
        if is_instruction(block_text):
            print(f"\nRULE_INFO: Classified as 'instruction'.")
            instructions.append(block_text)
            continue

        # Slow path: If not an instruction, use AI for accurate classification.
        ai_result = classify_remaining_with_ai(block_text)
        if not ai_result: continue

        result_type = ai_result.get('type')
        content = ai_result.get('content')

        if result_type == 'heading':
            headings.append(content)
        elif result_type == 'question_with_answers':
            if isinstance(content, dict) and 'question' in content and 'answers' in content:
                questions.append(content)
            else:
                print(f"AI_WARNING: 'question_with_answers' block has malformed content: {content}")

    return headings, instructions, questions

# ---------- MAIN EXECUTION ----------

def main():
    doc_filename = input("Enter DOCX filename (just file, no path): ").strip()
    path_in_cwd = os.path.join(os.getcwd(), doc_filename)
    path_in_samples = os.path.join(os.getcwd(), "samples", doc_filename)

    if os.path.isfile(path_in_cwd):
        path = path_in_cwd
    elif os.path.isfile(path_in_samples):
        path = path_in_samples
    else:
        print(f"ERROR: File '{doc_filename}' not found in current directory or in a 'samples' subdirectory.")
        return

    print(f"\nProcessing '{os.path.basename(path)}' using hybrid AI model: '{AI_MODEL}'...")
    
    headings, instructions, questions = process_document_hybrid(path)

    print("\n" + "=" * 80)
    print("QUESTIONS WITH ANSWERS")
    print("=" * 80)
    if not questions:
        print("No questions were identified.")
    for i, item in enumerate(questions, 1):
        print(f"\n{i}. QUESTION:")
        print(f"   {item['question']}")
        if item['answers']:
            print(f"\n   ANSWERS:")
            for j, ans in enumerate(item['answers'], 1):
                if ans.strip(): print(f"   {j}) {ans.strip()}")
        else:
            print(f"\n   ⚠️  WARNING: No answers were extracted for this question!")
        print("-" * 80)

    print("\n" + "=" * 80)
    print("INSTRUCTIONS")
    print("=" * 80)
    if not instructions:
        print("No instructions were identified.")
    for i, ins in enumerate(instructions, 1):
        print(f"{i}. {ins}\n")

    print("\n" + "=" * 80)
    print("HEADINGS")
    print("=" * 80)
    if not headings:
        print("No headings were identified.")
    for i, h in enumerate(headings, 1):
        print(f"{i}. {h}\n")
    
    total_questions = len(questions)
    questions_with_answers = sum(1 for q in questions if q['answers'])
    print("\n" + "=" * 80)
    print("SUMMARY")
    print("=" * 80)
    print(f"Total Questions Identified: {total_questions}")
    print(f"Questions with Answers: {questions_with_answers}")
    print(f"Questions WITHOUT Answers: {total_questions - questions_with_answers}")

if __name__ == "__main__":
    main()
