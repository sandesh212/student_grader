import os
import re
from docx import Document
from docx.shared import RGBColor

# ---------- UNIT CACHE LOADER ----------

def load_all_units(cache_dir="unit_cache"):
    """Load all unit information from cached files."""
    units = {}
    if not os.path.exists(cache_dir):
        print(f"WARNING: {cache_dir} directory not found.")
        return units
    
    for filename in os.listdir(cache_dir):
        if filename.endswith('.txt'):
            unit_code = filename.replace('.txt', '')
            filepath = os.path.join(cache_dir, filename)
            try:
                with open(filepath, 'r', encoding='utf-8') as f:
                    content = f.read()
                    units[unit_code] = parse_unit_content(content, unit_code)
            except Exception as e:
                print(f"WARNING: Could not load {filename}: {e}")
    
    print(f"Loaded {len(units)} units from cache.")
    return units

def parse_unit_content(content, unit_code):
    """Parse unit content to extract PC and KE items."""
    unit_data = {
        'code': unit_code,
        'title': '',
        'performance_criteria': [],
        'knowledge_evidence': [],
        'full_text': content.lower()  # For fuzzy matching
    }
    
    # Extract title (usually first few lines)
    lines = content.split('\n')
    for line in lines[:5]:
        if line.strip() and len(line.strip()) > 10:
            unit_data['title'] = line.strip()
            break
    
    # Extract Performance Criteria (PC)
    pc_pattern = r'(\d+\.\d+)\s+(.+?)(?=\d+\.\d+|Knowledge Evidence|Performance Evidence|$)'
    pc_matches = re.findall(pc_pattern, content, re.DOTALL)
    for pc_num, pc_text in pc_matches:
        unit_data['performance_criteria'].append({
            'number': pc_num,
            'text': pc_text.strip()[:200]  # First 200 chars
        })
    
    # Extract Knowledge Evidence items
    ke_section = re.search(r'Knowledge Evidence(.+?)(?=Performance Evidence|Assessment Conditions|$)', 
                          content, re.DOTALL | re.IGNORECASE)
    if ke_section:
        ke_content = ke_section.group(1)
        # Look for bullet points or numbered items
        ke_items = re.findall(r'(?:^|\n)\s*[•\-\*]\s*(.+?)(?=\n\s*[•\-\*]|$)', ke_content, re.DOTALL)
        unit_data['knowledge_evidence'] = [item.strip()[:200] for item in ke_items if len(item.strip()) > 20]
    
    return unit_data

# ---------- FAST COLOR-BASED EXTRACTION ----------

def is_red_text(run):
    """Check if text is red colored (answers)."""
    if run.font.color and run.font.color.rgb:
        r, g, b = run.font.color.rgb
        return r > 150 and g < 100 and b < 100
    return False

def extract_questions_and_answers_fast(doc):
    """Fast extraction from tables - works with your document format."""
    results = []
    
    for table in doc.tables:
        for row in table.rows:
            # Usually first cell contains the question
            if len(row.cells) == 0:
                continue
            
            question_cell = row.cells[0]
            question_text = question_cell.text.strip()
            
            # Skip header rows and empty cells
            if not question_text or len(question_text) < 20:
                continue
            
            # Skip if it looks like a header or section title (short, all caps, etc.)
            if (question_text.isupper() and len(question_text) < 50) or \
               question_text in ['S', 'NS', 'Emergency Preparedness', 'Vessel Handling and Watertight Integrity']:
                continue
            
            # Check if this looks like a question
            has_question_words = bool(re.search(
                r'\b(list|describe|name|state|explain|identify|give|provide|what|how|why|which)\b', 
                question_text.lower()
            ))
            has_question_mark = '?' in question_text
            has_blanks = '_' in question_text or 'action' in question_text.lower()
            
            if has_question_words or has_question_mark or has_blanks:
                # Extract answers from colored text (red = answers)
                answers = []
                for para in question_cell.paragraphs:
                    for run in para.runs:
                        if run.text.strip() and is_red_text(run):
                            answer_text = run.text.strip()
                            # Split multi-part answers
                            parts = re.split(r'\n|;|\|', answer_text)
                            for part in parts:
                                part = part.strip()
                                if part and len(part) > 2 and not part.startswith('_'):
                                    answers.append(part)
                
                # Clean up the question text
                question_clean = re.sub(r'_{3,}', '[BLANK]', question_text)
                question_clean = re.sub(r'\s+', ' ', question_clean).strip()
                
                results.append({
                    'question': question_clean,
                    'answers': answers
                })
    
    return results

# ---------- FAST UNIT MATCHING ----------

def match_question_to_units(question_text, answer_texts, all_units):
    """Match a question to relevant units and their PC/KE items."""
    matches = []
    
    # Combine question and answers for context
    full_context = (question_text + ' ' + ' '.join(answer_texts)).lower()
    
    # Extract explicit unit references from question (e.g., "K007-PC1.1" or "N008:K9")
    explicit_units = re.findall(r'\b([A-Z]\d{3})[:\-]', question_text)
    
    # Map abbreviated codes to full unit codes (K007 -> MARK007, N008 -> MARN008, etc.)
    # The pattern is: first letter + "MAR" + rest of the code
    # So K007 -> MAR + K + 007 = MARK007
    
    # Expand abbreviated codes
    full_unit_codes = []
    for abbr in explicit_units:
        # abbr is like "K007", "N008", "I003"
        # We need to build "MARK007", "MARN008", "MARI003"
        full_code = "MAR" + abbr  # K007 -> MARK007, N008 -> MARN008
        full_unit_codes.append(full_code)
    
    # Extract key terms (nouns, verbs) for semantic matching
    key_terms = set(re.findall(r'\b[a-z]{4,}\b', full_context))
    
    for unit_code, unit_data in all_units.items():
        score = 0
        matched_items = []
        
        # HIGH PRIORITY: Check if this unit was explicitly mentioned
        if unit_code in full_unit_codes:
            score += 200  # Very high confidence
            matched_items.append("Explicitly mentioned")
        
        # REMOVED: Don't match on partial code endings - too many false positives
        
        # Check if unit code is mentioned anywhere
        if unit_code.lower() in full_context:
            score += 100
        
        # Check title match
        title_words = set(re.findall(r'\b[a-z]{4,}\b', unit_data['title'].lower()))
        common_title_words = key_terms & title_words
        score += len(common_title_words) * 10
        
        # Check PC matches
        for pc in unit_data['performance_criteria']:
            pc_words = set(re.findall(r'\b[a-z]{4,}\b', pc['text'].lower()))
            common_pc = key_terms & pc_words
            if len(common_pc) >= 3:
                score += len(common_pc) * 5
                matched_items.append(f"PC {pc['number']}")
        
        # Check KE matches
        for ke in unit_data['knowledge_evidence']:
            ke_words = set(re.findall(r'\b[a-z]{4,}\b', ke.lower()))
            common_ke = key_terms & ke_words
            if len(common_ke) >= 3:
                score += len(common_ke) * 5
                matched_items.append("KE")
        
        if score > 20:  # Threshold for relevance
            matches.append({
                'unit_code': unit_code,
                'unit_title': unit_data['title'],
                'score': score,
                'matched_items': list(set(matched_items))
            })
    
    # Sort by score and return top matches
    matches.sort(key=lambda x: x['score'], reverse=True)
    return matches[:5]  # Top 5 matches

# ---------- MAIN EXECUTION ----------

def main():
    print("Loading unit cache...")
    all_units = load_all_units()
    
    if not all_units:
        print("ERROR: No units loaded. Make sure unit_cache directory exists with .txt files.")
        return
    
    doc_filename = input("Enter DOCX filename (just file, no path): ").strip()
    path_in_cwd = os.path.join(os.getcwd(), doc_filename)
    path_in_samples = os.path.join(os.getcwd(), "samples", doc_filename)
    
    if os.path.isfile(path_in_cwd):
        path = path_in_cwd
    elif os.path.isfile(path_in_samples):
        path = path_in_samples
    else:
        print(f"ERROR: File '{doc_filename}' not found.")
        return
    
    print(f"Processing '{os.path.basename(path)}'...")
    doc = Document(path)
    
    questions = extract_questions_and_answers_fast(doc)
    
    print("\n" + "=" * 100)
    print("QUESTIONS, ANSWERS & UNIT MAPPING")
    print("=" * 100)
    
    for i, item in enumerate(questions, 1):
        print(f"\n{'─' * 100}")
        print(f"Q{i}. {item['question']}")
        
        if item['answers']:
            print(f"\n   ANSWERS:")
            for j, ans in enumerate(item['answers'], 1):
                print(f"   {j}) {ans}")
        else:
            print(f"\n   ⚠️  No answers detected")
        
        # Match to units
        matches = match_question_to_units(item['question'], item['answers'], all_units)
        
        if matches:
            print(f"\n   📚 RELATED UNITS:")
            for match in matches:
                items_str = ', '.join(match['matched_items']) if match['matched_items'] else 'General match'
                print(f"      • {match['unit_code']}: {match['unit_title'][:60]}")
                print(f"        └─ Matched: {items_str} (Score: {match['score']})")
        else:
            print(f"\n   ⚠️  No unit matches found")
    
    print("\n" + "=" * 100)
    print(f"SUMMARY: {len(questions)} questions processed")
    print("=" * 100)

if __name__ == "__main__":
    main()
