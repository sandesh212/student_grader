import os
import re
from docx import Document
try:
    from openai import OpenAI
except ImportError:
    print("ERROR: OpenAI library not found. Run: pip install openai")
    exit()

# --- AI Configuration for Smart Matching ---
client = OpenAI(
    base_url='http://localhost:11434/v1',
    api_key='ollama'
)
AI_MODEL = "phi3:mini"

# ---------- UNIT CACHE LOADER ----------

def load_all_units(cache_dir="unit_cache"):
    """Load all unit information from cached files."""
    units = {}
    if not os.path.exists(cache_dir):
        return units
    
    for filename in os.listdir(cache_dir):
        if filename.endswith('.txt'):
            unit_code = filename.replace('.txt', '')
            filepath = os.path.join(cache_dir, filename)
            try:
                with open(filepath, 'r', encoding='utf-8') as f:
                    content = f.read()
                    units[unit_code] = {
                        'code': unit_code,
                        'title': extract_unit_title(content),
                        'content_sample': content[:2000]  # First 2000 chars for context
                    }
            except Exception as e:
                pass
    
    print(f"Loaded {len(units)} units from cache.")
    return units

def extract_unit_title(content):
    """Extract unit title from various formats."""
    patterns = [
        r'"title":"([^"]+)"',
        r'<title>([^<]+)</title>',
        r'Unit of Competency[:\s]+([^\n]+)',
    ]
    for pattern in patterns:
        match = re.search(pattern, content)
        if match:
            return match.group(1).strip()
    return ''

# ---------- QUESTION & ANSWER EXTRACTION ----------

def is_red_text(run):
    """Check if text is red colored."""
    if run.font.color and run.font.color.rgb:
        r, g, b = run.font.color.rgb
        return r > 150 and g < 100 and b < 100
    return False

def extract_from_document(doc):
    """Extract questions and answers from document."""
    results = []
    
    for table in doc.tables:
        for row in table.rows:
            if len(row.cells) == 0:
                continue
            
            question_cell = row.cells[0]
            question_text = question_cell.text.strip()
            
            if not question_text or len(question_text) < 20:
                continue
            
            # Skip headers
            skip_headers = ['S', 'NS', 'Emergency Preparedness', 'Vessel Handling', 
                          'Participant Name', 'RTO Name', 'Course Name', 'UoC']
            if any(header in question_text for header in skip_headers):
                continue
            
            # Check if this looks like a question
            has_question_words = bool(re.search(
                r'\b(list|describe|name|state|explain|identify|give|provide|what|how|why|which|match)\b', 
                question_text.lower()
            ))
            has_question_mark = '?' in question_text
            has_blanks = '_' in question_text
            
            if has_question_words or has_question_mark or has_blanks:
                # Extract answers
                answers = []
                for para in question_cell.paragraphs:
                    for run in para.runs:
                        if run.text.strip() and is_red_text(run):
                            answer_text = run.text.strip()
                            parts = re.split(r'\n|;(?!\w)', answer_text)
                            for part in parts:
                                part = part.strip()
                                if part and len(part) > 2 and not part.startswith('_'):
                                    answers.append(part)
                
                # Clean question
                question_clean = re.sub(r'_{3,}', '[___]', question_text)
                question_clean = re.sub(r'\s+', ' ', question_clean).strip()
                
                # Extract explicit unit codes
                explicit_codes = extract_explicit_unit_codes(question_text)
                
                results.append({
                    'question': question_clean,
                    'answers': answers,
                    'explicit_codes': explicit_codes
                })
    
    return results

def extract_explicit_unit_codes(text):
    """Extract unit codes mentioned in text (K007 -> MARK007)."""
    codes = re.findall(r'\b([A-Z]\d{3})[:\-]', text)
    full_codes = ["MAR" + code for code in codes]
    return list(set(full_codes))

# ---------- SMART AI-POWERED UNIT MATCHING ----------

def smart_match_to_units(question_data, all_units, use_ai=True):
    """
    Smart matching that understands meaning, not just keywords.
    Uses AI to understand semantic similarity.
    """
    question = question_data['question']
    answers = question_data['answers']
    explicit_codes = question_data['explicit_codes']
    
    # FAST PATH: If question explicitly mentions units, return those immediately
    if explicit_codes:
        matched = []
        for code in explicit_codes:
            if code in all_units:
                matched.append({
                    'unit_code': code,
                    'unit_title': all_units[code]['title'],
                    'score': 200,
                    'match_type': 'Explicit reference',
                    'confidence': 'Very High'
                })
        if matched:
            return matched
    
    # SMART PATH: Use AI to understand semantic meaning
    if not use_ai or len(all_units) == 0:
        return []
    
    try:
        # Build context for AI
        full_context = f"Question: {question}\nAnswers: {', '.join(answers)}"
        
        # Create a condensed list of MARITIME units for AI to analyze
        # Prioritize MAR* units (maritime) over other training packages
        maritime_units = {k: v for k, v in all_units.items() if k.startswith('MAR')}
        other_units = {k: v for k, v in all_units.items() if not k.startswith('MAR')}
        
        unit_list = []
        # Add maritime units first (most relevant)
        for code, data in list(maritime_units.items())[:40]:
            title = data['title'] if data['title'] else 'Maritime unit'
            unit_list.append(f"{code}: {title}")
        # Add a few other units
        for code, data in list(other_units.items())[:10]:
            title = data['title'] if data['title'] else 'Training unit'
            unit_list.append(f"{code}: {title}")
        
        units_text = "\n".join(unit_list)
        
        # Extract key concepts from question
        key_words = []
        if 'overboard' in full_context.lower() or 'falls' in full_context.lower():
            key_words.append('person overboard/emergency procedures')
        if 'engine' in full_context.lower() or 'motor' in full_context.lower():
            key_words.append('engine operation/maintenance')
        if 'weather' in full_context.lower() or 'storm' in full_context.lower():
            key_words.append('weather/meteorology')
        if 'fuel' in full_context.lower() or 'refuel' in full_context.lower():
            key_words.append('fuel safety/refueling')
        if 'vessel' in full_context.lower() or 'boat' in full_context.lower():
            key_words.append('vessel operations/handling')
        
        concepts_hint = f"\nKey concepts detected: {', '.join(key_words)}" if key_words else ""
        
        prompt = f"""You are an expert in Australian training.gov.au maritime competency units.

QUESTION AND ANSWERS:
{full_context}{concepts_hint}

AVAILABLE MARITIME UNITS (prioritized):
{units_text}

Common Maritime Units:
- MARK007: Handle a vessel up to 12 metres
- MARN008: Apply seamanship skills aboard a vessel
- MARI003: Comply with regulations to ensure safe operation
- MARC037: Perform basic servicing and maintenance on an engine

Task: Match this question to 1-3 units that best assess these competencies.
Consider MEANING not just keywords:
- "person falls overboard" = MARK007 or MARI003 (emergency/safety)
- "check engine" = MARC037 (engine maintenance)  
- "weather signs" = MARK007 (vessel operations/weather)
- "refuel safely" = MARN008 or MARC037 (seamanship/safety)

Return ONLY valid JSON:
{{"matches": [{{"code": "MARK007", "reason": "Emergency procedures", "confidence": "high"}}]}}

If uncertain, return: {{"matches": []}}"""

        response = client.chat.completions.create(
            model=AI_MODEL,
            messages=[
                {"role": "system", "content": "You are a competency mapping expert. Return only valid JSON."},
                {"role": "user", "content": prompt}
            ],
            response_format={"type": "json_object"},
            temperature=0.1,
            max_tokens=300
        )
        
        result = response.choices[0].message.content
        import json
        data = json.loads(result)
        
        matches = []
        for match in data.get('matches', []):
            code = match.get('code')
            if code in all_units:
                matches.append({
                    'unit_code': code,
                    'unit_title': all_units[code]['title'],
                    'score': 150 if match.get('confidence') == 'high' else 100,
                    'match_type': f"AI: {match.get('reason', 'Semantic match')}",
                    'confidence': match.get('confidence', 'medium').title()
                })
        
        return matches[:3]
        
    except Exception as e:
        print(f"AI matching error: {e}")
        return []

# ---------- MAIN PROCESSING ----------

def process_pair(question_file, answer_file, all_units, use_ai=True):
    """Process question file with smart AI matching."""
    print(f"\n{'='*100}")
    print(f"PROCESSING: {os.path.basename(question_file)}")
    print(f"{'='*100}\n")
    
    try:
        q_doc = Document(question_file)
        questions = extract_from_document(q_doc)
    except Exception as e:
        print(f"ERROR: {e}")
        return
    
    # Load answers from marking sheet if exists
    answers_from_marking = {}
    if answer_file and os.path.exists(answer_file):
        try:
            a_doc = Document(answer_file)
            answer_data = extract_from_document(a_doc)
            for i, item in enumerate(answer_data, 1):
                if item['answers']:
                    answers_from_marking[i] = item['answers']
        except:
            pass
    
    # Process each question with SMART matching
    for i, q_item in enumerate(questions, 1):
        print(f"\n{'─'*100}")
        print(f"Q{i}. {q_item['question'][:150]}")
        
        # Answers
        answers = q_item['answers'] or answers_from_marking.get(i, [])
        if answers:
            print(f"\n   ✓ ANSWERS ({len(answers)}):")
            for j, ans in enumerate(answers[:3], 1):  # Show first 3
                print(f"     {j}) {ans[:100]}")
            if len(answers) > 3:
                print(f"     ... and {len(answers)-3} more")
        else:
            print(f"\n   ⚠️  No answers found")
        
        # SMART MATCHING - AI understands meaning
        print(f"\n   🧠 AI SEMANTIC MATCHING...")
        matches = smart_match_to_units(q_item, all_units, use_ai=use_ai)
        
        if matches:
            print(f"\n   🎯 MATCHED UNITS:")
            for match in matches:
                print(f"      • {match['unit_code']}: {match['unit_title'][:60]}")
                print(f"        └─ {match['match_type']}")
                print(f"           Confidence: {match['confidence']} (Score: {match['score']})")
        else:
            print(f"\n   ⚠️  No confident matches found")
    
    print(f"\n{'='*100}")
    print(f"SUMMARY: {len(questions)} questions processed")
    print(f"{'='*100}\n")

# ---------- MAIN ----------

def main():
    print("🧠 Smart Grader - AI-Powered Semantic Matching")
    print("=" * 60)
    
    # Check if Ollama is running
    print("\nChecking AI availability...")
    try:
        test_response = client.chat.completions.create(
            model=AI_MODEL,
            messages=[{"role": "user", "content": "test"}],
            max_tokens=5
        )
        print("✅ AI is ready (Ollama connected)")
        use_ai = True
    except Exception as e:
        print(f"⚠️  AI not available: {e}")
        print("   Falling back to keyword matching only")
        use_ai = False
    
    # Load units
    all_units = load_all_units()
    if not all_units:
        print("ERROR: No units loaded.")
        return
    
    # Get file
    filename = input("\nEnter DOCX filename (or press Enter for all): ").strip()
    
    if not filename:
        # Process all
        samples_dir = "samples"
        if not os.path.exists(samples_dir):
            print(f"ERROR: {samples_dir} not found.")
            return
        
        files = [f for f in os.listdir(samples_dir) 
                if f.endswith('.docx') and not f.startswith('~$') and 'Marking' not in f]
        
        for file in files:
            question_path = os.path.join(samples_dir, file)
            
            # Find marking sheet
            base_name = file.replace('.docx', '')
            answer_path = None
            for suffix in ['-Marking-Sheet', ' Marking Sheet']:
                ans = os.path.join(samples_dir, f"{base_name}{suffix}.docx")
                if os.path.exists(ans):
                    answer_path = ans
                    break
            
            process_pair(question_path, answer_path, all_units, use_ai)
    else:
        # Process single file
        path_in_cwd = os.path.join(os.getcwd(), filename)
        path_in_samples = os.path.join(os.getcwd(), "samples", filename)
        
        if os.path.isfile(path_in_cwd):
            path = path_in_cwd
        elif os.path.isfile(path_in_samples):
            path = path_in_samples
        else:
            print(f"ERROR: File not found.")
            return
        
        # Find marking sheet
        base_name = os.path.basename(path).replace('.docx', '')
        dir_name = os.path.dirname(path)
        answer_path = None
        
        for suffix in ['-Marking-Sheet', ' Marking Sheet']:
            ans_path = os.path.join(dir_name, f"{base_name}{suffix}.docx")
            if os.path.exists(ans_path):
                answer_path = ans_path
                break
        
        process_pair(path, answer_path, all_units, use_ai)

if __name__ == "__main__":
    main()
