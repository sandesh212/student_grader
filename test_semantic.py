#!/usr/bin/env python3
"""Test semantic understanding of grader_smart.py"""
from grader_smart import smart_match_to_units, load_all_units

# Load units
print("Loading units...")
all_units = load_all_units()
print(f"Loaded {len(all_units)} units\n")

# Test cases - questions with different wording but same meaning
test_questions = [
    {
        'question': 'What actions should you take if someone falls into the water from your boat?',
        'answers': ['Throw life ring', 'Keep eyes on person', 'Call for help'],
        'explicit_codes': [],
        'expected_unit': 'MARK007 or MARI003 (person overboard procedures)'
    },
    {
        'question': 'Describe the steps to check your engine before departure',
        'answers': ['Check oil level', 'Test fuel system', 'Inspect cooling'],
        'explicit_codes': [],
        'expected_unit': 'MARC037 (engine maintenance)'
    },
    {
        'question': 'How do you safely refuel a vessel?',
        'answers': ['Turn off engine', 'No smoking', 'Check for fumes'],
        'explicit_codes': [],
        'expected_unit': 'MARN008 or MARC037 (fuel safety)'
    },
    {
        'question': 'What weather signs indicate an approaching storm?',
        'answers': ['Dark clouds', 'Falling pressure', 'Wind shift'],
        'explicit_codes': [],
        'expected_unit': 'MARK007 (weather awareness)'
    }
]

print("="*80)
print("TESTING SEMANTIC UNDERSTANDING (NO EXPLICIT UNIT CODES)")
print("="*80)

for i, test in enumerate(test_questions, 1):
    print(f"\n{'─'*80}")
    print(f"TEST {i}: {test['question']}")
    print(f"Answers: {', '.join(test['answers'])}")
    print(f"Expected: {test['expected_unit']}")
    print(f"\n🧠 AI Analyzing...")
    
    matches = smart_match_to_units(test, all_units, use_ai=True)
    
    if matches:
        print(f"\n✅ MATCHED UNITS:")
        for match in matches:
            print(f"   • {match['unit_code']}: {match['unit_title']}")
            print(f"     └─ {match['match_type']}")
            print(f"        Confidence: {match['confidence']}")
    else:
        print(f"\n❌ No matches found")

print(f"\n{'='*80}")
print("SEMANTIC MATCHING TEST COMPLETE")
print("="*80)
print("\nThe AI should match questions to units based on MEANING,")
print("not just keyword matching. Even with different wording,")
print("it should identify the correct competency being assessed.")
