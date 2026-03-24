#!/usr/bin/env python
"""Quick test of sd_chapter_classifier fixes"""
import json
import logging

logging.basicConfig(level=logging.DEBUG)

# Simulate what the classifier does
ALLOWED_GROUPS = {
    "Executive Summary & Product Overview",
    "Scope Boundaries & Prerequisites",
    "Transition Operations & Governance",
    "Commercial & Risk Management",
    "Internal Presales Alignment",
}

# Test case 1: With numbering (what Ollama actually returns)
test_json_1 = '{"group": "1. Executive Summary & Product Overview", "reason": "This is the overview"}'

# Test case 2: Without numbering  
test_json_2 = '{"group": "Executive Summary & Product Overview", "reason": "This is the overview"}'

print("=== TEST 1: With numbering (Ollama output) ===")
try:
    parsed = json.loads(test_json_1)
    group = str(parsed.get("group", "")).strip()
    
    print(f"Raw group: '{group}'")
    print(f"In ALLOWED_GROUPS: {group in ALLOWED_GROUPS}")
    
    # Apply fix: remove numbering
    if group and group[0].isdigit():
        parts = group.split(". ", 1)
        if len(parts) > 1:
            group = parts[1]
    
    print(f"Fixed group: '{group}'")
    print(f"In ALLOWED_GROUPS: {group in ALLOWED_GROUPS}")
    print(f"✓ SUCCESS" if group in ALLOWED_GROUPS else "✗ FAILED")
except Exception as e:
    print(f"✗ ERROR: {e}")

print("\n=== TEST 2: Without numbering ===")
try:
    parsed = json.loads(test_json_2)
    group = str(parsed.get("group", "")).strip()
    
    print(f"Raw group: '{group}'")
    print(f"In ALLOWED_GROUPS: {group in ALLOWED_GROUPS}")
    print(f"✓ SUCCESS" if group in ALLOWED_GROUPS else "✗ FAILED")
except Exception as e:
    print(f"✗ ERROR: {e}")
