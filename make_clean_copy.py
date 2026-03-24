#!/usr/bin/env python3
"""
Create clean SDT by copying filled version and replacing original
"""
import shutil
import os

source_filled = 'SDT - Cegeka DBMS for Oracle on Azure [PRD.0.8.001][PV1.0][DV1.0]_FILLED.docx'
target = 'SDT - Cegeka DBMS for Oracle on Azure [PRD.0.8.001][PV1.0][DV1.0].docx'

print(f"Preparing clean copy...")

# Remove old file if exists
if os.path.exists(target):
    try:
        os.remove(target)
        print(f"✓ Removed old file")
    except Exception as e:
        print(f"⚠ Could not remove old file: {e}")

# Copy filled version to main name
try:
    shutil.copy2(source_filled, target)
    print(f"✓ Copied {source_filled}")
    print(f"  → {target}")
    
    # Verify
    if os.path.exists(target):
        size = os.path.getsize(target)
        print(f"✓ File created: {size} bytes")
        
        # Quick validation
        from docx import Document
        doc = Document(target)
        sdts = doc.element.body.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sdt')
        print(f"✓ Validated: {len(sdts)} SDT fields, {len(doc.paragraphs)} paragraphs")
        print("\n✓ SUCCESS: Clean SDT file is ready to open!")
    
except Exception as e:
    print(f"✗ Error: {e}")
    import traceback
    traceback.print_exc()
