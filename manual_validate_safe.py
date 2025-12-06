import os
from validator_safe_renumber import ReferenceValidator

# Your file path
file_path = r"C:\Users\harikrishnam\Desktop\Citation reordering\S4C-Processed-Documents\Abuhamad9781975242831-ch002.docx"
output_path = r"C:\Users\harikrishnam\Desktop\Citation reordering\S4C-Processed-Documents\Abuhamad9781975242831-ch002-renumbered-SAFE.docx"

if not os.path.exists(file_path):
    print(f"File not found: {file_path}")
else:
    print(f"Processing {file_path}...")
    print("SAFE Validator (only updates numbers, preserves all references)\n")
    try:
        validator = ReferenceValidator(file_path)
        with validator:
            print("Step 1: Analyzing document...")
            results = validator.validate(auto_renumber=False)
            
            print(f"  Citation Format Detected: {results.get('citation_format_detected', 'Unknown')}")
            print(f"  Total References Found: {results['total_references']}")
            print(f"  Total Citations Found: {results['total_citations']}")
            print(f"  Sequence Status: {results.get('sequence_message', 'Unknown')}")
            
            if results.get('citation_sequence'):
                print(f"  Citation Sequence (first 20): {results['citation_sequence'][:20]}")
            
            # Only renumber if needed
            if 'NOT in sequence' in results.get('sequence_message', ''):
                print("\nStep 2: Renumbering citations (SAFE METHOD)...")
                results = validator.validate(auto_renumber=True, save_path=output_path)
                
                print("\n" + "="*70)
                print("RENUMBERING RESULTS")
                print("="*70)
                
                ren_result = results.get('renumber_attempt', {})
                if ren_result.get('renumbered'):
                    print(f"✅ SUCCESS: Citations renumbered")
                    print(f"  Format used: {ren_result.get('format')}")
                    print(f"  Output saved to: {output_path}")
                    
                    renumber_map = ren_result.get('map', {})
                    print(f"\n  Mapping applied (Old -> New):")
                    for i, (old, new) in enumerate(sorted(renumber_map.items())[:15]):
                        print(f"    {old:2d} -> {new:2d}")
                    if len(renumber_map) > 15:
                        print(f"    ... and {len(renumber_map) - 15} more mappings")
                else:
                    print(f"❌ FAILED: {ren_result.get('message')}")
            else:
                print("\nStep 2: Skipping renumbering...")
                print("  Reason: Citations already in sequence")
            
            print("\n" + "="*70)
            print("FINAL VALIDATION")
            print("="*70)
            print(f"Total References in Document: {results['total_references']}")
            print(f"Total Citations in Document: {results['total_citations']}")
            
            if results.get('missing_references'):
                print(f"\n⚠️  Missing References: {sorted(results['missing_references'])}")
            
            if results.get('unused_references'):
                print(f"\n⚠️  Unused References: {sorted(results['unused_references'])}")
            
            if not results.get('missing_references') and not results.get('unused_references'):
                print("\n✅ All references are cited!")
            
            print("\n" + "="*70)
        
    except Exception as e:
        print(f"❌ ERROR: {e}")
        import traceback
        traceback.print_exc()
