import os
from validator_enhanced import ReferenceValidator

# File mentioned by user
file_path = r"C:\Users\harikrishnam\Desktop\Citation reordering\S4C-Processed-Documents\Abuhamad9781975242831-ch002.docx"
output_path = r"C:\Users\harikrishnam\Desktop\Citation reordering\S4C-Processed-Documents\Abuhamad9781975242831-ch002-renumbered.docx"

if not os.path.exists(file_path):
    print(f"File not found: {file_path}")
else:
    print(f"Processing {file_path}...")
    print("Enhanced validator with physical reference reordering enabled.\n")
    try:
        validator = ReferenceValidator(file_path)
        with validator:
            # Enable auto-renumbering
            results = validator.validate(auto_renumber=True, save_path=output_path)
        
        print("="*60)
        print("VALIDATION REPORT")
        print("="*60)
        print(f"Total References: {results['total_references']}")
        print(f"Total Citations: {results['total_citations']}")
        print(f"\nSequence Status: {results.get('sequence_message', 'Unknown')}")
        
        if results.get('missing_references'):
            print(f"\nMissing References: {sorted(results['missing_references'])}")
        
        if results.get('unused_references'):
            print(f"Unused References: {sorted(results['unused_references'])}")
        
        print(f"\nRenumber Status: {results.get('renumber_attempt', {}).get('renumbered', False)}")
        
        if results.get('renumber_attempt', {}).get('renumbered'):
            print(f"\n✓ References have been PHYSICALLY REORDERED to match citation order")
            print(f"✓ Renumbered file saved to: {output_path}")
            
            renumber_map = results['renumber_attempt'].get('map', {})
            if renumber_map:
                print(f"\nRenumber Map (Old → New):")
                for old, new in sorted(renumber_map.items()):
                    print(f"  {old} → {new}")
        
        print("\n" + "="*60)
        
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
