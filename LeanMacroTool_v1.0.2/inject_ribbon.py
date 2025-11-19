#!/usr/bin/env python3
"""
Inject Custom Ribbon UI into Excel .xlam file
This adds a Ribbon tab to the add-in
"""

import zipfile
import os
import sys
import shutil
from pathlib import Path

def inject_ribbon(xlam_path, customui_xml_path, rels_xml_path):
    """Inject custom ribbon XML into an Excel add-in file"""
    
    if not os.path.exists(xlam_path):
        print(f"Error: {xlam_path} not found!")
        return False
    
    if not os.path.exists(customui_xml_path):
        print(f"Error: {customui_xml_path} not found!")
        return False
        
    if not os.path.exists(rels_xml_path):
        print(f"Error: {rels_xml_path} not found!")
        return False
    
    # Create backup
    backup_path = xlam_path + ".backup"
    shutil.copy2(xlam_path, backup_path)
    print(f"✓ Created backup: {backup_path}")
    
    # Create temp directory
    temp_dir = Path(xlam_path).parent / "temp_xlam"
    if temp_dir.exists():
        shutil.rmtree(temp_dir)
    temp_dir.mkdir()
    
    try:
        # Extract xlam
        print("Extracting .xlam...")
        with zipfile.ZipFile(xlam_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
        
        # Create customUI folder
        customui_dir = temp_dir / "customUI"
        customui_dir.mkdir(exist_ok=True)
        
        # Copy customUI14.xml
        shutil.copy2(customui_xml_path, customui_dir / "customUI14.xml")
        print("✓ Added customUI/customUI14.xml")
        
        # Update _rels/.rels to reference customUI
        rels_dir = temp_dir / "_rels"
        rels_dir.mkdir(exist_ok=True)
        
        # Read existing .rels
        rels_file = rels_dir / ".rels"
        if rels_file.exists():
            with open(rels_file, 'r') as f:
                rels_content = f.read()
            
            # Check if customUI relationship already exists
            if "customUI" not in rels_content:
                # Add customUI relationship before closing </Relationships>
                customui_rel = '  <Relationship Id="rIdCustomUI" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI14.xml"/>\n'
                rels_content = rels_content.replace('</Relationships>', customui_rel + '</Relationships>')
                
                with open(rels_file, 'w') as f:
                    f.write(rels_content)
                print("✓ Updated _rels/.rels")
        else:
            # Create new .rels with customUI reference
            shutil.copy2(rels_xml_path, rels_file)
            print("✓ Created _rels/.rels")
        
        # Update [Content_Types].xml
        content_types_file = temp_dir / "[Content_Types].xml"
        if content_types_file.exists():
            with open(content_types_file, 'r') as f:
                content = f.read()
            
            # Add customUI content type if not exists
            if "customUI" not in content and "extensibility" not in content:
                override_line = '  <Override PartName="/customUI/customUI14.xml" ContentType="application/xml"/>\n'
                content = content.replace('</Types>', override_line + '</Types>')
                
                with open(content_types_file, 'w') as f:
                    f.write(content)
                print("✓ Updated [Content_Types].xml")
        
        # Repack as .xlam
        print("Repacking .xlam...")
        os.remove(xlam_path)
        
        with zipfile.ZipFile(xlam_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    file_path = Path(root) / file
                    arcname = file_path.relative_to(temp_dir)
                    zipf.write(file_path, arcname)
        
        print(f"✓ Successfully injected ribbon into {xlam_path}")
        print("\n✅ Done! Your .xlam now has a 'Lean Macros' ribbon tab!")
        print("   Reopen Excel to see it.")
        
        return True
        
    except Exception as e:
        print(f"\n❌ Error: {e}")
        # Restore backup
        if os.path.exists(backup_path):
            shutil.copy2(backup_path, xlam_path)
            print(f"Restored from backup")
        return False
        
    finally:
        # Cleanup
        if temp_dir.exists():
            shutil.rmtree(temp_dir)


if __name__ == "__main__":
    if len(sys.argv) < 4:
        print("Usage: python3 inject_ribbon.py <xlam_file> <customUI14.xml> <_rels.xml>")
        print("\nExample:")
        print("  python3 inject_ribbon.py LeanMacroTools.xlam customUI14.xml _rels_dot_rels_for_customUI.xml")
        sys.exit(1)
    
    xlam_path = sys.argv[1]
    customui_xml = sys.argv[2]
    rels_xml = sys.argv[3]
    
    success = inject_ribbon(xlam_path, customui_xml, rels_xml)
    sys.exit(0 if success else 1)
