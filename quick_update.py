#!/usr/bin/env python3
"""
Simple script to process your Excel data and update KML
Just run this script after filling your template
"""
from datetime import date
from main import KMLPolygonEditor
import os

def quick_process():
    """Quick process using the standard template"""
    
    excel_file = 'main-data.xlsx'
    kml_file = 'MyArea.kml'
    
    print("Quick KML Update from Excel Template")
    print("=" * 40)
    
    # Check if files exist
    if not os.path.exists(excel_file):
        print(f"❌ Template file not found: {excel_file}")
        print("Run 'python create_template.py' first to create the template.")
        return False
    
    if not os.path.exists(kml_file):
        print(f"❌ KML file not found: {kml_file}")
        return False
    
    # Initialize editor
    editor = KMLPolygonEditor(kml_file)
    
    # Apply updates
    print("Applying updates from template...")
    summary = editor.update_polygons_from_excel(
        excel_file_path=excel_file,
        polygon_column='اسم/ رقم القطاع:',
        image_columns=['صورة للزاوية 1:_URL', 'صورة للزاوية 2:_URL', 'صورة للزاوية 3:_URL','صورة للزاوية 4:_URL'],
        description_columns=['اسم جامع البيانات:', 'نوع البناء:', 'ما هو عدد الشقق في البناء:'],
        merge_with_existing=True
    )
    
    if summary["success"]:
        print(f"✅ Success! Updated {len(summary['updated_polygons'])} polygons")
        print(f"   Images added: {summary['total_images_added']}")
        print(f"   Descriptions added: {summary['total_descriptions_added']}")
        # Save result
        output_file = f"{kml_file.replace('.kml', '')}_updated.kml"
        editor.save_kml(output_file)
        print(f"   Saved to: {output_file}")
        
        if summary['not_found_polygons']:
            print(f"⚠️  {len(summary['not_found_polygons'])} polygons not found in KML")
        
        return True
    else:
        print(f"❌ Failed: {summary['message']}")
        return False

if __name__ == "__main__":
    quick_process()
