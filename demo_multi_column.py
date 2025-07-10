#!/usr/bin/env python3
"""
Complete example demonstrating the multi-column Excel integration
This shows the full workflow from template creation to KML update
"""

from main import KMLPolygonEditor
import pandas as pd
import os

def demonstrate_multi_column_system():
    """Demonstrate the complete multi-column Excel system"""
    
    print("🎯 Multi-Column Excel Integration Demo")
    print("=" * 50)
    
    # Step 1: Create sample data with multiple columns
    print("1. Creating sample Excel data with multiple columns...")
    
    sample_data = {
        'Polygon_Name': [
            'القطاع A',  # Will match with fuzzy matching
            'القطاع A',
            'القطاع A',
            'القطاع A',
            'New Polygon',  # This won't be found
            'New Polygon'
        ],
        'Image_URL_1': [
            'https://example.com/photo1_main.jpg',
            'https://example.com/photo2_main.jpg',
            '',  # Empty - will be ignored
            'https://example.com/photo4_main.jpg',
            'https://example.com/new_photo1.jpg',
            ''
        ],
        'Image_URL_2': [
            'https://example.com/photo1_detail.jpg',
            '',  # Empty - will be ignored
            'https://example.com/photo3_detail.jpg',
            '',
            'https://example.com/new_photo2.jpg',
            'https://example.com/new_photo3.jpg'
        ],
        'Image_URL_3': [
            '',
            'https://example.com/photo2_wide.jpg',
            'https://example.com/photo3_wide.jpg',
            '',
            '',
            ''
        ],
        'Primary_Description': [
            'Multi-column demo: Primary observation 1',
            'Multi-column demo: Primary observation 2',
            'Multi-column demo: Primary observation 3',
            'Multi-column demo: Follow-up observation',
            'This polygon does not exist in KML',
            'Another entry for non-existent polygon'
        ],
        'Secondary_Description': [
            'Additional details: Infrastructure status good',
            'Additional details: Population survey completed',
            'Additional details: Environmental assessment',
            'Additional details: Maintenance required',
            'This will not be processed',
            'This will also not be processed'
        ],
        'Field_Notes': [
            'Field conditions: Sunny, clear visibility',
            'Field conditions: Overcast, moderate visibility',
            'Field conditions: Clear, excellent visibility',
            'Field conditions: Light rain, limited visibility',
            'Field conditions: N/A',
            'Field conditions: N/A'
        ],
        'Date': [
            '2025-07-01',
            '2025-07-02',
            '2025-07-03',
            '2025-07-10',
            '2025-07-05',
            '2025-07-06'
        ],
        'Team': [
            'Alpha Team',
            'Beta Team',
            'Alpha Team',
            'Gamma Team',
            'Delta Team',
            'Delta Team'
        ]
    }
    
    df = pd.DataFrame(sample_data)
    demo_file = 'multi_column_demo.xlsx'
    df.to_excel(demo_file, index=False)
    
    print(f"   ✅ Created demo file: {demo_file}")
    print(f"   📊 Data structure:")
    print(f"      - Polygon names: {len(set(sample_data['Polygon_Name']))} unique polygons")
    print(f"      - Image columns: 3 (Image_URL_1, Image_URL_2, Image_URL_3)")
    print(f"      - Description columns: 3 (Primary_Description, Secondary_Description, Field_Notes)")
    print(f"      - Total rows: {len(sample_data['Polygon_Name'])}")
    
    # Step 2: Initialize editor and preview
    print(f"\n2. Loading KML file and previewing updates...")
    
    try:
        editor = KMLPolygonEditor('MyArea.kml')
    except:
        print("   ⚠️  MyArea.kml not found, trying MyMap.kml...")
        try:
            editor = KMLPolygonEditor('MyMap.kml')
        except:
            print("   ❌ No KML file found. Please ensure you have a KML file.")
            return
    
    # Preview updates
    editor.preview_excel_updates(
        excel_file_path=demo_file,
        polygon_column='Polygon_Name',
        image_columns=['Image_URL_1', 'Image_URL_2', 'Image_URL_3'],
        description_columns=['Primary_Description', 'Secondary_Description', 'Field_Notes'],
        sheet_name=0
    )
    
    # Step 3: Apply updates
    print(f"\n3. Applying updates to KML...")
    
    summary = editor.update_polygons_from_excel(
        excel_file_path=demo_file,
        polygon_column='Polygon_Name',
        image_columns=['Image_URL_1', 'Image_URL_2', 'Image_URL_3'],
        description_columns=['Primary_Description', 'Secondary_Description', 'Field_Notes'],
        sheet_name=0,
        merge_with_existing=True
    )
    
    # Step 4: Show results
    print(f"\n4. Results Summary:")
    print("   " + "-" * 30)
    
    if summary["success"]:
        print(f"   ✅ Processing completed successfully!")
        print(f"   📈 Statistics:")
        print(f"      - Polygons updated: {len(summary['updated_polygons'])}")
        print(f"      - Polygons not found: {len(summary['not_found_polygons'])}")
        print(f"      - Total images added: {summary['total_images_added']}")
        print(f"      - Total descriptions added: {summary['total_descriptions_added']}")
        
        if summary['updated_polygons']:
            print(f"\n   ✅ Successfully updated polygons:")
            for polygon in summary['updated_polygons']:
                info = editor.get_polygon_info(polygon)
                if info:
                    print(f"      • {polygon}")
                    print(f"        - Images: {len(info['images_in_description'])}")
                    print(f"        - Description length: {len(info['description'])} chars")
        
        if summary['not_found_polygons']:
            print(f"\n   ⚠️  Polygons not found in KML:")
            for polygon in summary['not_found_polygons']:
                print(f"      • {polygon}")
        
        # Save result
        output_file = 'MyArea_multi_column_demo.kml'
        editor.save_kml(output_file)
        print(f"\n   💾 Saved updated KML to: {output_file}")
        
    else:
        print(f"   ❌ Processing failed: {summary['message']}")
    
    # Step 5: Cleanup and summary
    print(f"\n5. Demo Complete!")
    print("   " + "-" * 30)
    print(f"   📁 Files created:")
    print(f"      - {demo_file} (demo Excel data)")
    if summary.get("success"):
        print(f"      - MyArea_multi_column_demo.kml (updated KML)")
    
    print(f"\n   🎯 Key Features Demonstrated:")
    print(f"      ✅ Multiple image columns (3 columns)")
    print(f"      ✅ Multiple description columns (3 columns)")
    print(f"      ✅ Fuzzy polygon name matching")
    print(f"      ✅ Handling of empty cells")
    print(f"      ✅ Error handling for missing polygons")
    print(f"      ✅ Data merging with existing content")
    
    print(f"\n   📚 For your own data:")
    print(f"      1. Run: python create_template.py")
    print(f"      2. Fill the template with your data")
    print(f"      3. Run: python quick_update.py")

if __name__ == "__main__":
    demonstrate_multi_column_system()
