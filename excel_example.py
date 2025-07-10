#!/usr/bin/env python3
"""
Example script showing how to use the KML Polygon Editor with Excel files
"""

from main import KMLPolygonEditor
import pandas as pd



def example_excel_integration():
    """Example of how to integrate Excel data with KML polygons"""
    print("KML Polygon Editor - Excel Integration Example")
    print("=" * 50)
    
    # Initialize the editor
    editor = KMLPolygonEditor('MyMap.kml')
    
    # Create example Excel file
    
    
    
    # Preview updates
    editor.preview_excel_updates(
        excel_file_path=excel_file,
        polygon_column='اسم/ رقم القطاع:',
        image_column='Image_URL', 
        description_column='Description_Text',
        sheet_name=2
    )
    
    print(f"\n2. Applying updates from {excel_file}:")
    print("-" * 50)
    
    # Apply updates
    summary = editor.update_polygons_from_excel(
        excel_file_path=excel_file,
        polygon_column='Polygon_Name',
        image_column='Image_URL',
        description_column='Description_Text',
        sheet_name=0,
        merge_with_existing=True
    )
    
    if summary["success"]:
        print(f"\n✅ Update Summary:")
        print(f"   Updated polygons: {len(summary['updated_polygons'])}")
        print(f"   Not found polygons: {len(summary['not_found_polygons'])}")
        print(f"   Total images added: {summary['total_images_added']}")
        print(f"   Total descriptions added: {summary['total_descriptions_added']}")
        
        if summary['updated_polygons']:
            print(f"\n   Successfully updated: {', '.join(summary['updated_polygons'])}")
        
        if summary['not_found_polygons']:
            print(f"\n   Polygons not found in KML: {', '.join(summary['not_found_polygons'])}")
        
        # Save updated KML
        output_file = 'MyMap_excel_updated.kml'
        editor.save_kml(output_file)
        print(f"\n3. Saved updated KML to: {output_file}")
        
    else:
        print(f"❌ Update failed: {summary['message']}")

def example_custom_excel_format():
    """Example using a custom Excel format"""
    print("\n" + "=" * 50)
    print("Custom Excel Format Example")
    print("=" * 50)
    
    # Create custom Excel with different column names
    custom_data = {
        'Area_Name': ['القطاع A', 'القطاع A'],
        'Photo_Link': [
            'https://example.com/custom1.jpg',
            'https://example.com/custom2.jpg'
        ],
        'Notes': [
            'Custom note 1 for the area',
            'Custom note 2 with additional details'
        ],
        'Category': ['Residential', 'Commercial']
    }
    
    df = pd.DataFrame(custom_data)
    custom_excel = 'custom_polygon_data.xlsx'
    df.to_excel(custom_excel, index=False)
    print(f"Created custom Excel file: {custom_excel}")
    
    # Use with editor
    editor = KMLPolygonEditor('MyMap.kml')
    
    print(f"\nPreview from {custom_excel}:")
    editor.preview_excel_updates(
        excel_file_path=custom_excel,
        polygon_column='Area_Name',  # Different column name
        image_column='Photo_Link',   # Different column name
        description_column='Notes',  # Different column name
        sheet_name=0
    )

if __name__ == "__main__":
    # Run the example
    example_excel_integration()
    
    # Run custom format example
    example_custom_excel_format()
    
    print("\n" + "=" * 50)
    print("Examples completed!")
    print("Check the generated files:")
    print("- polygon_data.xlsx (example data)")
    print("- custom_polygon_data.xlsx (custom format)")
    print("- MyMap_excel_updated.kml (updated KML)")
    print("=" * 50)
