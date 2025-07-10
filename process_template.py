#!/usr/bin/env python3
"""
Process template Excel file and apply data to KML polygons
This script reads your filled template and updates the KML file
"""

from main import KMLPolygonEditor
import os

def process_template_excel(excel_file='polygon_data_template.xlsx', kml_file='MyArea.kml'):
    """
    Process the template Excel file and apply data to KML
    
    Args:
        excel_file (str): Path to the filled Excel template
        kml_file (str): Path to the KML file to update
    """
    
    print(f"Processing Excel template: {excel_file}")
    print(f"Target KML file: {kml_file}")
    print("=" * 50)
    
    # Check if files exist
    if not os.path.exists(excel_file):
        print(f"❌ Excel file not found: {excel_file}")
        print("Please run create_template.py first to create the template.")
        return
    
    if not os.path.exists(kml_file):
        print(f"❌ KML file not found: {kml_file}")
        print("Please make sure your KML file exists.")
        return
    
    # Initialize the editor
    editor = KMLPolygonEditor(kml_file)
    
    # Define the column structure
    polygon_column = 'Polygon_Name'
    image_columns = ['Image_URL_1', 'Image_URL_2', 'Image_URL_3']
    description_columns = ['Description_1', 'Description_2', 'Notes']
    
    print("\n1. Previewing updates...")
    print("-" * 30)
    
    # Preview the updates
    editor.preview_excel_updates(
        excel_file_path=excel_file,
        polygon_column=polygon_column,
        image_columns=image_columns,
        description_columns=description_columns,
        sheet_name=0
    )
    
    # Ask for confirmation
    confirm = input("\nDo you want to apply these updates? (y/n): ").strip().lower()
    if confirm != 'y':
        print("Operation cancelled.")
        return
    
    print("\n2. Applying updates...")
    print("-" * 30)
    
    # Apply the updates
    summary = editor.update_polygons_from_excel(
        excel_file_path=excel_file,
        polygon_column=polygon_column,
        image_columns=image_columns,
        description_columns=description_columns,
        sheet_name=0,
        merge_with_existing=True
    )
    
    if summary["success"]:
        print(f"\n✅ Updates applied successfully!")
        print(f"   Updated polygons: {len(summary['updated_polygons'])}")
        print(f"   Not found polygons: {len(summary['not_found_polygons'])}")
        print(f"   Total images added: {summary['total_images_added']}")
        print(f"   Total descriptions added: {summary['total_descriptions_added']}")
        
        if summary['updated_polygons']:
            print(f"\n   Successfully updated:")
            for polygon in summary['updated_polygons']:
                print(f"   • {polygon}")
        
        if summary['not_found_polygons']:
            print(f"\n   ⚠️  Polygons not found in KML:")
            for polygon in summary['not_found_polygons']:
                print(f"   • {polygon}")
        
        # Save the updated KML
        output_file = f"{kml_file.replace('.kml', '')}_updated.kml"
        editor.save_kml(output_file)
        print(f"\n3. Saved updated KML to: {output_file}")
        
        # Show summary of final state
        print(f"\n4. Final summary:")
        print("-" * 30)
        for polygon_name in summary['updated_polygons']:
            info = editor.get_polygon_info(polygon_name)
            if info:
                print(f"   {polygon_name}:")
                print(f"     - Images: {len(info['images_in_description'])}")
                print(f"     - Description length: {len(info['description'])} characters")
        
    else:
        print(f"❌ Update failed: {summary['message']}")

def process_example_file():
    """Process the example Excel file"""
    process_template_excel('polygon_data_example.xlsx', 'MyArea.kml')

def main():
    print("KML Polygon Data Processor")
    print("=" * 50)
    print("This script will process your Excel template and update your KML file.")
    print("\nOptions:")
    print("1. Process template file (polygon_data_template.xlsx)")
    print("2. Process example file (polygon_data_example.xlsx)")
    print("3. Process custom file")
    print("4. Exit")
    
    choice = input("\nEnter your choice (1-4): ").strip()
    
    if choice == '1':
        process_template_excel()
    elif choice == '2':
        process_example_file()
    elif choice == '3':
        excel_file = input("Enter Excel file path: ").strip()
        kml_file = input("Enter KML file path: ").strip()
        process_template_excel(excel_file, kml_file)
    elif choice == '4':
        print("Goodbye!")
    else:
        print("Invalid choice.")

if __name__ == "__main__":
    main()
