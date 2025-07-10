#!/usr/bin/env python3
"""
Simple example showing Excel integration with normalized polygon names
"""

from main import KMLPolygonEditor
import pandas as pd

def create_simple_excel():
    """Create a simple Excel file that works with fuzzy matching"""
    data = {
        'Polygon_Name': [
            'القطاع A',  # Without Unicode formatting characters
            'القطاع A', 
            'القطاع A'
        ],
        'Image_URL': [
            'https://kc.kobotoolbox.org/media/original?media_file=molhamteam%2Fattachments%2Fd29d77f4f6604d6b9c21f60a9ed2b4f0%2Ffe39de73-2b9a-4aba-9f3a-fae8590ea9a5%2F1749023464716.jpg',
            'https://kc.kobotoolbox.org/media/original?media_file=molhamteam%2Fattachments%2Fd29d77f4f6604d6b9c21f60a9ed2b4f0%2F446b089c-d08f-4697-8ee5-d380959fd106%2F1749024125108.jpg',
            'https://kc.kobotoolbox.org/media/original?media_file=molhamteam%2Fattachments%2Fd29d77f4f6604d6b9c21f60a9ed2b4f0%2F88ce007c-ab12-4dcc-a91e-4844b307a13c%2F1749024667220.jpg'
        ],
        'Description_Text': [
            'Excel Integration Test: First image and description added.',
            'Excel Integration Test: Second image showing different perspective.',
            'Excel Integration Test: Third image completing the documentation.'
        ],
        'Date_Added': [
            '2025-07-10',
            '2025-07-10', 
            '2025-07-10'
        ]
    }
    
    df = pd.DataFrame(data)
    df.to_excel('simple_polygon_data.xlsx', index=False)
    print("Created simple Excel file: simple_polygon_data.xlsx")
    return 'simple_polygon_data.xlsx'

def main():
    print("Simple Excel Integration Example")
    print("=" * 40)
    
    # Initialize the editor
    editor = KMLPolygonEditor('MyMap.kml')
    
    # Create simple Excel file
    excel_file = create_simple_excel()
    
    print(f"\n1. Current polygon info:")
    polygons = editor.list_polygons()
    
    if polygons:
        current_info = editor.get_polygon_info(polygons[0]['name'])
        if current_info:
            print(f"   Current images: {len(current_info['images_in_description'])}")
            print(f"   Current description length: {len(current_info['description'])}")
    
    print(f"\n2. Preview updates from {excel_file}:")
    editor.preview_excel_updates(
        excel_file_path=excel_file,
        polygon_column='Polygon_Name',
        image_column='Image_URL', 
        description_column='Description_Text'
    )
    
    print(f"\n3. Applying updates:")
    summary = editor.update_polygons_from_excel(
        excel_file_path=excel_file,
        polygon_column='Polygon_Name',
        image_column='Image_URL',
        description_column='Description_Text',
        merge_with_existing=True
    )
    
    if summary["success"]:
        print(f"\n✅ Success! Updated {len(summary['updated_polygons'])} polygons")
        print(f"   Images added: {summary['total_images_added']}")
        print(f"   Descriptions added: {summary['total_descriptions_added']}")
        
        # Save result
        editor.save_kml('MyMap_simple_updated.kml')
        print(f"\n4. Saved to: MyMap_simple_updated.kml")
        
        # Show updated info
        if summary['updated_polygons']:
            updated_name = summary['updated_polygons'][0]
            updated_info = editor.get_polygon_info(updated_name)
            if updated_info:
                print(f"\n5. Updated polygon now has:")
                print(f"   Images: {len(updated_info['images_in_description'])}")
                print(f"   Description length: {len(updated_info['description'])}")
    else:
        print(f"❌ Failed: {summary['message']}")

if __name__ == "__main__":
    main()
