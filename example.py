#!/usr/bin/env python3
"""
Example script showing how to use the KML Polygon Editor without Excel files
"""

from main import KMLPolygonEditor

def main():
    print("KML Polygon Editor Example")
    print("=" * 40)

    # Initialize the editor
    editor = KMLPolygonEditor('MyMap.kml')
    
    # List all polygons
    print("\n1. Listing all polygons:")
    polygons = editor.list_polygons()
    
    if polygons:
        # Get info for the first polygon
        first_polygon = polygons[0]
        print(f"\n2. Getting details for '{first_polygon['name']}':")
        info = editor.get_polygon_info(first_polygon['name'])
        
        if info:
            print(f"   Name: {info['name']}")
            print(f"   Current description: {info['description'][:100]}...")
            print(f"   Images: {len(info['images_in_description'])} found")
            
            # Example: Update the polygon
            print(f"\n3. Updating polygon '{first_polygon['name']}':")
            new_description = "Updated description from example script"
            new_images = [
                "https://kc.kobotoolbox.org/media/original?media_file=molhamteam%2Fattachments%2Fd29d77f4f6604d6b9c21f60a9ed2b4f0%2Ffe39de73-2b9a-4aba-9f3a-fae8590ea9a5%2F1749023464716.jpg",
                "https://kc.kobotoolbox.org/media/original?media_file=molhamteam%2Fattachments%2Fd29d77f4f6604d6b9c21f60a9ed2b4f0%2F446b089c-d08f-4697-8ee5-d380959fd106%2F1749024125108.jpg",
                "https://kc.kobotoolbox.org/media/original?media_file=molhamteam%2Fattachments%2Fd29d77f4f6604d6b9c21f60a9ed2b4f0%2F88ce007c-ab12-4dcc-a91e-4844b307a13c%2F1749024667220.jpg"
            ]
            
            success = editor.update_polygon(
                first_polygon['name'],
                new_description, 
                new_images
            )
            
            if success:
                print("   ✓ Polygon updated successfully!")
                
                # Save to a new file
                print("\n4. Saving updated KML:")
                editor.save_kml('MyMap_example_updated.kml')
                print("   ✓ Saved to 'MyMap_example_updated.kml'")
            else:
                print("   ✗ Failed to update polygon")
    else:
        print("No polygons found in the KML file.")

if __name__ == "__main__":
    main()
