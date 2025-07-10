#!/usr/bin/env python3
"""
Create a template Excel file for polygon data entry
This creates an empty template that you can fill with your own data
"""

import pandas as pd
from main import KMLPolygonEditor

def create_template_excel():
    """Create a template Excel file with multiple image and description columns"""
    
    # First, let's get the actual polygon names from your KML file
    try:
        editor = KMLPolygonEditor('MyArea.kml')
        polygons = editor.get_all_polygons()
        polygon_names = [p['name'] for p in polygons]
    except:
        # If KML file doesn't exist, use example polygon names
        polygon_names = ['Polygon 1', 'Polygon 2', 'Polygon 3']
    
    # Create template data with multiple image and description columns
    template_data = {
        'Polygon_Name': [],
        'Image_URL_1': [],
        'Image_URL_2': [],
        'Image_URL_3': [],
        'Description_1': [],
        'Description_2': [],
        'Notes': [],
        'Date': [],
        'Observer': []
    }
    
    # Add example rows for each polygon (you can fill these with your data)
    for polygon_name in polygon_names:
        # Add 3 example rows per polygon for demonstration
        for i in range(3):
            template_data['Polygon_Name'].append(polygon_name)
            template_data['Image_URL_1'].append('')  # Fill with your image URLs
            template_data['Image_URL_2'].append('')  # Fill with your image URLs
            template_data['Image_URL_3'].append('')  # Fill with your image URLs
            template_data['Description_1'].append('')  # Fill with your descriptions
            template_data['Description_2'].append('')  # Fill with your descriptions
            template_data['Notes'].append('')  # Fill with additional notes
            template_data['Date'].append('')  # Fill with date
            template_data['Observer'].append('')  # Fill with observer name
    
    # Create DataFrame and save to Excel
    df = pd.DataFrame(template_data)
    
    # Save template
    template_file = 'polygon_data_template.xlsx'
    df.to_excel(template_file, index=False)
    
    print(f"Created template Excel file: {template_file}")
    print("\nTemplate structure:")
    print("- Polygon_Name: Name of the polygon (must match KML)")
    print("- Image_URL_1, Image_URL_2, Image_URL_3: Multiple image URLs")
    print("- Description_1, Description_2: Multiple description fields")
    print("- Notes: Additional notes")
    print("- Date: Date of observation")
    print("- Observer: Name of observer")
    print("\nInstructions:")
    print("1. Open the Excel file")
    print("2. Fill in your data in the empty cells")
    print("3. You can delete rows you don't need")
    print("4. You can add more rows for additional data")
    print("5. Run the process_template.py script to apply the data")
    
    return template_file

def create_example_with_data():
    """Create an example Excel file with sample data"""
    
    # Example data with multiple columns
    example_data = {
        'Polygon_Name': [
            'القطاع A',
            'القطاع A',
            'القطاع A',
            'القطاع A',
            'Zone B',
            'Zone B'
        ],
        'Image_URL_1': [
            'https://example.com/area_a_photo1.jpg',
            'https://example.com/area_a_photo2.jpg',
            '',
            'https://example.com/area_a_photo4.jpg',
            'https://example.com/zone_b_photo1.jpg',
            ''
        ],
        'Image_URL_2': [
            'https://example.com/area_a_detail1.jpg',
            '',
            'https://example.com/area_a_detail3.jpg',
            '',
            'https://example.com/zone_b_detail1.jpg',
            'https://example.com/zone_b_detail2.jpg'
        ],
        'Image_URL_3': [
            '',
            'https://example.com/area_a_overview2.jpg',
            '',
            '',
            '',
            'https://example.com/zone_b_overview.jpg'
        ],
        'Description_1': [
            'Infrastructure assessment completed',
            'Population density survey',
            'Environmental impact evaluation',
            'Follow-up inspection conducted',
            'Initial zone survey',
            'Zone boundary verification'
        ],
        'Description_2': [
            'All systems operational',
            'High density residential area',
            'No environmental concerns',
            'Minor repairs needed',
            'Commercial district identified',
            'Boundary markers placed'
        ],
        'Notes': [
            'Weather: Clear, Good visibility',
            'Weather: Overcast, Limited visibility',
            'Weather: Clear, Excellent conditions',
            'Weather: Rainy, Poor visibility',
            'Weather: Clear, Good conditions',
            'Weather: Partly cloudy'
        ],
        'Date': [
            '2025-07-01',
            '2025-07-02',
            '2025-07-03',
            '2025-07-10',
            '2025-07-05',
            '2025-07-06'
        ],
        'Observer': [
            'Team Alpha',
            'Team Alpha',
            'Team Beta',
            'Team Alpha',
            'Team Gamma',
            'Team Gamma'
        ]
    }
    
    df = pd.DataFrame(example_data)
    example_file = 'polygon_data_example.xlsx'
    df.to_excel(example_file, index=False)
    
    print(f"\nCreated example Excel file: {example_file}")
    print("This file contains sample data showing how to structure your information.")
    
    return example_file

if __name__ == "__main__":
    print("Creating Excel templates for polygon data...")
    print("=" * 50)
    
    # Create empty template
    template_file = create_template_excel()
    
    # Create example with data
    example_file = create_example_with_data()
    
    print("\n" + "=" * 50)
    print("Files created:")
    print(f"1. {template_file} - Empty template for your data")
    print(f"2. {example_file} - Example with sample data")
    print("\nNext steps:")
    print("1. Edit the template file with your data")
    print("2. Run process_template.py to apply the data to your KML")
    print("=" * 50)
