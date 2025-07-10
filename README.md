# KML Polygon Editor

A Python tool for editing polygon descriptions and images in KML files. This tool allows you to modify polygon metadata such as descriptions and images based on polygon names.

## Features

- **List Polygons**: View all polygons in your KML file with their names
- **View Details**: See current description and images for any polygon
- **Edit Polygons**: Update descriptions and images for specific polygons
- **Excel Integration**: Bulk update polygons from Excel/CSV data
- **Fuzzy Matching**: Handles Unicode variations and whitespace in polygon names
- **Interactive Mode**: User-friendly command-line interface
- **Programmatic Mode**: Use as a Python library for batch operations
- **Safe Editing**: Creates backups and preserves original KML structure

## Installation

1. **Clone or download this repository**
2. **Install Python dependencies**:

   ```bash
   pip install -r requirements.txt
   ```

   Or install manually:

   ```bash
   pip install pykml lxml pandas openpyxl
   ```

## Usage

### Interactive Mode

Run the main script to use the interactive editor:

```bash
python main.py
```

This will open a menu with the following options:

1. **List all polygons** - Shows all polygon names in your KML file
2. **View polygon details** - Display current description and images for a specific polygon
3. **Edit polygon manually** - Modify description and images for a polygon
4. **Preview Excel updates** - See what changes would be made from an Excel file
5. **Update from Excel file** - Apply bulk updates from Excel data
6. **Save KML file** - Save your changes to a file
7. **Exit** - Close the editor

### Programmatic Mode

You can also use the KMLPolygonEditor class in your own scripts:

```python
from main import KMLPolygonEditor

# Initialize the editor
editor = KMLPolygonEditor('MyMap.kml')

# List all polygons
polygons = editor.list_polygons()

# Get details for a specific polygon
info = editor.get_polygon_info("القطاع A")

# Update a polygon
new_description = "This is the updated description"
new_images = [
    "https://example.com/image1.jpg",
    "https://example.com/image2.jpg"
]

editor.update_polygon("القطاع A", new_description, new_images)

# Save the changes
editor.save_kml('MyMap_updated.kml')
```

### Excel Integration

You can bulk update polygons from Excel files:

```python
from main import KMLPolygonEditor

# Initialize the editor
editor = KMLPolygonEditor('MyMap.kml')

# Update polygons from Excel file
summary = editor.update_polygons_from_excel(
    excel_file_path='polygon_data.xlsx',
    polygon_column='Polygon_Name',      # Column with polygon names
    image_column='Image_URL',           # Column with image URLs
    description_column='Description',   # Column with description text
    sheet_name=0,                      # Sheet index or name
    merge_with_existing=True           # Keep existing data
)

# Save the changes
editor.save_kml('updated_map.kml')
```

#### Excel File Format

Your Excel file should have columns for:

- **Polygon names** (must match KML polygon names)
- **Image URLs** (one image per row)
- **Description text** (text to add to polygon description)

Example Excel structure:
| Polygon_Name | Image_URL | Description_Text |
|-------------|-----------|------------------|
| القطاع A | https://example.com/img1.jpg | First observation |
| القطاع A | https://example.com/img2.jpg | Second observation |
| Sector B | https://example.com/img3.jpg | Infrastructure review |

#### Fuzzy Name Matching

The editor automatically handles:

- Unicode formatting characters (like RTL marks)
- Extra whitespace
- Case variations
- Different Unicode normalizations

This means "القطاع A" in Excel will match "القطاع A ‎" in your KML file.

### Example Script

Run the example script to see how the editor works:

```bash
python example.py
```

This script will:

1. Load your KML file
2. List all polygons
3. Show details for the first polygon
4. Update it with new description and images
5. Save the result to a new file

## KML Structure

The editor works with KML files that contain:

- **Polygons**: Geographic shapes defined by coordinates
- **Names**: Each polygon should have a name element
- **Descriptions**: HTML content that can include images and text
- **Extended Data**: Additional metadata including media links

### Example KML Structure

```xml
<Placemark>
  <name>القطاع A</name>
  <description><![CDATA[
    <img src="https://example.com/image.jpg" height="200" width="auto" /><br>
    <br>Description text here
  ]]></description>
  <ExtendedData>
    <Data name="gx_media_links">
      <value><![CDATA[https://example.com/image.jpg]]></value>
    </Data>
  </ExtendedData>
  <Polygon>
    <!-- Polygon coordinates -->
  </Polygon>
</Placemark>
```

## Methods

### KMLPolygonEditor Class

#### `__init__(kml_file_path)`

Initialize the editor with a KML file path.

#### `list_polygons()`

Returns a list of all polygons with their names.

#### `get_polygon_info(polygon_name)`

Get detailed information about a specific polygon including:

- Name
- Current description
- Images in description
- Media links

#### `update_polygon(polygon_name, new_description=None, new_images=None)`

Update a polygon's description and/or images:

- `polygon_name`: Name of the polygon to update
- `new_description`: New text description (optional)
- `new_images`: List of image URLs (optional)

#### `load_excel_data(excel_file_path, polygon_column, image_column, description_column, sheet_name=0)`

Load and parse data from an Excel file:

- `excel_file_path`: Path to the Excel file
- `polygon_column`: Column name containing polygon names
- `image_column`: Column name containing image URLs
- `description_column`: Column name containing description text
- `sheet_name`: Sheet name or index (default: 0)

#### `update_polygons_from_excel(excel_file_path, polygon_column, image_column, description_column, sheet_name=0, merge_with_existing=True)`

Bulk update polygons from Excel data:

- Returns a summary dictionary with update statistics
- `merge_with_existing`: Whether to keep existing data or replace it

#### `preview_excel_updates(excel_file_path, polygon_column, image_column, description_column, sheet_name=0)`

Preview what changes would be made without actually updating.

#### `normalize_polygon_name(name)`

Normalize polygon names to handle Unicode variations and whitespace.

#### `find_polygon_by_name_fuzzy(target_name)`

Find polygons using fuzzy matching for Unicode and case variations.

#### `save_kml(output_file=None)`

Save the modified KML file:

- `output_file`: Path to save file (optional, defaults to original file)

## Image Handling

The editor supports:

- **Multiple images per polygon**
- **Image URLs** (http/https links)
- **Automatic HTML formatting** for images in descriptions
- **Media links in Extended Data** (first image becomes the main media link)

Images are automatically formatted as:

```html
<img src="URL" height="200" width="auto" /><br />
```

## Error Handling

The editor includes error handling for:

- Invalid KML files
- Missing polygons
- File reading/writing errors
- Malformed image URLs

## File Output

When saving KML files, the editor:

- Preserves original KML structure and styling
- Maintains proper XML formatting
- Uses UTF-8 encoding
- Includes XML declaration

## Requirements

- Python 3.6+
- lxml >= 4.9.0
- pandas >= 1.3.0 (for Excel integration)
- openpyxl >= 3.0.0 (for Excel file support)

## Limitations

- Works with standard KML polygon structures
- Requires polygon names to be unique for identification
- Images must be accessible URLs (not local files)
- Preserves existing KML styles and formatting

## Examples

### Adding Multiple Images

```python
editor = KMLPolygonEditor('MyMap.kml')

images = [
    "https://example.com/photo1.jpg",
    "https://example.com/photo2.jpg",
    "https://example.com/photo3.jpg"
]

editor.update_polygon("My Polygon", "Updated description", images)
editor.save_kml('updated_map.kml')
```

### Batch Update All Polygons

```python
editor = KMLPolygonEditor('MyMap.kml')
polygons = editor.get_all_polygons()

for polygon in polygons:
    name = polygon['name']
    new_desc = f"Updated description for {name}"
    new_images = [f"https://example.com/images/{name.replace(' ', '_')}.jpg"]

    editor.update_polygon(name, new_desc, new_images)

editor.save_kml('batch_updated.kml')
```

## Troubleshooting

### Common Issues

1. **"Polygon not found"**: Check that the polygon name exactly matches (case-sensitive)
2. **File encoding errors**: Ensure your KML file is UTF-8 encoded
3. **Invalid XML**: Verify your KML file is properly formatted XML

### Debug Mode

To see detailed information during operation, check the console output for:

- File loading success/failure messages
- Polygon update confirmations
- Save operation results

## License

This tool is provided as-is for educational and development purposes.
