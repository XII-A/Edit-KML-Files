# Excel Template Usage Guide

## 📋 How to Use the Excel Template System

### Step 1: Create Your Template

Run this command to create Excel templates:

```bash
python create_template.py
```

This creates:

- `polygon_data_template.xlsx` - Empty template for your data
- `polygon_data_example.xlsx` - Example with sample data

### Step 2: Fill Your Template

Open `polygon_data_template.xlsx` in Excel and fill in your data:

| Column            | Description                               | Example                               |
| ----------------- | ----------------------------------------- | ------------------------------------- |
| **Polygon_Name**  | Must match your KML polygon names exactly | "القطاع A"                            |
| **Image_URL_1**   | First image URL                           | "https://example.com/photo1.jpg"      |
| **Image_URL_2**   | Second image URL                          | "https://example.com/photo2.jpg"      |
| **Image_URL_3**   | Third image URL                           | "https://example.com/photo3.jpg"      |
| **Description_1** | Main description                          | "Infrastructure assessment completed" |
| **Description_2** | Additional description                    | "All systems operational"             |
| **Notes**         | Additional notes                          | "Weather: Clear, Good visibility"     |
| **Date**          | Date of observation                       | "2025-07-10"                          |
| **Observer**      | Observer name                             | "Team Alpha"                          |

### Step 3: Process Your Data

#### Option A: Quick Update (Recommended)

```bash
python quick_update.py
```

#### Option B: Interactive Processing

```bash
python process_template.py
```

#### Option C: Manual Processing

```bash
python main.py
```

Then choose option 5 (Update from Excel file)

### 📊 Excel Template Features

#### Multiple Images per Polygon

- Each row can have up to 3 image URLs
- Leave blank if you have fewer images
- All images from all rows for the same polygon will be combined

#### Multiple Descriptions per Polygon

- Description_1: Main description
- Description_2: Additional details
- Notes: Extra notes or observations
- All descriptions from all rows for the same polygon will be combined

#### Flexible Data Entry

- You can have multiple rows for the same polygon
- Each row can have different combinations of images and descriptions
- Empty cells are automatically ignored

### 💡 Tips for Success

1. **Polygon Names Must Match**

   - Copy polygon names exactly from your KML file
   - The system handles Unicode characters automatically
   - Case doesn't matter, but spelling must be exact

2. **Image URLs**

   - Use complete URLs (starting with http:// or https://)
   - Test URLs in a browser to make sure they work
   - Supported formats: JPG, PNG, GIF, WebP

3. **Descriptions**

   - Keep descriptions concise but informative
   - Use clear, descriptive language
   - Multiple description fields will be combined

4. **Data Organization**
   - One row per observation/image set
   - Multiple rows per polygon are allowed
   - Delete unused rows to keep the file clean

### 🔧 Advanced Usage

#### Custom Column Names

If you want to use different column names, you can modify the scripts:

```python
# In your custom script
editor.update_polygons_from_excel(
    excel_file_path='your_file.xlsx',
    polygon_column='Your_Polygon_Column',
    image_columns=['Your_Image_Col1', 'Your_Image_Col2'],
    description_columns=['Your_Desc_Col1', 'Your_Desc_Col2'],
    merge_with_existing=True
)
```

#### Multiple Sheets

You can specify which sheet to use:

```python
editor.update_polygons_from_excel(
    excel_file_path='your_file.xlsx',
    polygon_column='Polygon_Name',
    image_columns=['Image_URL_1', 'Image_URL_2'],
    description_columns=['Description_1', 'Description_2'],
    sheet_name='Sheet2'  # or sheet_name=1 for second sheet
)
```

### 🚨 Troubleshooting

**Problem**: "Polygon not found in KML file"

- **Solution**: Check that polygon names match exactly
- Use the interactive editor (option 1) to see exact polygon names

**Problem**: "Excel file not found"

- **Solution**: Make sure you're in the correct directory
- Run `python create_template.py` first

**Problem**: "Column not found"

- **Solution**: Check that your Excel file has the correct column names
- Don't rename the columns in the template

**Problem**: Images not showing

- **Solution**: Verify image URLs are accessible
- Check that URLs start with http:// or https://

### 📁 File Structure

After running the scripts, you'll have:

```
📁 Your Project Folder
├── 📄 MyArea.kml (your original KML)
├── 📄 MyArea_updated.kml (updated KML)
├── 📊 polygon_data_template.xlsx (your template)
├── 📊 polygon_data_example.xlsx (example data)
├── 🐍 main.py (main editor)
├── 🐍 create_template.py (create templates)
├── 🐍 process_template.py (process data)
└── 🐍 quick_update.py (quick processing)
```

### 🎯 Example Workflow

1. **Create template**: `python create_template.py`
2. **Fill template**: Open Excel, add your data
3. **Process data**: `python quick_update.py`
4. **Check results**: Open the `*_updated.kml` file

Your KML polygons will now have all the images and descriptions from your Excel file! 🎉
