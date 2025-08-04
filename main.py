from lxml import etree
import re
import pandas as pd
import os

class KMLPolygonEditor:
    def __init__(self, kml_file_path):
        """Initialize the KML editor with the path to the KML file"""
        self.kml_file_path = kml_file_path
        self.tree = None
        self.root = None
        self.load_kml()
    
    def load_kml(self):
        """Load and parse the KML file"""
        try:
            self.tree = etree.parse(self.kml_file_path)
            self.root = self.tree.getroot()
            print(f"Successfully loaded KML file: {self.kml_file_path}")
        except Exception as e:
            print(f"Error loading KML file: {e}")
    
    def get_all_polygons(self):
        """Get all polygons with their names from the KML file"""
        polygons = []
        
        # Find all Placemark elements that contain Polygon elements
        for placemark in self.root.xpath('.//kml:Placemark[kml:Polygon]', 
                                       namespaces={'kml': 'http://www.opengis.net/kml/2.2'}):
            name_element = placemark.find('.//{http://www.opengis.net/kml/2.2}name')
            name = name_element.text if name_element is not None else "Unnamed Polygon"
            
            # Clean up the name (remove extra whitespace and newlines)
            name = re.sub(r'\s+', ' ', name.strip())
            
            polygons.append({
                'name': name,
                'placemark': placemark
            })
        
        return polygons
    
    def list_polygons(self):
        """List all polygons in the KML file"""
        polygons = self.get_all_polygons()
        
        if not polygons:
            print("No polygons found in the KML file.")
            return
        
        print("\nPolygons found in the KML file:")
        print("-" * 50)
        for i, polygon in enumerate(polygons, 1):
            print(f"{i}. {polygon['name']}")
        print("-" * 50)
        
        return polygons
    
    def normalize_polygon_name(self, name):
        """
        Normalize polygon names to handle Unicode and whitespace variations
        
        Args:
            name (str): The polygon name to normalize
            
        Returns:
            str: Normalized polygon name
        """
        if not name:
            return ""
        
        # Remove common Unicode formatting characters
        import unicodedata
        name = unicodedata.normalize('NFKC', name)
        
        # Remove directional marks and other formatting characters
        formatting_chars = ['\u200e', '\u200f', '\u202a', '\u202b', '\u202c', '\u202d', '\u202e']
        for char in formatting_chars:
            name = name.replace(char, '')
        
        # Normalize whitespace
        name = re.sub(r'\s+', ' ', name.strip())
        
        return name
    
    def find_polygon_by_name_fuzzy(self, target_name):
        """
        Find polygon by name with fuzzy matching to handle Unicode variations
        
        Args:
            target_name (str): The polygon name to search for
            
        Returns:
            dict or None: Polygon info if found, None otherwise
        """
        normalized_target = self.normalize_polygon_name(target_name)
        
        for polygon in self.get_all_polygons():
            original_name = polygon['name']
            normalized_name = self.normalize_polygon_name(original_name)
            
            # Exact match after normalization
            if normalized_name == normalized_target:
                return polygon
            
            # Case-insensitive match
            if normalized_name.lower() == normalized_target.lower():
                return polygon
        
        return None

    def get_polygon_info(self, polygon_name):
        """Get current description and images for a specific polygon"""
        polygons = self.get_all_polygons()
        
        for polygon in polygons:
            if polygon['name'].strip().lower() == polygon_name.strip().lower():
                placemark = polygon['placemark']
                
                # Get description
                desc_element = placemark.find('.//{http://www.opengis.net/kml/2.2}description')
                current_description = desc_element.text if desc_element is not None else ""
                
                # Extract images from description
                images = []
                if current_description:
                    # Find img tags
                    img_pattern = r'<img[^>]+src=["\']([^"\']+)["\'][^>]*>'
                    images = re.findall(img_pattern, current_description, re.IGNORECASE)
                
                # Get media links from ExtendedData
                media_links = []
                media_element = placemark.find('.//kml:Data[@name="gx_media_links"]/kml:value', 
                                            namespaces={'kml': 'http://www.opengis.net/kml/2.2'})
                if media_element is not None and media_element.text:
                    media_links = [media_element.text.strip()]
                
                return {
                    'name': polygon['name'],
                    'description': current_description,
                    'images_in_description': images,
                    'media_links': media_links,
                    'placemark': placemark
                }
        
        return None
    
    def calculate_centroid(self, coordinates_str):
        """Calculate the centroid of a polygon from its coordinates string"""
        # Parse coordinates string into list of points
        coords = []
        for point in coordinates_str.strip().split():
            lon, lat, _ = map(float, point.split(','))
            coords.append((lon, lat))
        
        # Calculate centroid
        if not coords:
            return None
        
        x_total = sum(x for x, _ in coords)
        y_total = sum(y for _, y in coords)
        count = len(coords)
        
        return f"{x_total/count},{y_total/count},0"

    def update_polygon(self, polygon_name, new_description=None, new_images=None):
        """Update a polygon's description and/or images"""
        polygon_info = self.get_polygon_info(polygon_name)
        
        if not polygon_info:
            print(f"Polygon '{polygon_name}' not found.")
            return False
        
        placemark = polygon_info['placemark']
        
        # Create or find Document element to store shared style
        ns = {'kml': 'http://www.opengis.net/kml/2.2'}
        document = self.root.find('kml:Document', namespaces=ns)
        if document is None:
            document = etree.Element('{http://www.opengis.net/kml/2.2}Document')
            self.root.insert(0, document)
        
        # Create shared style for invisible points if it doesn't exist
        style_id = "noIconStyle"
        shared_style = document.find(f'.//kml:Style[@id="{style_id}"]', namespaces=ns)
        if shared_style is None:
            shared_style = etree.SubElement(document, '{http://www.opengis.net/kml/2.2}Style')
            shared_style.set('id', style_id)
            
            icon_style = etree.SubElement(shared_style, '{http://www.opengis.net/kml/2.2}IconStyle')
            scale = etree.SubElement(icon_style, '{http://www.opengis.net/kml/2.2}scale')
            scale.text = '0'  # Make icon invisible
        
        # Convert to MultiGeometry if needed and add label point
        polygon = placemark.find('.//{http://www.opengis.net/kml/2.2}Polygon')
        if polygon is not None:
            # Get polygon coordinates
            coords = polygon.find('.//kml:coordinates', namespaces=ns)
            if coords is not None and coords.text:
                centroid = self.calculate_centroid(coords.text)
                
                # Create MultiGeometry
                multi_geom = placemark.find('.//{http://www.opengis.net/kml/2.2}MultiGeometry')
                if multi_geom is None:
                    multi_geom = etree.Element('{http://www.opengis.net/kml/2.2}MultiGeometry')
                    # Move polygon under MultiGeometry
                    multi_geom.append(polygon)
                    
                    # Add Point for label
                    point = etree.SubElement(multi_geom, '{http://www.opengis.net/kml/2.2}Point')
                    point_coords = etree.SubElement(point, '{http://www.opengis.net/kml/2.2}coordinates')
                    point_coords.text = centroid
                    
                    # Replace original polygon with MultiGeometry
                    for child in list(placemark):
                        if child.tag.endswith('Polygon'):
                            placemark.remove(child)
                    placemark.append(multi_geom)
                    
                    # Add style for invisible point
                    style_url = etree.SubElement(placemark, '{http://www.opengis.net/kml/2.2}styleUrl')
                    style_url.text = '#noIconStyle'

        # Update description
        if new_description is not None:
            desc_element = placemark.find('.//{http://www.opengis.net/kml/2.2}description')
            
            if desc_element is None:
                # Create description element if it doesn't exist
                desc_element = etree.SubElement(placemark, '{http://www.opengis.net/kml/2.2}description')
            
            # Build new description with images
            updated_description = ""
            
            if new_images:
                # Add images to description
                for img_url in new_images:
                    updated_description += f'<img src="{img_url}" height="200" width="auto" /><br>'
                updated_description += "<br>"
            
            # Add text description
            updated_description += new_description
            
            # Set content with CDATA
            desc_element.clear()
            desc_element.text = updated_description
        
        # Update media links using gx:Carousel structure
        # NOTE: KEPT FOR LEGACY SUPPORT, NOT USED IN CURRENT IMPLEMENTATION
        # if new_images:
        #     # Remove any existing ExtendedData
        #     old_extended_data = placemark.find('.//{http://www.opengis.net/kml/2.2}ExtendedData')
        #     if old_extended_data is not None:
        #         placemark.remove(old_extended_data)
            
        #     # Remove any existing Carousel
        #     old_carousel = placemark.find('.//{http://www.google.com/kml/ext/2.2}Carousel')
        #     if old_carousel is not None:
        #         placemark.remove(old_carousel)
            
        #     # Create carousel structure
        #     carousel = etree.SubElement(placemark, '{http://www.google.com/kml/ext/2.2}Carousel')
            
        #     # Add each image to the carousel
        #     for img_url in new_images:
        #         img_element = etree.SubElement(carousel, '{http://www.google.com/kml/ext/2.2}Image')
        #         # Create a simpler ID from the URL
        #         img_id = 'hosted_image_' + ''.join(c for c in img_url if c.isalnum())[:20] + str(hash(img_url))[-5:] # Ensure ID is unique
        #         img_element.set('{http://www.opengis.net/kml/2.2}id', img_id)
                
        #         # Add image URL
        #         img_url_element = etree.SubElement(img_element, '{http://www.google.com/kml/ext/2.2}imageUrl')
        #         img_url_element.text = img_url
        
        print(f"Successfully updated polygon '{polygon_name}'")
        return True
    
    def load_excel_data(self, excel_file_path, polygon_column, image_columns, description_columns, sheet_name=0):
        """
        Load data from Excel file for polygon updates
        
        Args:
            excel_file_path (str): Path to the Excel file
            polygon_column (str): Column name that contains polygon names
            image_columns (str or list): Column name(s) that contain image URLs
            description_columns (str or list): Column name(s) that contain description text
            sheet_name (str/int): Sheet name or index (default: 0 for first sheet)
        
        Returns:
            dict: Dictionary with polygon names as keys and lists of data as values
        """
        try:
            # Read Excel file
            df = pd.read_excel(excel_file_path, sheet_name=sheet_name)
            
            # Convert single columns to lists
            if isinstance(image_columns, str):
                image_columns = [image_columns]
            if isinstance(description_columns, str):
                description_columns = [description_columns]
            
            # Check if required columns exist
            all_required_columns = [polygon_column] + image_columns + description_columns
            missing_columns = [col for col in all_required_columns if col not in df.columns]
            
            if missing_columns:
                print(f"Error: Missing columns in Excel file: {missing_columns}")
                print(f"Available columns: {list(df.columns)}")
                return None
            
            # Group data by polygon name
            polygon_data = {}
            
            for _, row in df.iterrows():
                polygon_name = str(row[polygon_column]).strip()

                # Search for the first English letter
                match = re.search(r'[A-Za-z]', polygon_name)

                first_english_char = match.group(0) if match else None

                polygon_number = str(row['رقم الكتلة السكنية:']).strip().replace("'", "\\")

                polygon_name = f"{first_english_char}-{polygon_number}"

                print(polygon_name)                
                # Skip empty polygon names
                if not polygon_name or polygon_name.lower() == 'nan':
                    continue
                
                # Initialize polygon data if not exists
                if polygon_name not in polygon_data:
                    polygon_data[polygon_name] = {
                        'images': [],
                        'descriptions': [0,0,0,0,0,0,0,0,0]
                        # index 0: number of buildings under the polygon
                        # index 1: sum of the floors under the polygon
                        # index 2: sum of areas under the polygon
                        # index 3: number of safe buildings under the polygon
                        # index 4: number of semi-damaged buildings under the polygon
                        # index 5: number of damaged/needs to be removed buildings under the polygon
                        # index 6: number of completely destroyed buildings under the polygon
                        # index 7: sum of apartments under the polygon
                        # index 8: sum of the total costs of repairs for buildings under the polygon
                    }
                
                # Process all image columns
                for img_col in image_columns:
                    if img_col in df.columns:
                        image_url = str(row[img_col]).strip() if pd.notna(row[img_col]) else ""
                        if image_url and image_url.lower() != 'nan':
                            polygon_data[polygon_name]['images'].append(image_url)
                
                
                polygon_data[polygon_name]['descriptions'][0] += 1
                polygon_data[polygon_name]['descriptions'][1] += int(row['عدد الطوابق في البناء (رقما):']) if pd.notna(row['عدد الطوابق في البناء (رقما):']) else 0
                polygon_data[polygon_name]['descriptions'][2] += int(row['المساحة الاجمالية  (متر)']) if pd.notna(row['المساحة الاجمالية  (متر)']) else 0
                polygon_data[polygon_name]['descriptions'][7] += int(row['ما هو عدد الشقق في البناء:']) if pd.notna(row['ما هو عدد الشقق في البناء:']) else 0
                polygon_data[polygon_name]['descriptions'][8] += int(row['ما هي التكلفة التقديرية لترميم الشقق/المبنى؟']) if pd.notna(row['ما هي التكلفة التقديرية لترميم الشقق/المبنى؟']) else 0

                if  "سليم" in row['ما هو نوع الضرر الذي أصاب البناء:']:
                    polygon_data[polygon_name]['descriptions'][3] += 1
                elif "جزئي" in row['ما هو نوع الضرر الذي أصاب البناء:']:
                    polygon_data[polygon_name]['descriptions'][4] += 1
                elif "مدمر" in row['ما هو نوع الضرر الذي أصاب البناء:']:
                    polygon_data[polygon_name]['descriptions'][5] += 1
                elif "مهدوم" in row['ما هو نوع الضرر الذي أصاب البناء:']:
                    polygon_data[polygon_name]['descriptions'][6] += 1

                # Process all description columns
                for desc_col in description_columns:
                    if desc_col in df.columns:
                        description_text = str(row[desc_col]).strip() if pd.notna(row[desc_col]) else ""
                        if description_text and description_text.lower() != 'nan':
                            polygon_data[polygon_name]['descriptions'].append(description_text)
            
            print(f"Successfully loaded data for {len(polygon_data)} polygons from Excel file")
            return polygon_data
            
        except Exception as e:
            print(f"Error loading Excel file: {e}")
            return None
    
    def update_polygons_from_excel(self, excel_file_path, polygon_column, image_columns, description_columns, 
                                 sheet_name=0, merge_with_existing=True, border_color=None):
        """
        Update polygons with data from Excel file
        
        Args:
            excel_file_path (str): Path to the Excel file
            polygon_column (str): Column name that contains polygon names
            image_columns (str or list): Column name(s) that contain image URLs  
            description_columns (str or list): Column name(s) that contain description text
            sheet_name (str/int): Sheet name or index (default: 0 for first sheet)
            merge_with_existing (bool): Whether to merge with existing data or replace it
            border_color (str): Border color for polygons (HTML hex color, e.g. '#FF0000')
        
        Returns:
            dict: Summary of updates performed
        """
        # Load Excel data
        excel_data = self.load_excel_data(excel_file_path, polygon_column, image_columns, 
                                        description_columns, sheet_name)
        
        if not excel_data:
            return {"success": False, "message": "Failed to load Excel data"}
        
        # Get existing polygons
        existing_polygons = {p['name']: p for p in self.get_all_polygons()}
        
        update_summary = {
            "success": True,
            "updated_polygons": [],
            "not_found_polygons": [],
            "total_images_added": 0,
            "total_descriptions_added": 0
        }
        

        
        # Process each polygon from Excel data
        for polygon_name, data in excel_data.items():
            # Try fuzzy matching first
            polygon_match = self.find_polygon_by_name_fuzzy(polygon_name)
            
            if not polygon_match:
                update_summary["not_found_polygons"].append(polygon_name)
                print(f"Warning: Polygon '{polygon_name}' not found in KML file")
                continue
            
            # Use the actual polygon name from KML
            actual_polygon_name = polygon_match['name']
            
            # Get current polygon info
            current_info = self.get_polygon_info(actual_polygon_name)
            if not current_info:
                continue
            
            # Prepare new images
            new_images = data['images']
            if merge_with_existing:
                # Add to existing images
                existing_images = current_info['images_in_description']
                new_images = existing_images + [img for img in new_images if img not in existing_images]
            
            # Prepare new description
            new_description = ""
            if merge_with_existing and current_info['description']:
                # Extract existing text description (without HTML tags)
                existing_text = re.sub(r'<[^>]+>', '', current_info['description'])
                existing_text = re.sub(r'\s+', ' ', existing_text).strip()
                new_description = existing_text
            
            # Add new description parts
            if data['descriptions']:
                description_parts = data['descriptions'].copy()
                denominator = int(data['descriptions'][0]) # Total number of buildings
                
                description_parts[0] = f"عدد المباني: {description_parts[0]}"
    
                avgFloors = int(int(description_parts[1]) / denominator)
                description_parts[1] = f"عدد الطوابق في المباني: {avgFloors}"
                
                avgArea = int(int(description_parts[2]) / denominator)
                description_parts[2] = f"المساحة الاجمالية للمباني: {avgArea} متر"
                
                description_parts[3] = f"عدد المباني السليمة: {description_parts[3]}"
                description_parts[4] = f"عدد المباني المتضررة جزئيا: {description_parts[4]}"
                description_parts[5] = f"عدد المباني المتضررة بشكل كامل: {description_parts[5]}"
                description_parts[6] = f"عدد المباني المهدومة: {description_parts[6]}" 
                
                avgApartments = int(int(description_parts[7]) / denominator if denominator > 0 else 0)
                description_parts[7] = f"عدد الشقق في المباني: {avgApartments}"
                description_parts[8] = f"مجموع التكلفة التقديرية لترميم الشقق/المباني: {description_parts[8]}"
            
                new_description += "<br/>".join(
                    [
                        "اسم الكتلة السكنية: " + actual_polygon_name,
                        description_parts[0],
                        description_parts[1],
                        description_parts[2],
                        description_parts[7],
                        description_parts[8],
                        description_parts[3],
                        description_parts[4],
                        description_parts[5],
                        description_parts[6]
                    ]
                )
            
            # Update the polygon
            success = self.update_polygon(actual_polygon_name, new_description, new_images)
            
            if success:
                update_summary["updated_polygons"].append(actual_polygon_name)
                update_summary["total_images_added"] += len(data['images'])
                update_summary["total_descriptions_added"] += len(data['descriptions'])
        
        # Convert all polygons to include label points
        self.convert_all_polygons_to_labeled()
        
        # If border_color is specified, update all polygon styles
        if border_color:
            print("Setting border color for all polygons...")
            self.set_all_polygon_border_colors(border_color)

        return update_summary
    
    def convert_all_polygons_to_labeled(self):
        """Convert all polygons to use MultiGeometry with label points"""
        ns = {'kml': 'http://www.opengis.net/kml/2.2'}
        
        # Create or find Document element to store shared style
        document = self.root.find('kml:Document', namespaces=ns)
        if document is None:
            document = etree.Element('{http://www.opengis.net/kml/2.2}Document')
            self.root.insert(0, document)
        
        # Create shared style for invisible points if it doesn't exist
        style_id = "noIconStyle"
        shared_style = document.find(f'.//kml:Style[@id="{style_id}"]', namespaces=ns)
        if shared_style is None:
            shared_style = etree.SubElement(document, '{http://www.opengis.net/kml/2.2}Style')
            shared_style.set('id', style_id)
            
            icon_style = etree.SubElement(shared_style, '{http://www.opengis.net/kml/2.2}IconStyle')
            scale = etree.SubElement(icon_style, '{http://www.opengis.net/kml/2.2}scale')
            scale.text = '0'  # Make icon invisible
            
        # Process all placemarks with polygons
        for placemark in self.root.xpath('.//kml:Placemark[kml:Polygon]', namespaces=ns):
            polygon = placemark.find('.//{http://www.opengis.net/kml/2.2}Polygon')
            if polygon is not None:
                # Get polygon coordinates
                coords = polygon.find('.//kml:coordinates', namespaces=ns)
                if coords is not None and coords.text:
                    centroid = self.calculate_centroid(coords.text)
                    
                    # Skip if already has MultiGeometry
                    multi_geom = placemark.find('.//{http://www.opengis.net/kml/2.2}MultiGeometry')
                    if multi_geom is None:
                        multi_geom = etree.Element('{http://www.opengis.net/kml/2.2}MultiGeometry')
                        # Move polygon under MultiGeometry
                        multi_geom.append(polygon)
                        
                        # Add Point for label
                        point = etree.SubElement(multi_geom, '{http://www.opengis.net/kml/2.2}Point')
                        point_coords = etree.SubElement(point, '{http://www.opengis.net/kml/2.2}coordinates')
                        point_coords.text = centroid
                        
                        # Replace original polygon with MultiGeometry
                        for child in list(placemark):
                            if child.tag.endswith('Polygon'):
                                placemark.remove(child)
                        placemark.append(multi_geom)
                        
                        # Add style for invisible point if not already present
                        style_url = placemark.find('kml:styleUrl', namespaces=ns)
                        if style_url is None:
                            style_url = etree.SubElement(placemark, '{http://www.opengis.net/kml/2.2}styleUrl')
                        style_url.text = '#noIconStyle'
        
        print("Converted all polygons to include label points")

    def set_all_polygon_border_colors(self, border_color):
        """
        Set the border (line) color for all polygons in the KML file using a shared style at Document level.
        border_color: HTML hex color (e.g. '#FF0000' or 'red')
        """
        def html_color_to_kml(color):
            # Handle hex colors
            color = color.lstrip('#')
            if len(color) == 6:
                r, g, b = color[0:2], color[2:4], color[4:6]
                return f'ff{b}{g}{r}'  # KML format: aabbggrr (alpha=ff)
            return 'ff0000ff'  # Default red if invalid
            
        kml_color = html_color_to_kml(border_color)
        ns = {'kml': 'http://www.opengis.net/kml/2.2'}
        
        # Ensure we have a Document element at the root level
        document = self.root.find('kml:Document', namespaces=ns)
        if document is None:
            document = etree.Element('{http://www.opengis.net/kml/2.2}Document')
            self.root.insert(0, document)  # Insert as first child
            
        # Create shared style at Document level
        style_id = "shared_polygon_style"
        shared_style = document.find(f'.//kml:Style[@id="{style_id}"]', namespaces=ns)
        
        if shared_style is None:
            # Create new shared style
            shared_style = etree.Element('{http://www.opengis.net/kml/2.2}Style')
            shared_style.set('id', style_id)
            document.insert(0, shared_style)  # Insert at start of Document
            
        # Set up LineStyle
        line_style = shared_style.find('kml:LineStyle', namespaces=ns)
        if line_style is None:
            line_style = etree.SubElement(shared_style, '{http://www.opengis.net/kml/2.2}LineStyle')
            
        # Set color and width
        color_elem = line_style.find('kml:color', namespaces=ns)
        if color_elem is None:
            color_elem = etree.SubElement(line_style, '{http://www.opengis.net/kml/2.2}color')
        color_elem.text = kml_color
            
        width_elem = line_style.find('kml:width', namespaces=ns)
        if width_elem is None:
            width_elem = etree.SubElement(line_style, '{http://www.opengis.net/kml/2.2}width')
        width_elem.text = '2.5'
            
        # Set up PolyStyle
        poly_style = shared_style.find('kml:PolyStyle', namespaces=ns)
        if poly_style is None:
            poly_style = etree.SubElement(shared_style, '{http://www.opengis.net/kml/2.2}PolyStyle')
            
        # Ensure outline is visible and fill is transparent
        outline_elem = poly_style.find('kml:outline', namespaces=ns)
        if outline_elem is None:
            outline_elem = etree.SubElement(poly_style, '{http://www.opengis.net/kml/2.2}outline')
        outline_elem.text = '1'
            
        fill_elem = poly_style.find('kml:fill', namespaces=ns)
        if fill_elem is None:
            fill_elem = etree.SubElement(poly_style, '{http://www.opengis.net/kml/2.2}fill')
        fill_elem.text = '0'  # Make polygons transparent
            
        print(f"Created/updated shared style at Document level with color {kml_color}")
            
        # Apply shared style to all polygons
        for placemark in self.root.xpath('.//kml:Placemark[kml:Polygon]', namespaces=ns):
            name_element = placemark.find('.//{http://www.opengis.net/kml/2.2}name')
            name = name_element.text if name_element is not None else ""
            
            # Check if name contains any English letter
            if re.search(r'[A-Za-z]', name):
                print(f"Updating style for polygon: {name}")
                # Remove any existing inline styles
                for style in placemark.findall('kml:Style', namespaces=ns):
                    placemark.remove(style)
                    
                # Set or update styleUrl
                style_url = placemark.find('kml:styleUrl', namespaces=ns)
                if style_url is None:
                    style_url = etree.SubElement(placemark, '{http://www.opengis.net/kml/2.2}styleUrl')
                style_url.text = f'#{style_id}'
            else:
                print(f"Preserving existing style for polygon: {name}")    

    def preview_excel_updates(self, excel_file_path, polygon_column, image_columns, description_columns, 
                            sheet_name=0):
        """
        Preview what updates would be made from Excel file without actually updating
        
        Args:
            excel_file_path (str): Path to the Excel file
            polygon_column (str): Column name that contains polygon names
            image_columns (str or list): Column name(s) that contain image URLs
            description_columns (str or list): Column name(s) that contain description text  
            sheet_name (str/int): Sheet name or index (default: 0 for first sheet)
        """
        excel_data = self.load_excel_data(excel_file_path, polygon_column, image_columns, 
                                        description_columns, sheet_name)
        
        if not excel_data:
            return
        
        existing_polygons = {p['name']: p for p in self.get_all_polygons()}
        
        print("\n" + "="*60)
        print("PREVIEW: Excel Updates")
        print("="*60)
        
        for polygon_name, data in excel_data.items():
            print(f"\nPolygon: {polygon_name}")
            
            polygon_match = self.find_polygon_by_name_fuzzy(polygon_name)
            
            if not polygon_match:
                print("   ❌ Polygon NOT FOUND in KML file")
                continue
            
            print("   ✅ Polygon found in KML file")
            print(f"   📷 Images to add: {len(data['images'])}")
            for img in data['images'][:3]:  # Show first 3 images
                print(f"      - {img}")
            if len(data['images']) > 3:
                print(f"      ... and {len(data['images']) - 3} more")
            
            print(f"   📝 Description parts to add: {len(data['descriptions'])}")
            for desc in data['descriptions'][:2]:  # Show first 2 descriptions
                print(f"      - {desc[:100]}{'...' if len(desc) > 100 else ''}")
            if len(data['descriptions']) > 2:
                print(f"      ... and {len(data['descriptions']) - 2} more")
        
        print("\n" + "="*60)

    def save_kml(self, output_file=None):
        """Save the modified KML file"""
        if output_file is None:
            output_file = self.kml_file_path
        
        try:
            self.tree.write(output_file, pretty_print=True, xml_declaration=True, encoding='UTF-8')
            print(f"KML file saved to: {output_file}")
            return True
        except Exception as e:
            print(f"Error saving KML file: {e}")
            return False

def interactive_editor():
    """Interactive function to edit polygons"""
    editor = KMLPolygonEditor('MyArea_updated.kml')
    
    while True:
        print("\n" + "="*60)
        print("KML Polygon Editor")
        print("="*60)
        print("1. List all polygons")
        print("2. View polygon details")
        print("3. Edit polygon manually")
        print("4. Preview Excel updates")
        print("5. Update from Excel file")
        print("6. Save KML file")
        print("7. Exit")
        
        choice = input("\nEnter your choice (1-7): ").strip()
        
        if choice == '1':
            editor.list_polygons()
        
        elif choice == '2':
            polygon_name = input("Enter polygon name: ").strip()
            polygon_name = ""
            info = editor.get_polygon_info(polygon_name)
            
            if info:
                print(f"\nPolygon: {info['name']}")
                print(f"Description: {info['description']}")
                print(f"Images in description: {info['images_in_description']}")
                print(f"Media links: {info['media_links']}")
            else:
                print("Polygon not found.")
        
        elif choice == '3':
            polygon_name = input("Enter polygon name to edit: ").strip()
            info = editor.get_polygon_info(polygon_name)
            
            if not info:
                print("Polygon not found.")
                continue
            
            print(f"\nEditing polygon: {info['name']}")
            print(f"Current description: {info['description']}")
            print(f"Current images: {info['images_in_description']}")
            
            # Get new description
            print("\nEnter new description (or press Enter to keep current):")
            new_desc = input().strip()
            if not new_desc:
                # Extract text description without HTML tags
                current_text = re.sub(r'<[^>]+>', '', info['description'])
                current_text = re.sub(r'\s+', ' ', current_text).strip()
                new_desc = current_text
            
            # Get new images
            print("\nEnter image URLs (one per line, empty line to finish):")
            new_images = []
            while True:
                img_url = input().strip()
                if not img_url:
                    break
                new_images.append(img_url)
            
            if not new_images:
                new_images = info['images_in_description']
            
            # Update the polygon
            if editor.update_polygon(polygon_name, new_desc, new_images):
                print("Polygon updated successfully!")
        
        elif choice == '4':
            # Preview Excel updates
            excel_file = input("Enter Excel file path: ").strip()
            if not os.path.exists(excel_file):
                print("Excel file not found.")
                continue
            
            polygon_col = input("Enter polygon column name: ").strip()
            
            # Get image columns
            print("Enter image column names (one per line, empty line to finish):")
            image_cols = []
            while True:
                col = input().strip()
                if not col:
                    break
                image_cols.append(col)
            
            if not image_cols:
                image_cols = [input("Enter at least one image column name: ").strip()]
            
            # Get description columns
            print("Enter description column names (one per line, empty line to finish):")
            desc_cols = []
            while True:
                col = input().strip()
                if not col:
                    break
                desc_cols.append(col)
            
            if not desc_cols:
                desc_cols = [input("Enter at least one description column name: ").strip()]
            
            sheet = input("Enter sheet name/index (press Enter for first sheet): ").strip()
            
            if not sheet:
                sheet = 0
            else:
                try:
                    sheet = int(sheet)
                except ValueError:
                    pass  # Keep as string for sheet name
            
            editor.preview_excel_updates(excel_file, polygon_col, image_cols, desc_cols, sheet)
        
        elif choice == '5':
            # Update from Excel
            excel_file = input("Enter Excel file path: ").strip()
            if not os.path.exists(excel_file):
                print("Excel file not found.")
                continue
            
            polygon_col = input("Enter polygon column name: ").strip()
            
            # Get image columns
            print("Enter image column names (one per line, empty line to finish):")
            image_cols = []
            while True:
                col = input().strip()
                if not col:
                    break
                image_cols.append(col)
            
            if not image_cols:
                image_cols = [input("Enter at least one image column name: ").strip()]
            
            # Get description columns
            print("Enter description column names (one per line, empty line to finish):")
            desc_cols = []
            while True:
                col = input().strip()
                if not col:
                    break
                desc_cols.append(col)
            
            if not desc_cols:
                desc_cols = [input("Enter at least one description column name: ").strip()]
            
            sheet = input("Enter sheet name/index (press Enter for first sheet): ").strip()
            
            if not sheet:
                sheet = 0
            else:
                try:
                    sheet = int(sheet)
                except ValueError:
                    pass  # Keep as string for sheet name
            
            merge_choice = input("Merge with existing data? (y/n, default=y): ").strip().lower()
            merge_with_existing = merge_choice != 'n'
            
            print("\nUpdating polygons from Excel...")
            summary = editor.update_polygons_from_excel(
                excel_file, polygon_col, image_cols, desc_cols, sheet, merge_with_existing
            )
            
            if summary["success"]:
                print(f"\n✅ Update Summary:")
                print(f"   Updated polygons: {len(summary['updated_polygons'])}")
                print(f"   Not found polygons: {len(summary['not_found_polygons'])}")
                print(f"   Total images added: {summary['total_images_added']}")
                print(f"   Total descriptions added: {summary['total_descriptions_added']}")
                
                if summary['updated_polygons']:
                    print(f"\n   Updated: {', '.join(summary['updated_polygons'])}")
                
                if summary['not_found_polygons']:
                    print(f"\n   Not found: {', '.join(summary['not_found_polygons'])}")
            else:
                print(f"❌ Update failed: {summary['message']}")
        
        elif choice == '6':
            save_path = input("Enter save path (or press Enter for current file): ").strip()
            if not save_path:
                save_path = None
            editor.save_kml(save_path)
        
        elif choice == '7':
            print("Goodbye!")
            break
        
        else:
            print("Invalid choice. Please try again.")

# Example usage functions
def example_edit_polygon():
    """Example of how to edit a specific polygon programmatically"""
    editor = KMLPolygonEditor('MyMap.kml')
    
    # List all polygons first
    polygons = editor.list_polygons()
    
    if polygons:
        # Edit the first polygon as an example
        polygon_name = polygons[0]['name']
        
        # New description and images
        new_description = "This is an updated description for the polygon."
        new_images = [
            "https://example.com/image1.jpg",
            "https://example.com/image2.jpg"
        ]
        
        # Update the polygon
        editor.update_polygon(polygon_name, new_description, new_images)
        
        # Save the file
        editor.save_kml('MyMap_updated.kml')

if __name__ == "__main__":
    # Run the interactive editor
    interactive_editor()


