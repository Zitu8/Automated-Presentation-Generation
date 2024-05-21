from utilities import *

# Thhe class MapCreator is used to create a map of the technology landscape based on the provided material.
class MapCreator:
    def __init__(self, material):
        # Create a utility object and initialize the material
        self.utility = Utility()
        self.material = material
        self.template_path = TEMPLATE_PATH

        # Calculate Template Slide dimensions
        self.prs = Presentation(self.template_path)
        self.slide = self.prs.slides[0]
        self.slide_width = self.prs.slide_width
        self.slide_height = self.prs.slide_height

        # Set up the presentation
        self.center_charts()
        charts = self.get_charts()

        print("Analysing the PPTX file....")

        if len(charts) != 4:
            print("All charts are not found! Something might be wrong with the provided template.")
            print("Terminating the script...")
            exit()

        # Assign the charts to the class variables for easy access and manipulation
        self.tech_pie, self.fs_pie, self.fs_process, self.fs_line = charts

        print("Reading data from the JSON files....")

        # Load the JSON data from the technology data file
        with open(TECH_DATA_PATH, "r") as file:
            self.tech_json = json.load(file)

        # Create the technology_data list
        self.tech_list = [item["Technology"] for item in self.tech_json]
        print(f"The technologies: {self.tech_list}")
        self.number_of_tech = len(self.tech_list)
        print(f"Number of Technologies: {self.number_of_tech}")

        # Load the JSON data from the suppliers data file
        with open(SUPPLIERS_DATA_PATH, "r") as file:
            self.suppliers_json = json.load(file)

        # Create the suppliers list with unique OEM names
        unique_companies = set(item["OEM"] for item in self.suppliers_json)
        self.suppliers_list = list(unique_companies)
        self.number_of_suppliers = len(self.suppliers_list)
        print(f"Number of Companies: {self.number_of_suppliers}")

    # Return charts from the slide
    def get_charts(self):
        shape_array = []

        for shape in self.slide.shapes:
            if shape.has_chart:
                shape_array.append(shape)

        return shape_array


    # Center the charts on the slide horizontally and vertically for better alignment
    def center_charts(self):
        slide_center_x = self.slide_width // 2
        slide_center_y = self.slide_height // 2 + Cm(1).emu

        for chart in self.get_charts():
            # Calculate chart's half-width and half-height
            half_chart_width = chart.width // 2
            half_chart_height = chart.height // 2

            # Calculate new position
            new_left = slide_center_x - half_chart_width
            new_top = slide_center_y - half_chart_height

            # Convert new_left and new_top to integer if they are not
            new_left = int(new_left)
            new_top = int(new_top)

            # Move chart to the new position
            chart.left = new_left
            chart.top = new_top

    # Updates the technology pie chart
    def update_technology(self):
        # Set Technology pie chart (tech_pie) data
        tech_pie_cat_formatted = []
        tech_pie_values = []

        for tech in self.tech_list:
            tech_pie_cat_formatted.append(self.utility.format_text(tech))
            tech_pie_values.append(1)  # since each technology is considered unique and given a count of 1

        # Change Technology pie chart (tech_pie) data
        chart_data = CategoryChartData()
        chart_data.categories = tech_pie_cat_formatted
        chart_data.add_series('# Technologies', tech_pie_values)
        self.tech_pie.chart.replace_data(chart_data)

    # Updates the  feedstock chart
    def update_feedstock(self):
        feedstock_data = defaultdict(list)

        for item in self.tech_json:
            map_features = json.loads(item['MAP_Feature'])
            if item['Feedstock'] == 'Liquid':
                # For 'Liquid' feedstock, create keys combining Feedstock and MAP_Feature
                for feature in map_features:
                    key = (item['Feedstock'], feature)
                    feedstock_data[key].append(item)
            else:
                # For other feedstocks, use only Feedstock as the key
                key = item['Feedstock']
                feedstock_data[key].append(item)

        output_feedstock_data = []
        powder_sinter_count = 0

        for key, items in feedstock_data.items():
            count = len(items)

            # Handle key unpacking based on its type (string or tuple)
            if isinstance(key, tuple):
                feedstock, feature = key
            else:
                feedstock = key
                feature = None

            output_feedstock_data.append((feedstock, count))

            # Counting Powder Sinter-based
            if feedstock == 'Powder':
                if feature == "Sinter-based":
                    powder_sinter_count += count
                elif feature is None:
                    # For non-Liquid materials, check each item's MAP_Feature
                    for item in items:
                        if "Sinter-based" in json.loads(item['MAP_Feature']):
                            powder_sinter_count += 1


        # Sort the feedstock data based on the order defined in feedstock_order
        material_order = feedstock_order.get(self.material, [])
        output_feedstock_data.sort(key=lambda x: -x[1])
        if material_order:
            output_feedstock_data.sort(
                key=lambda x: material_order.index(x[0]) if x[0] in material_order else len(material_order))

        # Change the chart data for fs_pie and fs_line
        chart_data = CategoryChartData()
        chart_data.categories = [item[0] for item in output_feedstock_data]
        chart_data.add_series('# Technologies', [item[1] for item in output_feedstock_data])

        self.fs_pie.chart.replace_data(chart_data)
        self.fs_line.chart.replace_data(chart_data)

        # Adjust the starting angle for Metal material
        if self.material == "Metal":
            self.adjust_starting_angle(self.fs_pie, powder_sinter_count)
            self.adjust_starting_angle(self.fs_line, powder_sinter_count)

        # Rotate the data labels and update the colors for the feedstock's pie chart
        self.rotate_data_labels(self.fs_pie.chart, [item[1] for item in output_feedstock_data])
        self.update_chart_colors(self.fs_pie.chart, [item[0] for item in output_feedstock_data])

        print("Sorted Feedstock chart data for {}: {}".format(self.material, ', '.join(
            f"'{feedstock}': {count}" for feedstock, count in output_feedstock_data)))

    # Adjust the starting angle for the doughnut chart based on the feedstock data
    def adjust_starting_angle(self, chart_name, pow_sin_count=2):
        # Calculate the angle for the first slice for Metal material
        metal_fs_angle = int(360 - (360 / self.number_of_tech) * pow_sin_count)

        doughnutChart = chart_name.chart.plots[0]._element
        firstSliceAngs = doughnutChart.xpath('./c:firstSliceAng')

        if len(firstSliceAngs) == 0:
            # If the firstSliceAng element is not found, we need to create one
            from lxml.etree import SubElement
            firstSliceAng = SubElement(doughnutChart,'{http://schemas.openxmlformats.org/drawingml/2006/chart}firstSliceAng')
        else:
            firstSliceAng = firstSliceAngs[0]
            
        # Set the value of the firstSliceAng element to the calculated angle    
        firstSliceAng.set('val', str(metal_fs_angle))  


    # Adjust the data label rotation angles for the chart based on the values
    def rotate_data_labels(self, chart, values, pow_sin_count=2):
        angles = []
        angle_per_slice = 360 / sum(values)
        sum_of_angles = 0

        # # Adjust the starting point for metal material
        if self.material == "Metal":
            sum_of_angles = -(360 / self.number_of_tech) * pow_sin_count  

        # calculate data label rotation angles for each value
        for i, val in enumerate(values):
            angle = sum_of_angles + angle_per_slice * values[i] / 2  # Calculating the middle angle for each category

            if angle < 180:
                rot_angle = angle - 90
            elif angle < 270:
                rot_angle = angle - 270
            else:
                rot_angle = angle % 90

            # Convert the angle to the format required by PowerPoint and add it to the angles list
            angles.append(int(rot_angle * 60000))
            sum_of_angles += angle_per_slice * values[i]  # Move to the next category

        # Rotate the data labels for the chart based on the calculated angles
        counter = 0
        element = chart.plots[0].series[0]._element
        for i in range(0, len(element)):
            if re.sub("{.*}", '', element[i].tag) == "dLbls":
                # for the first element
                txPr = element[i].get_or_add_txPr()
                txPr.bodyPr.set('rot', str(angles[counter]))
                counter += 1

                # for all other elements
                for element2 in element[i]:
                    if re.sub("{.*}", '', element2.tag) == "dLbl":
                        txPr = element2.get_or_add_txPr()

                        try:
                            txPr.bodyPr.set('rot', str(angles[counter]))
                            counter += 1
                        except Exception as e:
                            continue

    # Update the colors of the chart based on the categories
    def update_chart_colors(self, chart, categories, chart_type="Feedstock"):
        cat_flag = 0

        # Go through every category and modify the color for Feedstock
        if chart_type == "Feedstock":
            for idx, cat in enumerate(categories):
                if cat == "Liquid":
                    cat_flag += 1

                # Blue is the default color if nothing else is found
                color = COLOR_THEME.get(cat, AM_BLUE)  

                if cat == "Liquid" and cat_flag > 1:
                    color = AM_GREEN  # For the second Liquid

                # Change the color of the point in the chart 
                point = chart.series[0].points[idx]
                point.format.fill.solid()
                point.format.fill.fore_color.rgb = RGBColor.from_string(color)

        # If the chart type is 'Process' color comes from map feature
        elif chart_type == "Process":
            map_feature = []

            # Change color according to the predefined MAP_Feature
            for idx, cat in enumerate(categories):
                # Find the matching entry in the tech_json where 'Technology' is equal to the category
                matching_entry = next((item for item in self.tech_json if item.get("Technology") == cat), None)

                # If a matching entry is found, print the category and length of MAP_Feature
                if matching_entry and matching_entry.get("MAP_Feature", "[]"):
                    map_feature = json.loads(matching_entry.get("MAP_Feature", "[]"))  # Default to empty list if none

                # Color for the features
                last_series = len(chart.series) - 1
                for feature in map_feature:
                    color = COLOR_THEME.get(feature, AM_BLUE)  # Blue is the default color if nothing else is found
                    if "Other" in cat:
                        color = AM_GREY  # Grey is the default color for Other

                    point = chart.series[last_series].points[idx]
                    point.format.fill.solid()
                    point.format.fill.fore_color.rgb = RGBColor.from_string(color)
                    last_series -= 1

    # Updates the fs_process chart
    def update_fs_process(self):
        # Check if the material is 'Polymer' and remove the fs_process chart
        if self.material == 'Polymer':
            # Assuming the fs_process chart is the third chart in the slide
            fs_process_chart_shape = self.get_charts()[2]
            self.slide.shapes._spTree.remove(fs_process_chart_shape._element)

            print("Polymer material detected - fs_process chart has been removed.")
            return

        # Check if the material is 'Ceramics' and set the fs_process_data for each tech with values [1, 1, 1]
        elif self.material == 'Ceramics':
            fs_process_data = {tech: [1, 1, 1] for tech in self.tech_list}
            print(f"Feedstock-Process chart data for Ceramics: {fs_process_data}")

            # Adjust the size of the doughnut hole
            doughnutChart = self.fs_process.chart.plots[0]._element
            holeSizes = doughnutChart.xpath('./c:holeSize')
            if len(holeSizes) == 0:
                raise ValueError('sorry, no c:holeSize element present')
            holeSize = holeSizes[0]
            holeSize.set('val', '85')

        # Check if the material is 'Metal' and set the fs_process_data for each tech with value 1
        elif self.material == 'Metal':
            fs_process_data = {tech: 1 for tech in self.tech_list}
            print(f"Feedstock-Process chart data for Metal: {fs_process_data}")

        else:
            print("Unknown Material!")
            return

        # Set fs_process_cat and fs_process_values
        fs_process_cat = []
        fs_process_values = []
        for key, value in fs_process_data.items():
            fs_process_cat.append(key)
            # Check if the value is a list (as it is for Ceramics), if not, simply append the value
            fs_process_values.append(value if isinstance(value, list) else [value])

        # Change the chart data for fs_process
        chart_data = CategoryChartData()
        chart_data.categories = fs_process_cat
        for i in range(len(fs_process_values[0])):  # Assuming all values are lists of the same length
            chart_data.add_series(f'# Technologies {i + 1}', [val[i] for val in fs_process_values])
        self.fs_process.chart.replace_data(chart_data)

        # Update the colors of the process charts
        self.update_chart_colors(self.fs_process.chart, fs_process_cat, "Process")

    # Create technology divider lines
    def create_technology_lines(self, oval_factor):
        # Arrays to store the endpoint values for the technology lines and the technology icons
        end_xlist = []
        end_ylist = []
        tech_lines = []
        mid_xlist = []
        mid_ylist = []

        # Calculate the angles and center points
        angle = 360 / self.number_of_tech  # Angle start from the right
        r = math.ceil(Emu(self.tech_pie.width).cm / 2) - MARGIN_CM - PADDING_CM  # The radius of the technology lines
        mid_r = r * 0.7  # 70% of the length of radius of the technology lines
        center_x_fs_pie = self.fs_pie.left + self.fs_pie.width / 2
        center_y_fs_pie = self.fs_pie.top + self.fs_pie.height / 2
        centre_x = Emu(center_x_fs_pie).cm
        centre_y = Emu(center_y_fs_pie).cm

        # Oval
        rc = r
        quarter = len(self.tech_list) / 4

        # Draw a line for each technology
        for i, tech in enumerate(self.tech_list):
            mod = (len(self.tech_list) / 20) * oval_factor

            if i == 0 or i == quarter * 2:
                rc = r + (1.5 * mod)
            elif i < quarter * 2:
                if i < quarter:
                    rc -= mod  # First Quarter
                else:
                    rc += mod  # Second Quarter

                if rc > r:
                    rc = r
            else:
                if i < quarter * 3:
                    rc -= mod  # Third Quarter
                else:
                    rc += mod  # Fourth Quarter

                if rc > r:
                    rc = r

            theta = math.radians(i * angle + 270)  # Adding 180 degree to start drawing line from left
            delta = math.radians(i * angle + 270 + angle / 2)

            # Adding the endpoints to the array
            end_x = centre_x + rc * math.cos(theta)
            end_y = centre_y + rc * math.sin(theta)
            end_xlist.append(end_x)
            end_ylist.append(end_y)

            # Adding the endpoints to the array
            mid_x = centre_x + mid_r * math.cos(delta)
            mid_y = centre_y + mid_r * math.sin(delta)
            mid_xlist.append(mid_x)
            mid_ylist.append(mid_y)

            # Stylise the lines and add the lines to the tech_lines array
            line = self.slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(centre_x), Cm(centre_y), Cm(end_x), Cm(end_y))
            self.utility.get_line_style(line)
            self.send_shape_to_back(line)
            tech_lines.append(line)

        return tech_lines, end_xlist, end_ylist, mid_xlist, mid_ylist

    # Add the technology icons
    def add_technology_icons(self,  mid_xlist, mid_ylist):
        icon_path = None

        # Get the icon path for the given tech_name
        try:
            for item in self.tech_json:
                tech = item["Technology"]
                icon_path = item["Logo"]
                position = item["MAP_Position"] - 1  # As arrays starts from 0

                # Construct the full image path
                img_path = os.path.join(MEDIA_PATH, icon_path)

                if not os.path.exists(img_path):
                    print(f"Image for {tech} not found. Using a default image.")
                    img_path = os.path.join(MEDIA_PATH, "2023\\10\\other.png")

                # Check if the image is an SVG and convert it to PNG if it is.
                if img_path.endswith('.svg'):
                    png_path = os.path.splitext(img_path)[0] + ".png"
                    self.utility.convert_svg_to_png(img_path, png_path)
                    img_path = png_path  # update img_path to point to the newly created PNG

                # Add the feedstock logo to the slide
                feedstock_logo = self.slide.shapes.add_picture(img_path, Cm(mid_xlist[position]),
                                                               Cm(mid_ylist[position]))

                # Check which dimension (width or height) is larger and resize accordingly
                if feedstock_logo.height > feedstock_logo.width:
                    aspect_ratio = feedstock_logo.width / feedstock_logo.height
                    feedstock_logo.height = FEEDSTOCK_LOGO_WIDTH
                    feedstock_logo.width = int(FEEDSTOCK_LOGO_WIDTH * aspect_ratio)
                else:
                    aspect_ratio = feedstock_logo.height / feedstock_logo.width
                    feedstock_logo.width = FEEDSTOCK_LOGO_WIDTH
                    feedstock_logo.height = int(FEEDSTOCK_LOGO_WIDTH * aspect_ratio)

                # Move the logo to the center of its position
                feedstock_logo.left = int(Cm(mid_xlist[position]) - feedstock_logo.width / 2)
                feedstock_logo.top = int(Cm(mid_ylist[position]) - feedstock_logo.height / 2)

        except Exception as e:
            print(e)

    # Creates the horizontal lines connected to the technology divider lines
    def create_hrline(self, end_xlist, end_ylist):
        for i in range(len(self.tech_list)):
            # Skip drawing the horizontal line for the top and bottom technology lines
            if end_ylist[i] == max(end_ylist) or end_ylist[i] == min(end_ylist):
                continue

            # Drawing the lines on the right side
            center_x = self.fs_pie.left + self.fs_pie.width / 2
            if end_xlist[i] > Emu(center_x).cm:
                line = self.slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(end_xlist[i]), Cm(end_ylist[i]),
                                                       self.slide_width - Cm(MARGIN_CM), Cm(end_ylist[i]))
            # Drawing the lines on the left side
            else:
                line = self.slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(end_xlist[i]), Cm(end_ylist[i]),
                                                       Cm(MARGIN_CM), Cm(end_ylist[i]))

            self.utility.get_line_style(line)
            self.send_shape_to_back(line)

    # Creates the horizontal lines connected to the technology divider lines
    def place_logos(self, end_xlist, end_ylist):
        center_x_fs_pie = self.fs_pie.left + self.fs_pie.width / 2
        center_y_fs_pie = self.fs_pie.top + self.fs_pie.height / 2

        image_packer = ImagePacker(self.utility, center_x_fs_pie)
        right_x = end_xlist[1]
        right_width = Emu(self.slide_width).cm - right_x - MARGIN_CM

        for i, tech in enumerate(self.tech_list):
            image_files = self.get_image_files_for_tech(tech, self.suppliers_json)

            if len(image_files) == 0:
                print(f"No images found for {tech}!")
            else:
                print(f"{len(image_files)} images found for {tech}.")

            # Calculate the logo placement area
            if i == 0 or i == math.floor(self.number_of_tech / 2):
                # Placing logos in the top right part
                if i == 0:
                    if i < self.number_of_tech - 1:
                        width = right_width - MARGIN_CM
                        height = MIN_HEIGHT
                        image_packer.fit_logos(self.slide, image_files, right_x + MARGIN_CM,
                                               end_ylist[1] - MIN_HEIGHT - PADDING_CM, width, height)

                # For the bottom left part
                if i == math.floor(self.number_of_tech / 2):
                    width = end_xlist[i+1] - MARGIN_CM * 2
                    height = end_ylist[i] - end_ylist[i + 1] - PADDING_CM * 2
                    height = MIN_HEIGHT - PADDING_CM
                    image_packer.fit_logos(self.slide, image_files, MARGIN_CM, end_ylist[i + 1] + PADDING_CM, width,
                                           height)
                continue

            # Calculate the logo placement area on the right side
            if end_xlist[i] > Emu(center_x_fs_pie).cm:
                # Placing on the right side
                if end_ylist[i] <= Emu(center_y_fs_pie).cm - 1:
                    right_x = end_xlist[i] + LOGO_OFFSET_CM
                else:
                    right_x = end_xlist[i] + LOGO_OFFSET_CM - MARGIN_CM

                width = Emu(self.slide_width).cm - right_x - MARGIN_CM
                height = end_ylist[i + 1] - end_ylist[i] - PADDING_CM * 2

                # For the last element at right side
                if height < 1:
                    height = MIN_HEIGHT - PADDING_CM

                image_packer.fit_logos(self.slide, image_files, right_x, end_ylist[i] + PADDING_CM, width, height)

            # Drawing the boundary lines on the left side
            else:
                if i < self.number_of_tech - 1:
                    if end_ylist[i] <= Emu(center_y_fs_pie).cm + 1:
                        width = end_xlist[i] - LOGO_OFFSET_CM
                    else:
                        width = end_xlist[i + 1] - LOGO_OFFSET_CM

                    height = end_ylist[i] - end_ylist[i + 1] - PADDING_CM * 2
                    image_packer.fit_logos(self.slide, image_files, MARGIN_CM, end_ylist[i + 1] + PADDING_CM, width, height)
                else:
                    # For the last one (index wise) on left side
                    width = end_xlist[i] - MARGIN_CM * 2
                    height = MIN_HEIGHT
                    image_packer.fit_logos(self.slide, image_files, MARGIN_CM,  end_ylist[1] - MIN_HEIGHT - PADDING_CM, width, height)


    # Get image files for each tech items
    def get_image_files_for_tech(self, tech_pie_cat_item, companies):
        # Extract image file paths based on the criteria
        image_files = []
        for suppliers in companies:
            if suppliers['Material_Cluster'] == self.material and suppliers['Technology'] == tech_pie_cat_item and suppliers['Logo']:
                image_files.append(os.path.join(MEDIA_PATH, suppliers['Logo']))

        return image_files

    # Creates header and footer of the slide
    def add_header_footer(self):
        # Assign the header and footer text
        header_text = self.material + " Additive Manufacturing technology landscape"
        version = "Version " + datetime.datetime.now().strftime("%d. %B %Y")

        # Update the Header and footer text
        # Loop through each shape in the slide
        for shape in self.slide.shapes:
            # Check if the shape is a text box
            if shape.has_text_frame:
                # Get the text frame
                text_frame = shape.text_frame

                # If "Template" is in the text frame it is header
                if "Template" in text_frame.text:
                    for paragraph in text_frame.paragraphs:
                        for run in paragraph.runs:
                            if "Template" in run.text:
                                # Replace the heading with header text
                                run.text = header_text
                                print("\nHeader text has been updated.")

                # If "Download" is in the text frame it is footer
                if "Download" in text_frame.text:
                    for paragraph in text_frame.paragraphs:
                        for run in paragraph.runs:
                            if "Version" in run.text:
                                run.text = version

                            if "Number of technologies" in run.text:
                                run.text = "Number of technologies: " + str(self.number_of_tech)

                            if "Number of suppliers" in run.text:
                                run.text = "Number of suppliers: " + str(self.number_of_suppliers)

                    print("Footer text has been updated.")

        self.add_footer_legend()

        # Shape IDs of the green bars to change the color
        shape_ids_to_change = [10, 12]

        # Loop through each shape in the slide and change color if the ID matches
        for shape in self.slide.shapes:
            if shape.shape_id in shape_ids_to_change:
                if shape.fill:
                    shape.fill.solid()
                    shape.fill.fore_color.rgb = RGBColor.from_string(COLOR_THEME.get(self.material))

    # Add the footer legend
    def add_footer_legend(self):
        # Extract unique features and sort them
        unique_features = set()
        for tech_data in self.tech_json:
            map_feature = tech_data.get("MAP_Feature", "")
            # Check if the map_feature is a non-empty string
            if isinstance(map_feature, str) and map_feature:
                # Remove "[" and "]" characters, split the string by ",", strip whitespace, and add to unique_features
                features_list = [feature.strip(' "') for feature in map_feature.strip("[]").split(",")]
                unique_features.update(features_list)

        unique_features = sorted(unique_features,
                                 key=lambda x: FEATURE_ORDER.index(x) if x in FEATURE_ORDER else len(FEATURE_ORDER))

        # Calculate starting position to center the elements
        total_width = sum([Cm(len(f) * 0.5) for f in unique_features])  # Estimate total width
        total_width += (len(unique_features) - 1) * Cm(0.5)  # Add spaces between elements
        current_left = (self.slide_width - total_width) / 2

        # Position the circles and text
        for feature in unique_features:
            # Add a circle (autoshape) for each feature
            top = self.slide_height - Cm(2)

            # Draw the circle without a border and shadow
            circle = self.slide.shapes.add_shape(MSO_SHAPE.OVAL, current_left, top, CIRCLE_DIAMETER, CIRCLE_DIAMETER)
            fill = circle.fill
            fill.solid()
            fill.fore_color.rgb = RGBColor.from_string(COLOR_THEME.get(feature, AM_BLUE))

            # Remove the border of the circle and shadow effects
            circle.line.fill.background()
            circle.shadow.inherit = False

            # Move the left position for the text
            current_left += CIRCLE_DIAMETER + Cm(0.25)

            # Add a textbox with the feature name
            textbox = self.slide.shapes.add_textbox(current_left, top - CIRCLE_DIAMETER / 10, width=Cm(5),
                                               height=CIRCLE_DIAMETER)
            text_frame = textbox.text_frame
            text_frame.margin_top = 0
            text_frame.margin_bottom = 0
            text_frame.text_anchor = MSO_ANCHOR.MIDDLE

            p = text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT  # Align text left
            run = p.add_run()
            run.text = feature

            # Set the font properties
            font = run.font
            font.size = LEGEND_FONT_SIZE
            font.name = LEGEND_FONT
            font.color.rgb = RGBColor.from_string(AM_BLUE)

            # Increment left position for the next element
            current_left += Cm(len(feature) * 0.5) + Cm(0.25)

        print("Chart Legends has been added.")

    # Brings the provided shape to the front of the z-order
    def bring_shape_to_front(self, shape):
        spTree = self.slide.shapes._spTree
        shape_elm = shape.element
        spTree.remove(shape_elm)
        spTree.insert(-1, shape_elm)

    # Sends the provided shape to the back of the z-order.
    def send_shape_to_back(self, shape):
        spTree = self.slide.shapes._spTree
        shape_elm = shape.element
        spTree.remove(shape_elm)
        spTree.insert(2, shape_elm)  # Inserting at 2, considering the usual spTree structure

