# Automated Presentation Generation Using Python Automation and 3D Packing Algorithms

## Overview

This project introduces an automated PowerPoint presentation system that utilizes Python automation to scrape, filter, and store relevant data from the internet into a MongoDB database. This database ensures data integrity and accessibility throughout the presentation generation process. Later, the system fetches pertinent data from the database to dynamically generate visually stunning slides, complete with graphs and charts, on demand.

Employing complex 3D packing algorithms and a suite of libraries including Selenium, BeautifulSoup4, PyAutoGUI, Python-PPTX, Matplotlib, NumPy, RectPack, and Pandas, seamless execution of tasks was achieved. These tools facilitated efficient data scraping, filtration, and graphical representation generation. Additionally, integration of APIs such as OpenAI bolstered data processing capabilities, enhancing the overall functionality of the system.

## Author

This project was developed by [H R Zitu](https://github.com/Zitu8).

### Associated with AMPOWER GmbH & Co. KG

## Project Structure

The project consists of several Python scripts and configuration files, each playing a specific role in the generation of the technology maps:

1. **technology_maps.py**
2. **technology_maps_config.py**
3. **utilities.py**
4. **utilities_pptx.py**

### 1. technology_maps.py

This is the main script that orchestrates the creation of the technology maps. It initializes the process, updates various charts, and generates the final PowerPoint presentation.

**Key Functions:**
- `main(material)`: 
  - Initializes the map object with the specified material.
  - Sets up logging to both the console and a file.
  - Calls methods to update various charts and elements in the presentation.
  - Saves the final PowerPoint presentation.
  - Handles exceptions and ensures proper cleanup.

### 2. technology_maps_config.py

This file contains the configuration settings and constants used throughout the project. It includes paths to data files, color themes, and various parameters for formatting and layout.

**Key Configurations:**
- **Paths:**
  - `TECH_DATA_PATH`: Path to the technology data JSON file.
  - `SUPPLIERS_DATA_PATH`: Path to the suppliers data JSON file.
  - `TEMPLATE_PATH`: Path to the PowerPoint template file.
  - `MEDIA_PATH`: Directory where logo images are stored.

- **Colors:**
  - `COLOR_THEME`: A dictionary mapping different categories to their respective colors.

- **Feedstock and Feature Orders:**
  - `feedstock_order`: Specifies the order of feedstock types for each material.
  - `FEATURE_ORDER`: Specifies the order of features for the Ceramics material.

- **Dimensions and Layouts:**
  - `LOGO_MAX_AREA`, `FEEDSTOCK_LOGO_WIDTH`, `MARGIN_CM`, etc.: Various settings for sizing and spacing elements within the presentation.

### 3. utilities.py

This file provides utility functions and classes used across the project. It includes logging redirection and various helper methods.

**Key Classes and Functions:**
- `DualOutput`: 
  - Redirects output to both a file and the console.
  - Methods: `write()`, `flush()`, `close()`.

- `Utility`:
  - `format_text(input_text)`: Formats text by adding new lines after a specified number of characters.
  - `get_line_style(line)`: Styles lines in the presentation.
  - `convert_svg_to_png(input_svg_path, output_png_path)`: Converts SVG images to PNG format.
  - `hex_to_rgb(value)`: Converts hex color codes to RGB.

- `ImagePacker`:
  - Manages the resizing and packing of images within a specified boundary.
  - Methods: `resize_image()`, `resize_images()`, `pack_images()`, `fit_logos()`, `draw_boundary()`.

### 4. utilities_pptx.py

This file contains the `MapCreator` class, which manages the creation of the technology map in the PowerPoint presentation.

**Key Functions:**

- **Initialization:**
  - `__init__(self, material)`: Sets up the presentation, loads data from JSON files, and initializes chart and supplier lists.

- **Chart Management:**
  - `get_charts()`: Retrieves chart objects from the slide.
  - `center_charts()`: Centers charts on the slide.
  - `update_technology()`: Updates the technology pie chart with formatted data.
  - `update_feedstock()`: Updates the feedstock chart with sorted data.
  - `update_fs_process()`: Updates the feedstock-process chart based on the material type.

- **Graphical Elements:**
  - `create_technology_lines()`: Creates dividing lines for different technologies.
  - `add_technology_icons()`: Adds icons for each technology.
  - `create_hrline()`: Adds horizontal connector lines.
  - `place_logos()`: Places company logos on the map.
  - `add_header_footer()`: Adds headers, footers, and legends to the slide.
  - `bring_shape_to_front()`, `send_shape_to_back()`: Methods for managing the z-order of shapes.

- **Helper Methods:**
  - `adjust_starting_angle()`, `rotate_data_labels()`, `update_chart_colors()`: Methods for adjusting chart visuals.
  - `get_image_files_for_tech(tech_pie_cat_item, companies)`: Retrieves image file paths for logos based on the technology and supplier data.

## Getting Started

### Prerequisites

- Python 3.x
- Required Python packages: `pptx`, `Pillow`, `rectpack`, `cairosvg`, `lxml`

Install the required packages using pip:

```sh
pip install python-pptx Pillow rectpack cairosvg lxml
```

### Running the Script

1. **Prepare the Data Files**:
   - Ensure the `technology_data.json` and `suppliers_data.json` files are in the same directory as the script.
   - Make sure the PowerPoint template file (e.g., `Metal Template.pptx`) is in place.

2. **Run the Script**:
   - Open a terminal or command prompt.
   - Navigate to the directory containing the script.
   - Execute the script with Python:

```sh
python technology_maps.py
```

### Note on Data and Logos

The data (`technology_data.json` and `suppliers_data.json`) and the logos are typically fetched and cleaned from the database of my employer, AMPOWER GmbH & Co. KG. For privacy and security reasons, this part of the code has been removed.

### Note on Dependencies

This project is designed to be built as a plugin and relies on other projects and components developed by other team members at AMPOWER GmbH & Co. KG. Due to privacy and security considerations, these dependencies are not included, and as such, the code may not work as expected in isolation.

## Customization

### Configuration

To customize the map generation process, modify the configuration values in `technology_maps_config.py`. You can change paths, color themes, feedstock orders, feature orders, and various layout parameters.

### Templates

To use a different PowerPoint template, update the `TEMPLATE_PATH` in `technology_maps_config.py` to point to the new template file.

## Troubleshooting

### Common Issues

- **File Not Found**: Ensure all file paths in `technology_maps_config.py` are correct and files are in place.
- **Python Package Errors**: Ensure all required Python packages are installed.
- **Chart Not Found**: Verify that the template PowerPoint file contains the expected number of charts.

### Logging

Check the log files in the `logs` directory for detailed error messages and execution logs.

## Conclusion

This project provides a comprehensive and automated way to generate technology landscape maps for different materials. By following the instructions and using the provided configuration options, you can customize the output to suit your needs. The project demonstrates advanced techniques in manipulating PowerPoint files using Python libraries and direct XML manipulation, offering valuable insights into automated presentation generation.
