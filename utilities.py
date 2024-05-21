from technology_maps_config import *

# The DualOutput class is used to redirect the output to a file, as well as the console
class DualOutput:
    def __init__(self, filename, std_out, std_err):
        self.file = open(filename, 'w')
        self.stdout = std_out
        self.stderr = std_err

    def write(self, message):
        self.stdout.write(message)
        self.file.write(message)

    def flush(self):
        # Flush both streams
        self.stdout.flush()
        self.file.flush()

    def close(self):
        self.file.close()


# The Utility class contains various utility functions used in the script
class Utility:
    # Formats labels by adding new line after max_chars
    def format_text(self, input_text):
        words = input_text.split()
        lines = []
        current_line = []

        char_count = 0
        for word in words:
            if char_count + len(word) > MAX_CHARS:
                lines.append(' '.join(current_line))
                current_line = []
                char_count = 0

            current_line.append(word)
            char_count += len(word) + 1  # +1 for the space

        # Append any remaining words
        if current_line:
            lines.append(' '.join(current_line))

        return '\n'.join(lines)

    # Stylize the technology divider lines # PPTX !!!!!!!!!!!!!!!!
    def get_line_style(self, line):
        # Setting the attributes of the line
        line.line.color.rgb = self.hex_to_rgb(AM_GREEN)
        line.line.width = Pt(DIV_LINE_WIDTH_PX)
        line.line.dash_style = MSO_LINE_DASH_STYLE.ROUND_DOT

        return line

    # Converts svg files to png so that python-pptx library can handle them
    def convert_svg_to_png(self, input_svg_path, output_png_path):
        cairosvg.svg2png(url=input_svg_path, write_to=output_png_path)

    # Convert hex to RGBColor
    def hex_to_rgb(self, value):
        value = value.lstrip('#')
        length = len(value)
        return RGBColor(int(value[0:length // 3], 16), int(value[length // 3:2 * length // 3], 16),
                        int(value[2 * length // 3:], 16))


# The ImagePacker class is used to resize and pack images within a specified boundary
# The class contains methods to resize images, pack images, and fit logos within a boundary
# The class also contains a method to draw a boundary on the slide for visualization purposes
class ImagePacker:
    # Initialize the ImagePacker with the Utility class and the center x-coordinate
    def __init__(self, utility, center_x):
        self.utility = utility
        self.center_x = center_x

    # Resize an image based on the maximum area and maximum width and height
    def resize_image(self, img_path, max_area, max_width_cm=LOGO_MAX_AREA, max_height_cm=LOGO_MAX_AREA):
        try:
            # Convert SVG to PNG if the image is in SVG format
            if img_path.endswith('.svg'):
                png_path = img_path.replace('.svg', '.png')
                self.utility.convert_svg_to_png(img_path, png_path)
                img_path = png_path

            # Open the image and get the dimensions and aspect ratio of the image
            with Image.open(img_path) as img:
                width, height = img.size
                aspect_ratio = width / height

                # Convert the width and height to centimeters
                width_cm = (width / DPI) * 2.54
                height_cm = width_cm / aspect_ratio
                area_cm = width_cm * height_cm

                # If the area exceeds the maximum area, resize the image
                if area_cm > max_area:
                    area_ratio = area_cm / max_area
                    side_ratio = math.sqrt(area_ratio)
                    width_cm /= side_ratio
                    height_cm /= side_ratio

                # If the width or height exceeds the maximum width or height, resize the image
                if max_width_cm and width_cm > max_width_cm:
                    width_cm = max_width_cm
                    height_cm = width_cm / aspect_ratio

                if max_height_cm and height_cm > max_height_cm:
                    height_cm = max_height_cm
                    width_cm = height_cm * aspect_ratio

                # Return the resized image path and dimensions 
                return img_path, int(width_cm * 100), int(height_cm * 100)
        except Exception as e:
            print(f"Unexpected error encountered while resizing {img_path}: {e}")

    # Resize multiple images based on the maximum area by calling the resize_image method
    def resize_images(self, image_files, max_area):
        return [self.resize_image(img_file, max_area) for img_file in image_files if img_file is not None]

    # Pack images within a specified boundary using 2D bin packing algorithm 
    def pack_images(self, rectangles, max_width, max_height, start_x, start_y):
        # Function to try packing the images within the boundary
        def try_pack(images):
            packed = []
            x = start_x
            y = start_y
            prev_h = 0

            # Convert center_x to cm and calculate direction based on start_x
            center_x = Emu(self.center_x).cm * 100
            direction = 1 if start_x < center_x else -1

            # Iterate through each image data
            for image_data in images:
                if image_data is None:
                    continue

                path, w, h = image_data
                img_padding = max(w, h) * 0.05

                # Add padding to width and height
                w_with_padding = w + 2 * img_padding
                h_with_padding = h + 2 * img_padding

                # Check if image fits horizontally within the boundary, if not reset x and increment y
                if (direction == 1 and x + w_with_padding > start_x + max_width) or (direction == -1 and x - w_with_padding < start_x):
                    x = start_x if direction == 1 else start_x + max_width
                    y += prev_h
                    prev_h = h_with_padding
                else:
                    prev_h = max(prev_h, h_with_padding)

                # Check if image fits vertically within the boundary
                if y + h_with_padding > start_y + max_height:
                    return None

                # Calculate final x position and add to packed list
                final_x = x if direction == 1 else x - w_with_padding
                packed.append((path, final_x, y, w, h))
                x += direction * w_with_padding

            return packed

        # Sort rectangles by area (width * height) in descending order
        shuffled_rectangles = sorted(rectangles, key=lambda item: item[1] * item[2], reverse=True)
        return try_pack(shuffled_rectangles)

    # Fit logos within a specified boundary on the slide using the pack_images method
    def fit_logos(self, slide, image_files, start_x, start_y, width, height):
        packed = False

        if not image_files:
            return

        current_max_area = LOGO_MAX_AREA
        padding_reduction_factor = 1.0

        # Try packing images, reducing max area and padding factor if not successful
        while current_max_area > 0:
            try:
                # Resize images based on current max area
                rectangles = self.resize_images(image_files, current_max_area)
                rectangles = [(path, int(w * padding_reduction_factor), int(h * padding_reduction_factor)) for path, w, h in rectangles if path is not None]

                # Attempt to pack images within the specified boundary
                packed = self.pack_images(rectangles, width * 100, height * 100, start_x * 100, start_y * 100)

                # If images are packed successfully, break the loop
                if packed:
                    break
                
                # Reduce the max area and padding factor for packing the images in the next iteration
                current_max_area -= STEP
                padding_reduction_factor -= 0.1
            except Exception as e:
                print(f"Unexpected error: {e}")
                break

        # If packing failed, print error message
        if not packed or not isinstance(packed, list):
            print("Even after resizing, the images cannot be packed within the specified boundary!")
        else:
            # Add packed images to the slide
            for path, x, y, w, h in packed:
                try:
                    slide.shapes.add_picture(path, Cm(x / 100), Cm(y / 100), Cm(w / 100), Cm(h / 100))
                except Exception as e:
                    print(f"Error adding image from path {path}: {e}")

    # Draw a boundary on the slide for visualization purposes during development and testing
    # The boundaray is drawn to check if the images are packed within the specified boundary
    def draw_boundary(self, slide, start_x, start_y, width, height):
        left = Cm(start_x)
        top = Cm(start_y)

        # Add a rectangle shape to represent the boundary
        boundary_shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            left, top, Cm(width), Cm(height)
        )
        
        # Set the fill to be transparent
        fill = boundary_shape.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(255, 255, 255)
        fill.transparency = 1.0

        # Set the border color and width of the boundary
        line = boundary_shape.line
        line.color.rgb = RGBColor(255, 0, 0)
        line.width = Pt(2.0)
