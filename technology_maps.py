from utilities_pptx import *

def main(material):
    # Create a new map object with the specified material
    map = MapCreator(material)

    # Construct the filename based on current date and time
    current_time = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M")
    filename = os.path.join("logs", f"{material} map output and error log {current_time}.txt")

    # Redirect stdout and stderr to both console and log file
    dual_output = DualOutput(filename, sys.stdout, sys.stderr)
    sys.stdout = dual_output
    sys.stderr = dual_output

    try:
        print("\nUpdating the Charts....")
        
        # Update technology chart
        tech_pie_cat = map.update_technology()

        # Update feedstock chart with the feedstock data
        map.update_feedstock()

        # Update the process chart with the process data
        map.update_fs_process()

        print("\nCreating the presentation.... ")

        # Add the technology dividing lines, and get the coordinates of the end and middle 
        # points of the lines, as well as the coordinates of the technology icons
        # The funcction takes the oval factor as an argument, which is used to determine 
        # the roundness of the pie chart
        tech_lines, end_xlist, end_ylist, mid_xlist, mid_ylist = map.create_technology_lines(0)

        # Add the technology icons
        map.add_technology_icons(mid_xlist, mid_ylist)

        # Add the horizontal connector lines
        map.create_hrline(end_xlist, end_ylist)

        # Place the company logos
        map.place_logos(end_xlist, end_ylist)

        # Add footer text info box
        map.add_header_footer()

    except Exception as e:
        print(f"\nAn error occurred: {e}")

    finally:
        # Save the modified chart as a new PowerPoint file
        output_path = "Map Output\\Test Outputs\\" + material + " Map " + CURRENT_TIME + ".pptx"
        map.prs.save(output_path)
        print(f"\nPresentation saved successfully at {os.getcwd()}\\{output_path}")

        # Ensure log file is closed properly
        if 'dual_output' in locals():
            dual_output.close()

        # Reset stdout and stderr
        sys.stdout = sys.__stdout__
        sys.stderr = sys.__stderr__

# The main function is called when the script is run directly
if __name__ == "__main__":
    # Define the materials to be used
    materials = ["Metal", "Polymer", "Ceramics"]
    
    # Call the main function for the selected material, in this case, metal
    material = materials[0]
    main(material)
