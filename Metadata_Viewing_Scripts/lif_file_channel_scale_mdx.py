# This file does not contribute to the functionality of the app.
# The purpose of this file is to generate a comprehensive structured json/dictionary 
# mapping out correct channel configurations and scalebar metrics for EVERY image in a LIF file.

import os
from readlif.reader import LifFile

def extract_all_images_metadata(file_path):
    if not os.path.exists(file_path):
        print(f"Error: Could not find the file at '{file_path}'")
        return {}

    try:
        # Load the LIF container
        reader = LifFile(file_path)
        root = reader.xml_root
        
        # This will hold the structured data to pass into your main app
        lif_metadata_manifest = {}
        
        # Step 1: Loop through every individual image series inside the LIF container
        for img_idx in range(reader.num_images):
            img_obj = reader.get_image(img_idx)
            img_name = img_obj.name
            
            # Default fallback data structure per image
            lif_metadata_manifest[img_name] = {
                "image_index": img_idx,
                "dimensions": {
                    "width_pixels": img_obj.dims.x,
                    "height_pixels": img_obj.dims.y,
                    "z_planes": img_obj.dims.z,
                    "time_points": img_obj.dims.t
                },
                "scalebar": {
                    "microns_per_pixel_x": None,
                    "microns_per_pixel_y": None,
                    "unit": "µm"
                },
                "channels": []
            }
            
            # Populate channels with initial positional indexes
            num_channels = img_obj.channels
            for c_idx in range(num_channels):
                lif_metadata_manifest[img_name]["channels"].append({
                    "channel_index": c_idx,
                    "name": f"Channel {c_idx + 1}",
                    "hex_color": "Unknown"
                })

        # Step 2: Parse the XML tree to extract matching scale and channel parameters safely
        for element in root.iter("Element"):
            element_name = element.get("Name")
            
            # Ensure this XML node matches a series we processed
            if element_name in lif_metadata_manifest:
                
                # --- A. SCALEBAR & CALIBRATION EXTRACTION ---
                # Leica records pixel scales under DimensionDescription nodes
                dim_nodes = element.findall(".//DimensionDescription")
                for dim in dim_nodes:
                    dim_id = dim.get("DimID")  # 1 = X-axis, 2 = Y-axis
                    num_elements = dim.get("NumberOfElements") # Total pixels
                    length_m = dim.get("Length")               # Physical size in meters
                    
                    if num_elements and length_m:
                        try:
                            pixels = float(num_elements)
                            length_meters = float(length_m)
                            
                            if pixels > 0:
                                # Convert meters to microns (1m = 1,000,000 microns)
                                length_microns = length_meters * 1_000_000
                                microns_per_pixel = length_microns / pixels
                                
                                if dim_id == "1":   # X-Axis Scale
                                    lif_metadata_manifest[element_name]["scalebar"]["microns_per_pixel_x"] = round(microns_per_pixel, 5)
                                elif dim_id == "2": # Y-Axis Scale
                                    lif_metadata_manifest[element_name]["scalebar"]["microns_per_pixel_y"] = round(microns_per_pixel, 5)
                        except ValueError:
                            pass

                # --- B. CHANNEL DETAILS & RGB HEX COLORS EXTRACTION ---
                channel_nodes = element.findall(".//ChannelDescription")
                for c_idx, ch_node in enumerate(channel_nodes):
                    if c_idx >= len(lif_metadata_manifest[element_name]["channels"]):
                        break
                    
                    # Update real names (e.g., Blue, Green, DAPI, FITC)
                    c_name = ch_node.get("Name") or ch_node.get("LUTName")
                    if c_name:
                        lif_metadata_manifest[element_name]["channels"][c_idx]["name"] = c_name
                    
                    # Compute Hex display colors from Leica's internal signed 32-bit integer system
                    raw_color_val = ch_node.get("Color")
                    if raw_color_val:
                        try:
                            int_color = int(raw_color_val)
                            # Handle negative bit numbers
                            if int_color < 0:
                                int_color = (1 << 32) + int_color
                            
                            # Standardize to a frontend-ready #RRGGBB structure
                            hex_color = f"#{int_color & 0xFFFFFF:06X}"
                            lif_metadata_manifest[element_name]["channels"][c_idx]["hex_color"] = hex_color
                        except ValueError:
                            pass
                            
        return lif_metadata_manifest
    
    except Exception as e:
        print(f"An error occurred while parsing metadata: {e}")
        return {}

# Run the full container diagnostic manifest
file_path = r"IHC input images\.lif\V6.lif"
print(f"Analyzing LIF Container: {file_path}\n" + "="*50)

all_images_meta = extract_all_images_metadata(file_path)

# Print clean console structure mapping
for img_name, meta in all_images_meta.items():
    print(f"\n📸 SERIES NAME: '{img_name}' (Index: {meta['image_index']})")
    print(f"  └── 📏 Size: {meta['dimensions']['width_pixels']}x{meta['dimensions']['height_pixels']} px")
    print(f"  └── 🔬 Scale/Resolution: X: {meta['scalebar']['microns_per_pixel_x']} {meta['scalebar']['unit']}/px | Y: {meta['scalebar']['microns_per_pixel_y']} {meta['scalebar']['unit']}/px")
    print(f"  └── 🎨 Channels:")
    for ch in meta['channels']:
         print(f"      • Index {ch['channel_index']}: {ch['name']} (Hex Color: {ch['hex_color']})")
