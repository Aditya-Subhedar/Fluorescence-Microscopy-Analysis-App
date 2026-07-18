# This file does not contribute to the functionality of the app.
# The purpose of this file is to generate a comprehensive structured json/dictionary 
# mapping out correct channel configurations and scalebar metrics for EVERY image in an OIB file.

import os
from oiffile import OifFile


def extract_oib_metadata(file_path):
    """
    Parses Olympus .oib/.oif containers, extracting structured dimension metadata,
    dynamically calculating spatial scaling from physical bounds, and mapping channels.
    """
    if not os.path.exists(file_path):
        print(f"Error: Could not find the file at '{file_path}'")
        return {}

    try:
        with OifFile(file_path) as oib:
            main_settings = oib.mainfile
            series_axes = oib.axes
            series_shape = oib.shape
            dim_map = dict(zip(series_axes.upper(), series_shape))
            
            image_name = os.path.basename(file_path)
            
            # Application metadata schema layout
            oib_metadata_manifest = {
                image_name: {
                    "image_index": 0,
                    "dimensions": {
                        "width_pixels": dim_map.get('X', 0),
                        "height_pixels": dim_map.get('Y', 0),
                        "z_planes": dim_map.get('Z', 1),
                        "time_points": dim_map.get('T', 1)
                    },
                    "scalebar": {
                        "microns_per_pixel_x": None,
                        "microns_per_pixel_y": None,
                        "unit": "µm"
                    },
                    "channels": []
                }
            }
            
            # --- A. SCALEBAR & CALIBRATION EXTRACTION ---
            # Extract spatial dimensions dynamically using physical boundaries
            for axis_idx, axis_key in [('x', 'Axis 0 Parameters Common'), ('y', 'Axis 1 Parameters Common')]:
                axis_data = main_settings.get(axis_key, {})
                
                # Method 1: Check for explicit pre-computed Scale key
                scale_val = axis_data.get('Scale')
                
                # Method 2: Dynamic formula fallback based on your specific system metadata
                if scale_val is None:
                    try:
                        start_pos = float(axis_data.get('StartPosition', 0.0))
                        end_pos = float(axis_data.get('EndPosition', 0.0))
                        max_size = float(axis_data.get('MaxSize', 0.0))
                        
                        physical_delta = abs(end_pos - start_pos)
                        if physical_delta > 0 and max_size > 0:
                            scale_val = physical_delta / max_size
                    except (ValueError, TypeError):
                        scale_val = None
                
                # Apply rounded precision parameter to the manifest structure
                if scale_val is not None:
                    oib_metadata_manifest[image_name]["scalebar"][f"microns_per_pixel_{axis_idx}"] = round(float(scale_val), 5)

            # --- B. CHANNEL DETAILS & COLORS EXTRACTION ---
            num_channels = dim_map.get('C', 1)
            default_hex_colors = ["#00FF00", "#FF0000", "#0000FF", "#00FFFF", "#FF00FF", "#FFFF00"]

            for c_idx in range(num_channels):
                ch_key = f'Channel {c_idx + 1} Parameters'
                ch_settings = main_settings.get(ch_key, {})
                
                # Clean up display tags and eliminate blank software 'None' strings
                ch_name = ch_settings.get('DyeName') or ch_settings.get('Name')
                if not ch_name or str(ch_name).strip().lower() == 'none':
                    ch_name = f"Channel {c_idx + 1}"
                
                channel_data = {
                    "channel_index": c_idx,
                    "name": ch_name,
                    "hex_color": "Unknown"
                }
                
                # Compute signed bit color masks to #RRGGBB values
                try:
                    raw_color = ch_settings.get('Color')
                    if raw_color is not None:
                        val = int(raw_color)
                        if val < 0:
                            val = (1 << 32) + val
                        channel_data["hex_color"] = f"#{val & 0xFFFFFF:06X}"
                    else:
                        channel_data["hex_color"] = default_hex_colors[c_idx % len(default_hex_colors)]
                except Exception:
                    channel_data["hex_color"] = default_hex_colors[c_idx % len(default_hex_colors)]
                    
                oib_metadata_manifest[image_name]["channels"].append(channel_data)
                
            return oib_metadata_manifest
    
    except Exception as e:
        print(f"An error occurred while parsing OIB metadata: {e}")
        return {}


if __name__ == "__main__":
    file_path = r"  \file_path\... "
    print(f"Analyzing OIB Container: {file_path}\n" + "="*50)
    all_images_meta = extract_oib_metadata(file_path)

    for img_name, meta in all_images_meta.items():
        print(f"\n📸 FILE NAME: '{img_name}' (Index: {meta['image_index']})")
        print(f"  └── 📏 Size: {meta['dimensions']['width_pixels']}x{meta['dimensions']['height_pixels']} px")
        print(f"  └── 🔬 Scale/Resolution: X: {meta['scalebar']['microns_per_pixel_x']} {meta['scalebar']['unit']}/px | Y: {meta['scalebar']['microns_per_pixel_y']} {meta['scalebar']['unit']}/px")
        print(f"  └── 🎨 Channels:")
        for ch in meta['channels']:
             print(f"      • Index {ch['channel_index']}: {ch['name']} (Hex Color: {ch['hex_color']})")
