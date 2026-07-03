import os
from pylibCZIrw import czi as pyczi

def extract_czi_images_metadata(file_path):
    if not os.path.exists(file_path):
        print(f"Error: Could not find the file at '{file_path}'")
        return {}

    try:
        with pyczi.open_czi(file_path) as czidoc:
            # 1. Fetch core pixel dimensions from the document bounding box
            # 'C' represents channels, 'Z' is Z-stack planes, 'T' is timepoints
            total_dims = czidoc.total_bounding_box
            width = total_dims.get('X', (0, 0))[1]
            height = total_dims.get('Y', (0, 0))[1]
            z_planes = total_dims.get('Z', (0, 1))[1]
            time_points = total_dims.get('T', (0, 1))[1]
            
            # CZI typically treats multi-scene layouts ('S') as independent image points
            num_images = total_dims.get('S', (0, 1))[1]

            metadata_dict = czidoc.metadata
            
            # Helper Function for Recursive XML Element Search
            def find_keys(data, target_key):
                found = []
                if isinstance(data, dict):
                    for key, value in data.items():
                        if key == target_key:
                            if isinstance(value, list):
                                found.extend(value)
                            else:
                                found.append(value)
                        else:
                            found.extend(find_keys(value, target_key))
                elif isinstance(data, list):
                    for item in data:
                        found.extend(find_keys(item, target_key))
                return found

            # --- 1. Deduplicated Channel Extraction Pipeline ---
            raw_channels = find_keys(metadata_dict, 'Channel')
            unique_channels_dict = {}
            fallback_idx = 0

            for ch in raw_channels:
                # Prioritise nodes containing valid Channel Index IDs
                c_id_raw = ch.get("@Id") or ch.get("Id")
                c_name = ch.get("@Name") or ch.get("Name") or f"Channel {fallback_idx + 1}"
                raw_color = ch.get("Color") or "Unknown"
                
                # Extract clean Index position from structural strings like "Channel:0"
                if c_id_raw and "Channel:" in str(c_id_raw):
                    try:
                        c_idx = int(str(c_id_raw).split(":")[-1])
                    except ValueError:
                        c_idx = fallback_idx
                else:
                    c_idx = fallback_idx

                # Standardise CZI alpha channel colors (#AARRGGBB) to web standard #RRGGBB
                hex_color = "Unknown"
                if raw_color != "Unknown" and str(raw_color).startswith("#"):
                    clean_color = str(raw_color).strip()
                    if len(clean_color) == 9:  # #AARRGGBB -> Remove Alpha channel bytes
                        hex_color = "#" + clean_color[3:]
                    else:
                        hex_color = clean_color

                # Keep entries containing valid indices or more detailed wavelength details
                if c_idx not in unique_channels_dict or ch.get("EmissionWavelength"):
                    unique_channels_dict[c_idx] = {
                        "channel_index": c_idx,
                        "name": c_name,
                        "hex_color": hex_color.upper()
                    }
                
                fallback_idx += 1

            # Convert tracking dict into a sorted index list block
            channel_list = [unique_channels_dict[k] for k in sorted(unique_channels_dict.keys())]

            # If deduplication dropped everything, build basic index fallbacks
            if not channel_list:
                num_channels = total_dims.get('C', (0, 1))[1]
                for c_idx in range(num_channels):
                    channel_list.append({
                        "channel_index": c_idx,
                        "name": f"Channel {c_idx + 1}",
                        "hex_color": "Unknown"
                    })

            # --- 2. Scaling / Distance Extraction Pipeline ---
            raw_distances = find_keys(metadata_dict, 'Distance')
            microns_x = None
            microns_y = None
            
            for dist in raw_distances:
                axis_id = dist.get("@Id") or dist.get("Id")
                val_str = dist.get("Value")
                
                if axis_id and val_str:
                    try:
                        val_meters = float(val_str)
                        val_um = round(val_meters * 1e6, 5)  # Convert to micrometers
                        
                        if str(axis_id).upper() == 'X':
                            microns_x = val_um
                        elif str(axis_id).upper() == 'Y':
                            microns_y = val_um
                    except ValueError:
                        pass

            # --- 3. Build the Universal Structured Output Manifest ---
            czi_metadata_manifest = {}
            for img_idx in range(num_images):
                # Match image point naming scheme from LIF and ND2 scripts
                img_name = f"Series_{img_idx + 1}" if num_images > 1 else os.path.basename(file_path)
                
                czi_metadata_manifest[img_name] = {
                    "image_index": img_idx,
                    "dimensions": {
                        "width_pixels": width,
                        "height_pixels": height,
                        "z_planes": z_planes,
                        "time_points": time_points
                    },
                    "scalebar": {
                        "microns_per_pixel_x": microns_x,
                        "microns_per_pixel_y": microns_y,
                        "unit": "µm"
                    },
                    "channels": channel_list
                }
                
            return czi_metadata_manifest

    except Exception as e:
        print(f"An error occurred while parsing CZI metadata: {e}")
        return {}


# --- Main Console Execution Loop ---
if __name__ == "__main__":
    file_path = r"IHC input images\.czi\3 Channel\10_A1.czi"
    print(f"Analyzing CZI File: {file_path}\n" + "="*50)
    
    all_images_meta = extract_czi_images_metadata(file_path)
    
    # Print out identical clean layout matching your ND2 console design
    for img_name, meta in all_images_meta.items():
        print(f"\n📸 SERIES NAME: '{img_name}' (Index: {meta['image_index']})")
        print(f"  └── 📏 Size: {meta['dimensions']['width_pixels']}x{meta['dimensions']['height_pixels']} px")
        print(f"  └── 🔬 Scale/Resolution: X: {meta['scalebar']['microns_per_pixel_x']} {meta['scalebar']['unit']}/px | Y: {meta['scalebar']['microns_per_pixel_y']} {meta['scalebar']['unit']}/px")
        print(f"  └── 🎨 Channels:")
        for ch in meta['channels']:
             print(f"      • Index {ch['channel_index']}: {ch['name']} (Hex Color: {ch['hex_color']})")
