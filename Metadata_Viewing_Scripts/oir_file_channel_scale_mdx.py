import os

try:
    import oirfile
except ImportError:
    raise ImportError("The 'oirfile' library is required. Please run '!pip install oirfile' first.")

def extract_oir_metadata(file_path):
    if not os.path.exists(file_path):
        print(f"Error: Could not find the file at '{file_path}'")
        return {}

    image_name = os.path.basename(file_path)

    manifest = {
        image_name: {
            "image_index": 0,
            "dimensions": {
                "width_pixels": 0,
                "height_pixels": 0,
                "z_planes": 1,
                "time_points": 1,
            },
            "scalebar": {
                "microns_per_pixel_x": None,
                "microns_per_pixel_y": None,
                "unit": "µm",
            },
            "channels": [],
        }
    }

    try:
        with oirfile.OirFile(file_path) as oir:
            manifest_core = manifest[image_name]
            
            # --- A. EXTRACT DYNAMIC DIMENSIONS ---
            sizes = oir.sizes
            manifest_core["dimensions"]["width_pixels"] = sizes.get("X", 0)
            manifest_core["dimensions"]["height_pixels"] = sizes.get("Y", 0)
            manifest_core["dimensions"]["z_planes"] = sizes.get("Z", 1)
            manifest_core["dimensions"]["time_points"] = sizes.get("T", 1)

            # --- B. EXTRACT SPATIAL CALIBRATIONS ---
            scales = oir.coord_scales
            units = oir.coord_units
            
            pixel_size_x = scales.get("X")
            pixel_size_y = scales.get("Y")
            
            if pixel_size_x is not None:
                manifest_core["scalebar"]["microns_per_pixel_x"] = round(float(pixel_size_x), 5)
            if pixel_size_y is not None:
                manifest_core["scalebar"]["microns_per_pixel_y"] = round(float(pixel_size_y), 5)
            
            unit_x = units.get("X", "µm")
            if unit_x:
                unit_str = str(unit_x).lower()
                if unit_str in ["micrometer", "micrometers", "um"]:
                    manifest_core["scalebar"]["unit"] = "µm"
                else:
                    manifest_core["scalebar"]["unit"] = str(unit_x)

            # --- C. CHANNEL GENERATION (TRUE HARDWARE METADATA) ---
            actual_channel_count = sizes.get("C", 1)
            
            for c_idx in range(actual_channel_count):
                ch_name = f"Channel {c_idx + 1}"
                assigned_color = None
                
                if c_idx < len(oir.channels):
                    ch = oir.channels[c_idx]
                    
                    # 1. Get the true channel name
                    ch_name = getattr(ch, "name", None) or ch_name
                    
                    # 2. Extract the actual hardware color integer saved by Olympus
                    raw_color = getattr(ch, "color", None)
                    
                    if raw_color is not None:
                        try:
                            if isinstance(raw_color, int):
                                # Convert raw color integer to standard Hex, masking out the alpha channel
                                assigned_color = f"#{raw_color & 0xFFFFFF:06X}"
                            elif isinstance(raw_color, (tuple, list)) and len(raw_color) >= 3:
                                # In case oirfile parses it as an RGB tuple
                                assigned_color = f"#{int(raw_color[0]):02X}{int(raw_color[1]):02X}{int(raw_color[2]):02X}"
                        except Exception:
                            pass
                
                # 3. Absolute fallback only if the microscope saved no color data at all
                if assigned_color is None:
                    fallback_colors = ["#00FF00", "#FF0000", "#0000FF", "#00FFFF", "#FF00FF"]
                    assigned_color = fallback_colors[c_idx % len(fallback_colors)]
                
                manifest_core["channels"].append({
                    "channel_index": c_idx,
                    "name": ch_name,
                    "hex_color": assigned_color
                })

    except Exception as e:
        print(f"Failed to decode file properties using OirFile engine: {e}")

    return manifest


if __name__ == "__main__":
    file_path = r"  \file_path\...  "
    
    print(f"Analyzing OIR via oirfile Pipeline: {file_path}\n" + "=" * 50)
    all_images_meta = extract_oir_metadata(file_path)

    for img_name, meta in all_images_meta.items():
        print(f"\n📸 FILE NAME: '{img_name}' (Index: {meta['image_index']})")
        print(f"  └── 📏 Size: {meta['dimensions']['width_pixels']}x{meta['dimensions']['height_pixels']} px")
        print(f"  └── 🎞️ Video Sequence: {meta['dimensions']['time_points']} frames")
        print(f"  └── 🧮 Z Planes: {meta['dimensions']['z_planes']} layers")
        print(f"  └── 🔬 Scale/Resolution: X: {meta['scalebar']['microns_per_pixel_x']} {meta['scalebar']['unit']}/px | Y: {meta['scalebar']['microns_per_pixel_y']} {meta['scalebar']['unit']}/px")
        print(f"  └── 🎨 Channels:")
        for ch in meta['channels']:
            print(f"      • Index {ch['channel_index']}: {ch['name']} (Hex Color: {ch['hex_color']})")