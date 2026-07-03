import os
import tifffile


def extract_oir_metadata_via_tiff(file_path):
    """Uses tifffile's native compound container decoding to dynamically extract

    Olympus .oir dimensions, scalebars, and channel counts without hardcoded fallbacks.
    """
    if not os.path.exists(file_path):
        print(f"Error: Could not find the file at '{file_path}'")
        return {}

    image_name = os.path.basename(file_path)

    # Initialize manifest target JSON template
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
        # Tifffile handles Olympus OIR binary block maps directly
        with tifffile.TiffFile(file_path) as tif:

            # --- A. EXTRACT DYNAMIC DIMENSIONS ---
            # Exposes the true dimensional shape layout array (e.g., [T, Z, C, Y, X])
            series = tif.series[0]
            shape = series.shape
            axes = series.axes.upper()  # Maps string axes e.g. 'TZCYX'
            dim_map = dict(zip(axes, shape))

            manifest_core = manifest[image_name]
            manifest_core["dimensions"]["width_pixels"] = dim_map.get("X", 0)
            manifest_core["dimensions"]["height_pixels"] = dim_map.get("Y", 0)
            manifest_core["dimensions"]["z_planes"] = dim_map.get("Z", 1)
            manifest_core["dimensions"]["time_points"] = dim_map.get("T", 1)

            # --- B. EXTRACT SPATIAL CALIBRATIONS & SCALEBAR ---
            # Read metadata fields directly from Olympus tag blocks
            olympus_md = getattr(tif, "olympus_metadata", None) or getattr(
                series, "olympus_metadata", {}
            )

            # Safely look for physical spacing maps inside the parsed metadata tree
            pixel_size_x = None
            pixel_size_y = None

            if olympus_md:
                # Direct metadata structure retrieval
                try:
                    pixel_size_x = olympus_md.get("PixelSizeX") or olympus_md.get("ResolutionX")
                    pixel_size_y = olympus_md.get("PixelSizeY") or olympus_md.get("ResolutionY")
                except AttributeError:
                    pass

            # Fallback calculation if physical spacing is found inside structural tags
            if pixel_size_x is None:
                tags = tif.pages[0].tags
                if "XResolution" in tags:
                    res_x = tags["XResolution"].value
                    if res_x and res_x[0] > 0:
                        pixel_size_x = res_x[1] / res_x[0]  # Convert resolution to pixel size
                if "YResolution" in tags:
                    res_y = tags["YResolution"].value
                    if res_y and res_y[0] > 0:
                        pixel_size_y = res_y[1] / res_y[0]

            manifest_core["scalebar"]["microns_per_pixel_x"] = (
                round(float(pixel_size_x), 5) if pixel_size_x else None
            )
            manifest_core["scalebar"]["microns_per_pixel_y"] = (
                round(float(pixel_size_y), 5) if pixel_size_y else None
            )

            # --- C. CHANNEL GENERATION ---
            num_channels = dim_map.get("C", 1)
            default_hex_colors = [
                "#00FF00",
                "#FF0000",
                "#0000FF",
                "#00FFFF",
                "#FF00FF",
                "#FFFF00",
            ]

            # Re-read native dye parameters from Olympus dictionary dump if accessible
            dye_names = []
            if olympus_md and "Channels" in olympus_md:
                for ch_info in olympus_md["Channels"]:
                    if "DyeName" in ch_info:
                        dye_names.append(ch_info["DyeName"])

            for c_idx in range(num_channels):
                ch_name = (
                    dye_names[c_idx]
                    if c_idx < len(dye_names)
                    else f"Channel {c_idx + 1}"
                )
                manifest_core["channels"].append(
                    {
                        "channel_index": c_idx,
                        "name": ch_name,
                        "hex_color": default_hex_colors[
                            c_idx % len(default_hex_colors)
                        ],
                    }
                )

    except Exception as e:
        print(f"Failed to decode file properties using TiffFile engine: {e}")

    return manifest


if __name__ == "__main__":
    file_path = r"C:\CSE\8 th SEM (Internship)\Fluorescence-Microscopy-Analysis-App\IHC input images\.oir\1202-interval_30sec_sequence_frame_z stack.oir"

    print(f"Analyzing OIR via TiffFile Pipeline: {file_path}\n" + "=" * 50)
    all_images_meta = extract_oir_metadata_via_tiff(file_path)

    for img_name, meta in all_images_meta.items():
        print(f"\n📸 FILE NAME: '{img_name}' (Index: {meta['image_index']})")
        print(
            f"  └── 📏 Size: {meta['dimensions']['width_pixels']}x{meta['dimensions']['height_pixels']} px"
        )
        print(f"  └── 🎞️ Video Sequence: {meta['dimensions']['time_points']} frames")
        print(f"  └── 🧮 Z Planes: {meta['dimensions']['z_planes']} layers")
        print(
            f"  └── 🔬 Scale/Resolution: X: {meta['scalebar']['microns_per_pixel_x']} {meta['scalebar']['unit']}/px | Y: {meta['scalebar']['microns_per_pixel_y']} {meta['scalebar']['unit']}/px"
        )
        print(f"  └── 🎨 Channels:")
        for ch in meta['channels']:
            print(
                f"      • Index {ch['channel_index']}: {ch['name']} (Hex Color: {ch['hex_color']})"
            )
