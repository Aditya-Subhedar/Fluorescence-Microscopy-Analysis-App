import os
import re
import tifffile


def extract_tiff_metadata(file_path):
    """Robust metadata-driven parser for standalone, multi-layer, list, or RGB TIFFs.

    Handles flat RGB channel splitting, list-based image series, and deep text
    tag scanning.
    """
    if not os.path.exists(file_path):
        print(f"Error: Could not find file at '{file_path}'")
        return {}

    image_name = os.path.basename(file_path)

    # Core target JSON manifest layout structure
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
    manifest_core = manifest[image_name]

    try:
        with tifffile.TiffFile(file_path) as tif:
            # --- 1. DYNAMIC GEOMETRY RESOLUTION (LIST AND HYPERSTACK SAFE) ---
            series = tif.series if tif.series else None

            # Safe verification: extract standard geometry from the first page fallback
            first_page = tif.pages[0]
            width = first_page.imagewidth
            height = first_page.imagelength
            z_planes = len(tif.pages)
            time_points = 1

            # Check samples per pixel to find flat RGB image arrays
            num_channels = (
                first_page.samplesperpixel
                if hasattr(first_page, "samplesperpixel")
                else 1
            )

            # If it's a valid scientific series object (not a flat page list)
            if series is not None and not isinstance(series, list):
                try:
                    axes_str = series.axes.upper()
                    shape_arr = series.shape
                    dim_map = dict(zip(axes_str, shape_arr))

                    width = dim_map.get("X", width)
                    height = dim_map.get("Y", height)
                    z_planes = dim_map.get("Z", z_planes)
                    time_points = dim_map.get("T", time_points)
                    num_channels = dim_map.get("C", num_channels)

                    if num_channels == 1 and "S" in dim_map:
                        # Check for RGB/RGBA samples
                        if dim_map["S"] in [3, 4]:
                            num_channels = dim_map["S"]
                except AttributeError:
                    pass
            elif isinstance(series, list) and len(series) > 0:
                # If it's a list, process dimensions using the first block element safely
                try:
                    first_series = series[0]
                    axes_str = first_series.axes.upper()
                    shape_arr = first_series.shape
                    dim_map = dict(zip(axes_str, shape_arr))

                    width = dim_map.get("X", width)
                    height = dim_map.get("Y", height)
                    # If pages are split across list indexes, align metrics
                    if "C" in dim_map:
                        num_channels = dim_map["C"]
                except AttributeError:
                    pass

            manifest_core["dimensions"]["width_pixels"] = width
            manifest_core["dimensions"]["height_pixels"] = height
            manifest_core["dimensions"]["z_planes"] = z_planes
            manifest_core["dimensions"]["time_points"] = time_points

            # --- 2. DEEP SCALEBAR RESOLUTION & METADATA LOOPS ---
            pixel_size_x, pixel_size_y = None, None
            page_tags = first_page.tags

            # Method A: Process baseline hardware resolution fraction tags
            if "XResolution" in page_tags and "YResolution" in page_tags:
                try:
                    res_x_tag = page_tags["XResolution"].value
                    res_y_tag = page_tags["YResolution"].value
                    unit_type = (
                        page_tags["ResolutionUnit"].value
                        if "ResolutionUnit" in page_tags
                        else 2
                    )

                    if isinstance(res_x_tag, (tuple, list)) and len(res_x_tag) == 2:
                        val_x = (
                            float(res_x_tag[0]) / float(res_x_tag[1])
                            if res_x_tag[1] != 0
                            else 0.0
                        )
                    else:
                        val_x = float(res_x_tag) if res_x_tag else 0.0

                    if isinstance(res_y_tag, (tuple, list)) and len(res_y_tag) == 2:
                        val_y = (
                            float(res_y_tag[0]) / float(res_y_tag[1])
                            if res_y_tag[1] != 0
                            else 0.0
                        )
                    else:
                        val_y = float(res_y_tag) if res_y_tag else 0.0

                    if val_x > 0 and val_y > 0:
                        if val_x < 50:
                            pixel_size_x = 1.0 / val_x
                            pixel_size_y = 1.0 / val_y
                        else:
                            if unit_type == 3:  # Centimeters base unit
                                pixel_size_x = 10000.0 / val_x
                                pixel_size_y = 10000.0 / val_y
                            elif unit_type == 2:  # Inches base unit
                                pixel_size_x = 25400.0 / val_x
                                pixel_size_y = 25400.0 / val_y
                except Exception:
                    pass

            # Method B: Scan raw OME-XML strings
            if (pixel_size_x is None or pixel_size_y is None) and hasattr(
                tif, "ome_metadata"
            ):
                if tif.ome_metadata:
                    match_x = re.search(
                        r'PhysicalSizeX="([0-9.]+)"',
                        tif.ome_metadata,
                        re.IGNORECASE,
                    )
                    match_y = re.search(
                        r'PhysicalSizeY="([0-9.]+)"',
                        tif.ome_metadata,
                        re.IGNORECASE,
                    )
                    if match_x:
                        pixel_size_x = float(match_x.group(1))
                    if match_y:
                        pixel_size_y = float(match_y.group(1))

            # Method C: Deep text scraping fallback (Targeting raw metadata descriptions)
            text_dump = ""
            for tag_name in ["ImageDescription", "Software", "Artist"]:
                if tag_name in page_tags:
                    text_dump += "\n" + str(page_tags[tag_name].value)

            if pixel_size_x is None or pixel_size_y is None:
                match_desc_x = re.search(
                    r"(?:pixel_size_x|physical_size_x|spacing_x|scale_x|width_pixel_size|spacing)\s*[:=]\s*([0-9.]+)",
                    text_dump,
                    re.IGNORECASE,
                ) or re.search(
                    r"scale\s*=\s*([0-9.]+)\s*(?:µm|um)",
                    text_dump,
                    re.IGNORECASE,
                )
                match_desc_y = re.search(
                    r"(?:pixel_size_y|physical_size_y|spacing_y|scale_y|height_pixel_size)\s*[:=]\s*([0-9.]+)",
                    text_dump,
                    re.IGNORECASE,
                )

                if match_desc_x:
                    pixel_size_x = float(match_desc_x.group(1))
                    pixel_size_y = (
                        float(match_desc_y.group(1))
                        if match_desc_y
                        else pixel_size_x
                    )

            manifest_core["scalebar"]["microns_per_pixel_x"] = (
                round(pixel_size_x, 5) if pixel_size_x else None
            )
            manifest_core["scalebar"]["microns_per_pixel_y"] = (
                round(pixel_size_y, 5) if pixel_size_y else None
            )

            # --- 3. CHANNEL MANIFEST BUILDER ---
            default_hex_colors = [
                "#00FF00",
                "#FF0000",
                "#0000FF",
                "#00FFFF",
                "#FF00FF",
                "#FFFF00",
            ]
            channel_names = []

            if hasattr(tif, "ome_metadata") and tif.ome_metadata:
                channel_names = re.findall(
                    r'<Channel[^>]*Name="([^"]+)"',
                    tif.ome_metadata,
                    re.IGNORECASE,
                )

            # Assign standard sequential naming configurations for RGB plane spaces
            if not channel_names and num_channels in [3, 4]:
                channel_names = [
                    "Red Channel",
                    "Green Channel",
                    "Blue Channel",
                    "Alpha Channel",
                ]
                default_hex_colors = ["#FF0000", "#00FF00", "#0000FF", "#FFFFFF"]

            for c_idx in range(num_channels):
                ch_name = (
                    channel_names[c_idx]
                    if c_idx < len(channel_names)
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
        print(f"Failed to process image container tags: {e}")

    return manifest


if __name__ == "__main__":
    # Correct relative target file path
    file_path = r"  \file_path\... "

    all_images_meta = extract_tiff_metadata(file_path)
    
    # ---------------------------------------------------------
    # FORMATTED TERMINAL OUTPUT
    # ---------------------------------------------------------
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