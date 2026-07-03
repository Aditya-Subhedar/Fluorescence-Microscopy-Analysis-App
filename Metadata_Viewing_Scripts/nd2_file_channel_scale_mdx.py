import os
import nd2

def extract_nd2_images_metadata(file_path):
    if not os.path.exists(file_path):
        print(f"Error: Could not find the file at '{file_path}'")
        return {}

    try:
        with nd2.ND2File(file_path) as reader:
            # 1. Base Dimensions
            width = reader.sizes.get('X', None)
            height = reader.sizes.get('Y', None)
            z_planes = reader.sizes.get('Z', 1)
            time_points = reader.sizes.get('T', 1)
            num_images = reader.sizes.get('P', 1)
            
            # 2. EXTRACT MICROSCOPIC CALIBRATION (Deep Target Fallbacks)
            microns_x = None
            microns_y = None
            
            # Target A: Pull calibration attached directly to Frame 0 (Common in Reconstructed/SIM pipelines)
            try:
                # Read structural frame-level metadata parameters without loading actual pixel pixel chunks
                frame_meta = reader.frame_metadata(0)
                if hasattr(frame_meta, 'channels') and frame_meta.channels:
                    # Check first frame channel parameters mapping
                    f_ch = frame_meta.channels[0]
                    if hasattr(f_ch, 'volume') and f_ch.volume:
                        microns_x = round(f_ch.volume.axesCalibration[0], 5)
                        microns_y = round(f_ch.volume.axesCalibration[1], 5)
            except Exception:
                pass

            # Target B: Native structured file attributes
            if microns_x is None:
                try:
                    if hasattr(reader, 'attributes') and reader.attributes:
                        attrs = reader.attributes
                        if hasattr(attrs, 'pixelToMicron') and attrs.pixelToMicron:
                            microns_x = round(attrs.pixelToMicron, 5)
                            microns_y = round(attrs.pixelToMicron, 5)
                except Exception:
                    pass

            # Target C: Text Data Hard-Scraping from the Description Summary block
            if microns_x is None and hasattr(reader, 'text_info') and reader.text_info:
                txt_data = reader.text_info.get('description', '')
                for line in txt_data.split('\n'):
                    # Match lines containing standard "Calibration:", "Scale:", or "um/px"
                    if any(k in line.lower() for k in ['scale', 'calib', 'μm', 'um']):
                        words = line.replace(':', ' ').replace('=', ' ').split()
                        for word in words:
                            try:
                                # Target strings that resemble pure decimal scaling integers
                                if '.' in word and word.replace('.', '', 1).replace('-', '').isdigit():
                                    val = abs(float(word))
                                    if 0.0001 < val < 100.0:  # Validation sanity range for objective scales
                                        microns_x = round(val, 5)
                                        microns_y = round(val, 5)
                                        break
                            except ValueError:
                                pass
                        if microns_x is not None:
                            break

            # 3. Parse Channels & Hex Colors
            channel_list = []
            if hasattr(reader, 'metadata') and reader.metadata and reader.metadata.channels:
                for c_idx, ch_info in enumerate(reader.metadata.channels):
                    c_name = ch_info.channel.name if hasattr(ch_info, 'channel') else f"Channel {c_idx + 1}"
                    
                    hex_color = "Unknown"
                    if hasattr(ch_info, 'channel') and hasattr(ch_info.channel, 'colorRGB') and ch_info.channel.colorRGB is not None:
                        hex_color = f"#{ch_info.channel.colorRGB & 0xFFFFFF:06X}"
                    
                    # Wavelength color mapping backup if metadata flag is missing
                    if hex_color == "Unknown":
                        name_lower = c_name.lower()
                        if any(k in name_lower for k in ['488', 'gfp', 'fitc', 'alexa488']):
                            hex_color = "#00FF00"  # Green
                        elif any(k in name_lower for k in ['dapi', '405', 'hoechst']):
                            hex_color = "#0000FF"  # Blue
                        elif any(k in name_lower for k in ['561', '555', 'tritc', 'cy3']):
                            hex_color = "#FF0000"  # Red
                        elif any(k in name_lower for k in ['647', 'cy5']):
                            hex_color = "#800080"  # Purple
                    
                    channel_list.append({
                        "channel_index": c_idx,
                        "name": c_name,
                        "hex_color": hex_color
                    })
            else:
                num_channels = reader.sizes.get('C', 1)
                for c_idx in range(num_channels):
                    channel_list.append({
                        "channel_index": c_idx,
                        "name": f"Channel {c_idx + 1}",
                        "hex_color": "Unknown"
                    })

            # 4. Generate structured manifest dict mapping
            nd2_metadata_manifest = {}
            for img_idx in range(num_images):
                img_name = f"Series_{img_idx + 1}" if num_images > 1 else os.path.basename(file_path)
                
                nd2_metadata_manifest[img_name] = {
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
                
        return nd2_metadata_manifest
    
    except Exception as e:
        print(f"An error occurred while parsing ND2 metadata: {e}")
        return {}


# Executing against target path
file_path = r"IHC input images\.nd2\C3_N01Ato5_nucleo-trans_div3_mSG_b2s_200ms_30%_Reconstructed.nd2"
print(f"Analyzing ND2 File: {file_path}\n" + "="*50)

all_images_meta = extract_nd2_images_metadata(file_path)

# Print out console diagnostics block
for img_name, meta in all_images_meta.items():
    print(f"\n📸 SERIES NAME: '{img_name}' (Index: {meta['image_index']})")
    print(f"  └── 📏 Size: {meta['dimensions']['width_pixels']}x{meta['dimensions']['height_pixels']} px")
    print(f"  └── 🔬 Scale/Resolution: X: {meta['scalebar']['microns_per_pixel_x']} {meta['scalebar']['unit']}/px | Y: {meta['scalebar']['microns_per_pixel_y']} {meta['scalebar']['unit']}/px")
    print(f"  └── 🎨 Channels:")
    for ch in meta['channels']:
         print(f"      • Index {ch['channel_index']}: {ch['name']} (Hex Color: {ch['hex_color']})")
