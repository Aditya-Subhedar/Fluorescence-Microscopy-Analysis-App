import os
import nd2

def extract_nd2_images_metadata(file_path):
    if not os.path.exists(file_path):
        print(f"Error: Could not find the file at '{file_path}'")
        return {}

    try:
        with nd2.ND2File(file_path) as reader:
            # 1. Unstructured metadata bypasses SDK mapping entirely
            raw_meta = reader.unstructured_metadata()
            
            # Fetch base sizes safely
            width = reader.sizes.get('X', None)
            height = reader.sizes.get('Y', None)
            z_planes = reader.sizes.get('Z', 1)
            time_points = reader.sizes.get('T', 1)
            num_images = reader.sizes.get('P', 1)
            
            # 2. Extract Microscopic Calibration (Bypassing .voxel_size)
            microns_x = None
            microns_y = None
            
            # Navigate the raw dictionary tree where Nikon saves hardware configurations
            try:
                meta_seq = raw_meta.get('ImageMetadataSeqLV|0', {})
                pic_meta = meta_seq.get('SLxPictureMetadata', {})
                
                # Check for direct calibration metrics
                if 'PixelToMicron' in pic_meta:
                    microns_x = round(float(pic_meta['PixelToMicron']), 5)
                    microns_y = round(float(pic_meta['PixelToMicron']), 5)
                elif 'dPixelToMicron' in pic_meta:
                    microns_x = round(float(pic_meta['dPixelToMicron']), 5)
                    microns_y = round(float(pic_meta['dPixelToMicron']), 5)
            except Exception:
                pass # Keep as None if scaling parameters don't exist

            # 3. Parse Channels & Colors out of the clean Nikon internal struct
            channel_list = []
            if hasattr(reader, 'metadata') and reader.metadata and reader.metadata.channels:
                for c_idx, ch_info in enumerate(reader.metadata.channels):
                    c_name = ch_info.channel.name if hasattr(ch_info, 'channel') else f"Channel {c_idx + 1}"
                    
                    hex_color = "Unknown"
                    if hasattr(ch_info, 'channel') and hasattr(ch_info.channel, 'colorRGB'):
                        int_color = ch_info.channel.colorRGB
                        if int_color is not None:
                            hex_color = f"#{int_color & 0xFFFFFF:06X}"
                    
                    channel_list.append({
                        "channel_index": c_idx,
                        "name": c_name,
                        "hex_color": hex_color
                    })
            else:
                # Fallback if channel properties block is entirely stripped
                num_channels = reader.sizes.get('C', 1)
                for c_idx in range(num_channels):
                    channel_list.append({
                        "channel_index": c_idx,
                        "name": f"Channel {c_idx + 1}",
                        "hex_color": "Unknown"
                    })

            # 4. Generate final JSON-ready manifest output
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


# Target path based on your terminal run
file_path = r"C:\Users\Asus\Downloads\C3_N01Ato5_nucleo-trans_div3_mSG_b2s_200ms_30%_Reconstructed.nd2"
print(f"Analyzing ND2 File: {file_path}\n" + "="*50)

all_images_meta = extract_nd2_images_metadata(file_path)

# Print out clean console structure layout mapping
for img_name, meta in all_images_meta.items():
    print(f"\n📸 SERIES NAME: '{img_name}' (Index: {meta['image_index']})")
    print(f"  └── 📏 Size: {meta['dimensions']['width_pixels']}x{meta['dimensions']['height_pixels']} px")
    print(f"  └── 🔬 Scale/Resolution: X: {meta['scalebar']['microns_per_pixel_x']} {meta['scalebar']['unit']}/px | Y: {meta['scalebar']['microns_per_pixel_y']} {meta['scalebar']['unit']}/px")
    print(f"  └── 🎨 Channels:")
    for ch in meta['channels']:
         print(f"      • Index {ch['channel_index']}: {ch['name']} (Hex Color: {ch['hex_color']})")
