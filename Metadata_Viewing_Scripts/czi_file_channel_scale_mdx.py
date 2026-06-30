import os
from pylibCZIrw import czi as pyczi

def extract_czi_comprehensive_metadata(file_path):
    """
    Extracts both channel configurations and physical scaling (pixel size)
    metadata from a CZI image in a single pass.
    """
    if not os.path.exists(file_path):
        print(f"Error: Could not find the file at '{file_path}'")
        return None

    try:
        with pyczi.open_czi(file_path) as czidoc:
            metadata_dict = czidoc.metadata
            
            # --- Helper Functions for Recursive Search ---
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

            # --- 1. Channel Extraction Pipeline ---
            raw_channels = find_keys(metadata_dict, 'Channel')
            channels_info = []
            
            for ch in raw_channels:
                c_name = ch.get("@Name") or ch.get("Name") or "Unknown"
                c_id = ch.get("@Id") or ch.get("Id") or "Unknown"
                c_color = ch.get("Color") or "Unknown"
                wavelength = ch.get("EmissionWavelength") or "N/A"
                
                channels_info.append({
                    "ID": c_id,
                    "Name": c_name,
                    "Color": c_color,
                    "Wavelength": wavelength
                })

            # --- 2. Scaling / Distance Extraction Pipeline ---
            raw_distances = find_keys(metadata_dict, 'Distance')
            scaling_info = {}
            
            for dist in raw_distances:
                axis_id = dist.get("@Id") or dist.get("Id")
                val_str = dist.get("Value")
                
                if axis_id and val_str:
                    try:
                        val_meters = float(val_str)
                        val_um = val_meters * 1e6  # Convert to micrometers
                        
                        if axis_id.upper() in ['X', 'Y', 'Z']:
                            scaling_info[axis_id.upper()] = {
                                "Meters": val_meters,
                                "Micrometers": round(val_um, 4)
                            }
                    except ValueError:
                        pass
            
            # --- Consolidate Output ---
            return {
                "Channels": channels_info,
                "Scaling": scaling_info
            }
            
    except Exception as e:
        print(f"An error occurred: {e}")
        return None

# --- Main Console Execution ---
if __name__ == "__main__":
    file_path = r"IHC input images\czi\3 Channel\10_A1.czi"
    print(f"Analyzing File: {file_path}")
    print("=" * 50)
    
    metadata = extract_czi_comprehensive_metadata(file_path)
    
    if metadata:
        # Print Channel Data
        print("\n[ CHANNEL METADATA ]")
        if metadata["Channels"]:
            for i, ch in enumerate(metadata["Channels"]):
                print(f"Index {i}: {ch['Name']} (ID: {ch['ID']}, Color: {ch['Color']}, Wave: {ch['Wavelength']})")
        else:
            print("No channel information found.")
            
        # Print Scale Bar Data
        print("\n[ SCALING / PIXEL METADATA ]")
        if metadata["Scaling"]:
            for axis, sizes in metadata["Scaling"].items():
                print(f"Axis {axis}: {sizes['Micrometers']} µm per pixel (Raw: {sizes['Meters']} m)")
        else:
            print("No scaling metadata found. Image might be uncalibrated.")
            
    else:
        print("\nFailed to extract metadata.")
