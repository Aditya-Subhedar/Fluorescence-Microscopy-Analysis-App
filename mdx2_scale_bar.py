from pylibCZIrw import czi as pyczi
import os

def extract_czi_scale_metadata(file_path):
    if not os.path.exists(file_path):
        print(f"Error: Could not find the file at '{file_path}'")
        return None

    try:
        with pyczi.open_czi(file_path) as czidoc:
            metadata_dict = czidoc.metadata
            
            # Recursive search specifically for 'Distance' tags
            def find_distances(data):
                found = []
                if isinstance(data, dict):
                    for key, value in data.items():
                        if key == 'Distance':
                            # Distance can be a list (X, Y, Z) or a single dict
                            if isinstance(value, list):
                                found.extend(value)
                            else:
                                found.append(value)
                        else:
                            found.extend(find_distances(value))
                elif isinstance(data, list):
                    for item in data:
                        found.extend(find_distances(item))
                return found
            
            raw_distances = find_distances(metadata_dict)
            scaling_info = {}
            
            for dist in raw_distances:
                # The axis is usually stored as an 'Id' (X, Y, or Z)
                axis_id = dist.get("@Id") or dist.get("Id")
                # The physical size is stored as 'Value'
                val_str = dist.get("Value")
                
                if axis_id and val_str:
                    try:
                        val_meters = float(val_str)
                        # Convert meters to micrometers (um) for standard microscopy usage
                        val_um = val_meters * 1e6
                        
                        # Only keep X, Y, and Z axes to filter out any junk data
                        if axis_id.upper() in ['X', 'Y', 'Z']:
                            scaling_info[axis_id.upper()] = {
                                "Meters": val_meters,
                                "Micrometers": round(val_um, 4)
                            }
                    except ValueError:
                        pass
                        
        return scaling_info
    
    except Exception as e:
        print(f"An error occurred: {e}")
        return None

# --- Example Usage ---
file_path = r"IHC input images\czi\3 Channel\10_A1.czi"
print(f"Analyzing Scaling for: {file_path}\n" + "-"*40)

scale_data = extract_czi_scale_metadata(file_path)

if scale_data:
    for axis, sizes in scale_data.items():
        print(f"Axis {axis}: {sizes['Micrometers']} µm per pixel (Raw: {sizes['Meters']} m)")
else:
    print("No scaling metadata found. The image might be uncalibrated.")