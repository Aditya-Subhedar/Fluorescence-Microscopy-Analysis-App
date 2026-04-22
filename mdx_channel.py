from pylibCZIrw import czi as pyczi
import os

def extract_czi_channel_metadata(file_path):
    if not os.path.exists(file_path):
        print(f"Error: Could not find the file at '{file_path}'")
        return []

    try:
        with pyczi.open_czi(file_path) as czidoc:
            # It's already a dictionary! No XML parsing needed.
            metadata_dict = czidoc.metadata
            
            # Helper function to recursively hunt for 'Channel' keys
            def find_channels(data):
                found = []
                if isinstance(data, dict):
                    for key, value in data.items():
                        if key == 'Channel':
                            # It might be a list of multiple channels or just one dict
                            if isinstance(value, list):
                                found.extend(value)
                            else:
                                found.append(value)
                        else:
                            found.extend(find_channels(value))
                elif isinstance(data, list):
                    for item in data:
                        found.extend(find_channels(item))
                return found
            
            raw_channels = find_channels(metadata_dict)
            channels_info = []
            
            for ch in raw_channels:
                # xmltodict often uses '@' for XML attributes like <Channel Name="...">
                c_name = ch.get("@Name") or ch.get("Name") or "Unknown"
                c_id = ch.get("@Id") or ch.get("Id") or "Unknown"
                
                # Standard tags become standard dictionary keys
                c_color = ch.get("Color") or "Unknown"
                wavelength = ch.get("EmissionWavelength") or "N/A"
                
                channels_info.append({
                    "ID": c_id,
                    "Name": c_name,
                    "Color": c_color,
                    "Wavelength": wavelength
                })
                
        return channels_info
    
    except Exception as e:
        print(f"An error occurred: {e}")
        return []

# Run the diagnostic
file_path = r"IHC input images\czi\3 Channel\10_A1.czi"
print(f"Analyzing: {file_path}\n" + "-"*40)

metadata = extract_czi_channel_metadata(file_path)

if metadata:
    for i, ch in enumerate(metadata):
        print(f"Index {i}: {ch['Name']} (Color: {ch['Color']}, Wave: {ch['Wavelength']})")
else:
    print("No channel metadata found or file could not be read.")