# parser.py
import re
import logging
import chardet

def parse_txt_file(file_path):
    data = {}
    camera_found = False
    sections = {
        "System": {},
        "Processor": {},
        "Memory": {},
        "Monitor": {},
        "Drive": {},
        "Network": {},
        "Audio": {},
        "Battery": {},
        "Video": {},
        "BIOS": {}
    }
    pending_subsection_value = None
    keyname_fallback = {"Video Chipset": [], "Maximum Link Speed": []}
    
    # List of common VID/PID for laptop webcams
    webcam_vid_pid = [
        "VID_0BDA&PID_5520", "VID_1BCF&PID_2B95", "VID_04F2&PID_B5A7", "VID_322E&PID_2025",
        "VID_5986&PID_118A", "VID_0C45&PID_60B0", "VID_04F2&PID_B61E", "VID_0BDA&PID_5852",
        "VID_0BDA&PID_5846", "VID_0C45&PID_6513", "VID_0408&PID_5489", "VID_04F2&PID_B578",
        "VID_0BDA&PID_57F8", "VID_0C45&PID_6366", "VID_13D3&PID_56A2", "VID_058F&PID_5608",
        "VID_0408&PID_A031", "VID_0AC8&PID_307B", "VID_0553&PID_0100", "VID_0C45&PID_62C0",
        "VID_1BCF&PID_2C99", "VID_0BDA&PID_58F4", "VID_04F2&PID_B6BF", "VID_0C45&PID_6A06",
        "VID_5986&PID_9102", "VID_0BDA&PID_58F0", "VID_0C45&PID_64AB", "VID_322E&PID_2501",
        "VID_0BDA&PID_58F2", "VID_OC45&PID_63F9", "VID_OC45&PID_63F8", "VID_OC45&PID_63E9",
        "VID_13D3&PID_5671", "VID_13D3&PID_5675", "VID_13D3&PID_5657", "VID_13D3&PID_5659",
        "VID_13D3&PID_5661", "VID_13D3&PID_5701", "VID_13D3&PID_5702", "VID_13D3&PID_5659",
        "VID_0BDA&PID_5520", "VID_1BCF&PID_2B95", "VID_04F2&PID_B5A7", "VID_322E&PID_2025",
        "VID_13D3&PID_5666", "VID_13D3&PID_5671", "VID_13D3&PID_5675", "VID_13D3&PID_5657",
        "VID_13D3&PID_5659", "VID_13D3&PID_5661", "VID_13D3&PID_5701", "VID_13D3&PID_5702",
        "VID_04F2&PID_0113", "VID_04F2&PID_100D", "VID_04F2&PID_100F", "VID_0C45&PID_63E9",
        "VID_0C45&PID_63F8", "VID_0C45&PID_63F9", "VID_046D&PID_0825", "VID_413C&PID_81E0"
    ]

    try:
        with open(file_path, 'rb') as file:
            raw_data = file.read()
            result = chardet.detect(raw_data)
            encoding = result['encoding']
        with open(file_path, 'r', encoding=encoding, errors='replace') as file:
            lines = file.readlines()
    except Exception as e:
        logging.error(f"Failed to read file {file_path}: {str(e)}")
        return data, camera_found, keyname_fallback

    logging.debug(f"Processing {len(lines)} lines from {file_path} with encoding {encoding}")

    for i, line in enumerate(lines):
        line = re.sub(r'\t+', ' ', line)
        line = re.sub(r'\s+', ' ', line).strip()
        if not line:
            logging.debug(f"Line {i+1}: Skipped (empty)")
            continue

        # Check for camera in line or VID/PID
        if 'camera' in line.lower() or any(vid_pid in line for vid_pid in webcam_vid_pid):
            camera_key = 'Camera' if 'camera' in line.lower() else 'Webcam VID/PID'
            camera_value = line if 'camera' in line.lower() else next((vid_pid for vid_pid in webcam_vid_pid if vid_pid in line), 'Unknown')
            camera_found = True
            logging.debug(f"Line {i+1}: {camera_key} detected: {camera_value}")

        if line.startswith('[') and line.endswith(']') and line != '[General Information]':
            pending_subsection_value = line.strip('[]')
            logging.debug(f"Line {i+1}: Found subsection: {pending_subsection_value}")
            continue
        elif pending_subsection_value:
            if 'Supported Video Modes' in pending_subsection_value:
                resolution_match = re.match(r'(\d+\s*x\s*\d+)', line)
                if resolution_match:
                    resolution = resolution_match.group(1).strip()
                    sections["Monitor"]["Supported Video Modes"] = resolution
                    logging.debug(f"Line {i+1}: Stored subsection value: {pending_subsection_value} = {resolution}")
                else:
                    logging.debug(f"Line {i+1}: Could not extract resolution from: {line}")
            elif line:
                keyname_fallback[pending_subsection_value] = line
                logging.debug(f"Line {i+1}: Stored subsection value: {pending_subsection_value} = {line}")
            pending_subsection_value = None
            continue
        if ':' in line:
            try:
                key, value = map(str.strip, line.split(':', 1))
                if not key or not value:
                    logging.debug(f"Line {i+1}: Skipped (invalid key-value pair: {line})")
                    continue

                logging.debug(f"Line {i+1}: Found key-value pair: {key} = {value}")

                if key == "Computer Brand Name":
                    sections["System"][key] = value
                elif key == "Product Serial Number":
                    sections["System"][key] = value
                elif key == "SKU Number":
                    sections["System"][key] = value
                elif key == "CPU Brand Name":
                    sections["Processor"][key] = value
                elif key == "Total Memory Size":
                    sections["Memory"][key] = value
                elif key == "Memory Speed":
                    sections["Memory"][key] = value
                elif key == "Drive Model":
                    # Extract drive size from value
                    size_match = re.search(r'(\d+)\s*(GB|TB)', value, re.IGNORECASE)
                    if size_match:
                        size = int(size_match.group(1))
                        unit = size_match.group(2).upper()
                        if unit == "TB":
                            size *= 1000  # Convert TB to GB
                        if size >= 128:  # Only include drives >= 128GB
                            sections["Drive"][key] = value
                        else:
                            logging.debug(f"Line {i+1}: Skipped drive size {size}GB (below 128GB)")
                    else:
                        logging.debug(f"Line {i+1}: No valid drive size found in: {value}")
                elif key == "Network Card" and "wi-fi" in value.lower() and "ethernet" not in value.lower():
                    sections["Network"][key] = value
                elif key == "Monitor Name (Manuf)":
                    sections["Monitor"][key] = value
                elif key == "Audio Adapter":
                    sections["Audio"][key] = value
                elif key == "Wear Level":
                    sections["Battery"][key] = value
                elif key in ["BIOS Version", "UEFI Boot"]:
                    sections["BIOS"][key] = value

                if key == "Video Chipset" and "Codename" not in key:
                    keyname_fallback["Video Chipset"].append(value)
                    logging.debug(f"Line {i+1}: Stored in fallback: {key} = {value}")
                elif key == "Operating System":
                    keyname_fallback[key] = value
                    logging.debug(f"Line {i+1}: Stored in fallback: {key} = {value}")
                elif key == "Maximum Link Speed" and "wi-fi" in line.lower():
                    speed_match = re.search(r'\d+\s*Mbps', value)
                    if speed_match:
                        keyname_fallback["Maximum Link Speed"].append(speed_match.group(0))
                        logging.debug(f"Line {i+1}: Stored in fallback: {key} = {speed_match.group(0)}")
            except ValueError:
                logging.debug(f"Line {i+1}: Skipped (malformed line: {line})")
                continue
        else:
            logging.debug(f"Line {i+1}: Skipped (no colon: {line})")

    for section_name, section_data in sections.items():
        if section_data:
            data[section_name] = section_data

    logging.debug(f"Parsed section data: {data}")
    logging.debug(f"Parsed fallback data: {keyname_fallback}")
    return data, camera_found, keyname_fallback