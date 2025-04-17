import os
import logging
from datetime import datetime

def remove_xml_stylesheet(directory):
    modified_count = 0
    target_string1 = "<?xml-stylesheet type='text/xsl' href='../../Ressourcen/Stylesheet.xsl'?>"
    target_string2 = "<?xml-stylesheet type=\"text/xsl\" href=\"../../Ressourcen/Stylesheet.xsl\"?>"
    for root, _, files in os.walk(directory):
        for file in files:
            if file.endswith('.xml'):
                file_path = os.path.join(root, file)
                try:
                    with open(file_path, 'r', encoding='utf-8') as f:
                        content = f.read()
                except Exception as e:
                    logging.error("Error reading file: %s, error: %s", file_path, e)
                    continue
                if target_string1 in content or target_string2 in content:
                    updated_content = content.replace(target_string1, "").replace(target_string2, "")
                    try:
                        with open(file_path, 'w', encoding='utf-8') as f:
                            f.write(updated_content)
                        logging.info("Modified: %s", file_path)
                        modified_count += 1
                    except Exception as e:
                        logging.error("Error writing file: %s, error: %s", file_path, e)
    return modified_count

if __name__ == "__main__":
    # Create a dynamic log filename with datetime
    log_filename = f'logs/log_{datetime.now().strftime("%Y%m%d_%H%M%S")}.txt'
    logging.basicConfig(filename=log_filename, level=logging.INFO)
    data_directory = "/Users/stazh/Documents/zszh-data-recovery/zszh-data/data/RRB"
    logging.info("Processing directory: %s", os.path.abspath(data_directory))
    modified_count = remove_xml_stylesheet(data_directory)
    logging.info("Total modified files: %d", modified_count)