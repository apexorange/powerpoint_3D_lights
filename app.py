import zipfile
import os
from lxml import etree


def unzip_pptx(file_path, extract_path):
    with zipfile.ZipFile(file_path, 'r') as zip_ref:
        zip_ref.extractall(extract_path)


def list_elements(root, element_tag):
    namespaces = root.nsmap
    namespaces[
        'am3d'] = 'http://schemas.microsoft.com/office/drawing/2017/model3d'  # Ensure 'am3d' is in the namespace map
    elements = root.findall(f'.//{element_tag}', namespaces)
    return elements


def update_rgb_values(root, element_tag, position, r, g, b):
    elements = list_elements(root, element_tag)
    if 0 < position <= len(elements):
        elem = elements[position - 1]  # Adjust for zero-based index
        scrgbClr = elem.find('.//a:scrgbClr', namespaces=root.nsmap)
        if scrgbClr is not None:
            scrgbClr.set('r', str(int(r * 1000)))
            scrgbClr.set('g', str(int(g * 1000)))
            scrgbClr.set('b', str(int(b * 1000)))
            print(f"Updated {element_tag} at position {position} to r={r}, g={g}, b={b}")
        else:
            print(f"a:scrgbClr not found in {element_tag} at position {position}")
    else:
        print(f"{element_tag} at position {position} not found")


def update_intensity(root, element_tag, position, percentage):
    elements = list_elements(root, element_tag)
    if 0 < position <= len(elements):
        elem = elements[position - 1]  # Adjust for zero-based index
        namespaces = root.nsmap
        namespaces['am3d'] = 'http://schemas.microsoft.com/office/drawing/2017/model3d'
        intensity = elem.find('.//am3d:intensity', namespaces=namespaces)
        if intensity is not None:
            # Convert percentage to n and d values
            n = int(percentage * 1000000)
            d = 1000000
            intensity.set('n', str(n))
            intensity.set('d', str(d))
            print(f"Updated {element_tag} at position {position} to {percentage}% brightness")
        else:
            print(f"am3d:intensity not found in {element_tag} at position {position}")
    else:
        print(f"{element_tag} at position {position} not found")


def update_ambient_light_rgb(root, position, r, g, b):
    update_rgb_values(root, 'am3d:ambientLight', position, r, g, b)


def update_ambient_light_intensity(root, position, percentage):
    elements = list_elements(root, 'am3d:ambientLight')
    if 0 < position <= len(elements):
        elem = elements[position - 1]  # Adjust for zero-based index
        namespaces = root.nsmap
        namespaces['am3d'] = 'http://schemas.microsoft.com/office/drawing/2017/model3d'
        illuminance = elem.find('.//am3d:illuminance', namespaces=namespaces)
        if illuminance is not None:
            # Convert percentage to n and d values
            n = int(percentage * 1000000)
            d = 1000000
            illuminance.set('n', str(n))
            illuminance.set('d', str(d))
            print(f"Updated am3d:ambientLight at position {position} to {percentage}% brightness")
        else:
            print(f"am3d:illuminance not found in am3d:ambientLight at position {position}")
    else:
        print(f"am3d:ambientLight at position {position} not found")


def save_xml(file_path, root):
    tree = etree.ElementTree(root)
    tree.write(file_path, pretty_print=True, xml_declaration=True, encoding='UTF-8')


def repack_pptx(extract_path: str, output_path: str):
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
        for folder_name, subfolders, filenames in os.walk(extract_path):
            for filename in filenames:
                file_path = os.path.join(folder_name, filename)
                zip_ref.write(file_path, os.path.relpath(file_path, extract_path))


# Paths
pptx_path = 'demo.pptx'
extract_path = 'extracted/'
slide_xml_path = os.path.join(extract_path, 'ppt/slides/slide1.xml')
new_pptx_path = 'demo_edited.pptx'

# Process
unzip_pptx(pptx_path, extract_path)

# Parse the XML and get the root
parser = etree.XMLParser(remove_blank_text=True)
tree = etree.parse(slide_xml_path, parser)
root = tree.getroot()

# List and print elements for debugging
elements_ambient_light = list_elements(root, 'am3d:ambientLight')
elements_pt_light = list_elements(root, 'am3d:ptLight')

print(f"Found {len(elements_ambient_light)} am3d:ambientLight elements:")
for i, elem in enumerate(elements_ambient_light, start=1):
    print(f"{i}: {etree.tostring(elem, pretty_print=True).decode('utf-8')}")

print(f"Found {len(elements_pt_light)} am3d:ptLight elements:")
for i, elem in enumerate(elements_pt_light, start=1):
    print(f"{i}: {etree.tostring(elem, pretty_print=True).decode('utf-8')}")

# Update RGB values and intensity for point lights
update_rgb_values(root, 'am3d:ptLight', 1, 255, 0, 0)  # Example: Update 1st ptLight to red
update_rgb_values(root, 'am3d:ptLight', 2, 80, 25, 0)  # Example: Update 2nd ptLight to green
update_rgb_values(root, 'am3d:ptLight', 3, 0, 0, 255)  # Example: Update 3rd ptLight to blue

update_intensity(root, 'am3d:ptLight', 1, 75)  # Example: Update 1st ptLight to 75% brightness
update_intensity(root, 'am3d:ptLight', 2, 50)  # Example: Update 2nd ptLight to 50% brightness
update_intensity(root, 'am3d:ptLight', 3, 25)  # Example: Update 3rd ptLight to 25% brightness

# Update RGB values and intensity for ambient light
update_ambient_light_rgb(root, 1, 5, 5, 5)  # Example: Update ambientLight to gray
update_ambient_light_intensity(root, 1, 10)  # Example: Update ambientLight to 60% brightness

# Save the modified XML
save_xml(slide_xml_path, root)

# Repack the PowerPoint file
repack_pptx(extract_path, new_pptx_path)
