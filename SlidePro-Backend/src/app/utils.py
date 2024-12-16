import json
import os
import re
from collections import Counter
from dotenv import load_dotenv
from langchain_openai import AzureChatOpenAI
from src.app.constants import CATEGORY_PROMPT


def extract_template_name(json_obj):
    for shape in json_obj['shapes']:
        if shape['Name'] == "TextBox 99999":
            return [content['Text'] for content in shape['Content']]
    return []


def compare_template_names(json1, json2):
    name1 = extract_template_name(json1)
    name2 = extract_template_name(json2)
    if name1 and name2 and name1 == name2:
        return True
    else:
        return False


def filter_dict(input_dict):
    filtered_shapes = []

    # List of keys to keep
    keys_to_keep = ['ShapeType', 'Name', 'Content', "Top", "Left"]
    # Iterate over each shape in the 'shapes' list
    for shape in input_dict.get("shapes", []):

        # Filter out shapes based on certain ShapeTypes
        if shape.get('Name') == 'TextBox 99999':
            # print("Skipping placeholder id")
            continue

        # If content length is 0, also skip
        elif shape.get('Content') is None or len(shape.get('Content')) == 0:
            # print("Skipping empty content")
            continue

        elif shape.get('ShapeType') not in ['Unknown', 'Graphic', 'EmbeddedOLEObject', 'Picture']:
            # Create a new dict with only the keys we want to keep
            filtered_shape = {key: shape[key] for key in keys_to_keep if key in shape}
            filtered_shapes.append(filtered_shape)

    # Return a new dictionary with the filtered shapes
    return {"shapes": filtered_shapes}


def get_llm(temperature: float = 0.01, timeout=None, max_tokens=None, json_mode=False):
    """
    :param temperature: Temperature of LLM model
    :param timeout: Timeout for LLM API call
    :param max_tokens: Max tokens to output
    :return: LLM Object
    """
    load_dotenv()
    if json_mode:
        llm = AzureChatOpenAI(
            azure_deployment="Slidepro",  # or your deployment
            api_version="2023-06-01-preview",  # or your api version
            temperature=temperature,
            max_tokens=max_tokens,
            timeout=timeout,
            max_retries=2,
            model_kwargs={"response_format": {"type": "json_object"}},
        )
        return llm
    else:
        llm = AzureChatOpenAI(
            azure_deployment="Slidepro",  # or your deployment
            api_version="2023-06-01-preview",  # or your api version
            temperature=temperature,
            max_tokens=max_tokens,
            timeout=timeout,
            max_retries=2,
        )
        return llm


def load_json_files_from_directory(category: str):
    json_files = []

    # Define directory paths
    input_folder = os.path.join("Data", category, "input")

    # Check if input folder exists
    if not os.path.exists(input_folder):
        raise FileNotFoundError(f"The folder '{input_folder}' does not exist.")

    # Loop through files in the input directory
    for filename in os.listdir(input_folder):
        if filename.endswith('.json'):
            file_path = os.path.join(input_folder, filename)
            try:
                # with open(file_path, 'r') as f:
                with open(file_path, 'r', encoding='utf-8') as f:
                    json_data = json.load(f)
                    json_files.append((json_data, filename))
                    # json_files.append(json_data)
            except json.JSONDecodeError as e:
                print(f"Error decoding JSON file {filename}: {e}")
            except Exception as e:
                print(f"Error reading file {filename}: {e}")

    return json_files


# Function to generate the slide category
def get_slide_category(input_slide: dict):
    llm = get_llm(temperature=0.001, json_mode=True)
    response = llm.invoke(CATEGORY_PROMPT.format(input_slide=json.dumps(input_slide)))
    res = json.loads(response.content)
    return res.get("category")


def extract_slide_metadata(input_slide: dict, category: str, prompt_str):
    filtered_slide = filter_dict(input_slide)
    llm = get_llm(temperature=0.001, json_mode=True)
    response = llm.invoke(prompt_str.format(ppt=filtered_slide))
    return json.loads(response.content)


def extract_elements(json_data):
    """
    Extracts and counts the relevant elements (text, bullet, image, rectangle, figure) from the JSON data.
    """
    elements = []
    for shape in json_data.get("shapes", []):
        shape_type = shape.get("ShapeType", "")

        # Check for text and bullets in Placeholder and TextBox shapes
        if shape_type in ["Placeholder", "TextBox"]:
            content = shape.get("Content", [])
            for item in content:
                if item.get("Text"):
                    elements.append("text")
                elif item.get("Bullet"):
                    elements.append("bullet")

    return Counter(elements)


def extract_text_count_json(data):
    texts = []

    # Check if 'shapes' key exists
    if 'shapes' in data:
        for shape in data['shapes']:
            # Check if 'Content' key exists within each shape
            if 'Content' in shape:
                for content_item in shape['Content']:
                    # Extract 'Text' if present
                    if 'Text' in content_item:
                        texts.append(content_item['Text'])
                    # Extract 'Bullet' if present
                    if 'Bullet' in content_item and content_item['Bullet'] is not None:
                        texts.append(content_item['Bullet'])

    # remove None from list
    texts = [text for text in texts if text is not None]
    final_str = ' '.join(texts)
    return len(final_str)


def calculate_content_count_similarity(json1, json2):
    """
    Calculates the similarity between two JSON objects based on the count of text content.
    """
    count1 = extract_text_count_json(json1)
    count2 = extract_text_count_json(json2)
    # Handle the case where both counts are zero
    if count1 == 0 and count2 == 0:
        return 1.0  # Both are considered identical if no text content is present

    # Avoid division by zero
    if max(count1, count2) == 0:
        return 0.0

    return min(count1, count2) / max(count1, count2)


def calculate_jaccard_similarity(counter1, counter2):
    """Calculates the Jaccard similarity between two Counters of elements."""
    intersection = sum((counter1 & counter2).values())
    union = sum((counter1 | counter2).values())

    if union == 0:
        # print("Both are empty")
        return 1.0  # If both are empty, they are perfectly similar

    return intersection / union


def compare_jsons_to_multiple(json1, json_list, filenames, top_n):
    """
    Compares a primary JSON to multiple other JSONs and finds the closest match.
    """
    counter1 = extract_elements(json1)
    similarities = []

    for i, json2 in enumerate(json_list):
        # print(" Filename : ", filenames[i])
        counter2 = extract_elements(json2)
        similarity_jaccard = calculate_jaccard_similarity(counter1, counter2)
        count_similarity = calculate_content_count_similarity(json1, json2)

        # Calculate the overall similarity score
        similarity_score = (similarity_jaccard + count_similarity) / 2
        similarities.append((i, similarity_score))

    top_matches = sorted(similarities, key=lambda x: x[1], reverse=True)[:top_n]

    # Return the top N matches as tuples (index, filename, similarity)
    return [(i, filenames[i], similarity) for i, similarity in top_matches]


def generate_placeholders(input_slide: str, placeholders: list, category_prompt: str):
    filter_slide = filter_dict(json.loads(input_slide))

    # llm = ChatOpenAI(model="gpt-4o-mini", temperature=0.01, model_kwargs={"response_format": {"type": "json_object"}})
    llm = get_llm(temperature=0.01, json_mode=True)
    template_json = {}
    for placeholder in placeholders:
        template_json[placeholder] = "[PLACEHOLDER_VALUE]"

    response = llm.invoke(category_prompt.format(input_slide=filter_slide, placeholders=placeholders,
                                                                placeholder_format=json.dumps(template_json)))
    return response.content



def fill_placeholders(template_Str, placeholders_json):
    for key, value in placeholders_json.items():
        # placeholder_key = f"<<{key}>>"
        if value is None:
            value = ""
        if value == "[PLACEHOLDER_VALUE]":
            value = "..."

        template_Str = template_Str.replace(key, value)
        # Manually remove << and >> symbols
        # template_Str = template_Str.replace("<<", "").replace(">>", "")
    return template_Str


def get_placeholders(template_string):
    # Adjust the regex to match placeholders within << >>
    placeholders = re.findall(r'<<.*?>>', template_string)
    return placeholders


def filter_jsons_by_metadata(json_list, filenames, filters):
    """
    Filters JSON objects and their corresponding filenames based on metadata filters.

    :param json_list: List of JSON objects
    :param filenames: List of filenames corresponding to each JSON
    :param filters: Dictionary of filter criteria
    :return: Tuple of filtered JSONs and filtered filenames
    """
    filtered_jsons = []
    filtered_filenames = []

    # Ensure both lists are the same length
    if len(json_list) != len(filenames):
        raise ValueError("json_list and filenames must have the same length")

    for json_obj, filename in zip(json_list, filenames):
        # Check if the JSON object contains a 'metadata' key and the filter matches
        metadata = json_obj.get('metadata', {})
        if all(metadata.get(key) == value for key, value in filters.items()):
            filtered_jsons.append(json_obj)
            filtered_filenames.append(filename)

    return filtered_jsons, filtered_filenames


def is_valid_hex_color(hex_color):
    """
    Check if the given string is a valid hex color.
    """
    if isinstance(hex_color, str) and hex_color.startswith("#") and len(hex_color) == 7:
        try:
            int(hex_color[1:], 16)  # Convert the hex to an integer to verify if it's a valid hex code
            return True
        except ValueError:
            return False
    return False


def is_grey_or_white(hex_color):
    """
    Check if the given color is a shade of grey or white.
    This function specifically checks for greys and whites, ignoring other light colors.
    """
    if not is_valid_hex_color(hex_color):
        return False

    # Convert hex to RGB
    hex_color = hex_color.lstrip('#')
    rgb = tuple(int(hex_color[i:i + 2], 16) for i in (0, 2, 4))

    # All RGB values should be close to each other to be a shade of grey or white
    r, g, b = rgb

    # Define a threshold for how close the RGB values need to be to consider it grey/white
    threshold = 15  # Allow some small variation between RGB values for greys
    if abs(r - g) <= threshold and abs(g - b) <= threshold and abs(r - b) <= threshold:
        return True

    return False


def apply_master_colors(input_slide_json, template_json):
    # Extract slide master details (if it's a list, take the first element)
    slide_master = input_slide_json.get("slidemaster", [])

    # Check if slide_master is a list and not empty
    if isinstance(slide_master, list) and slide_master:
        slide_master = slide_master[0]  # Access the first dictionary in the list

    # Extract FillColor and FontColor from slide_master
    shape_color = slide_master.get("FillColor") if isinstance(slide_master, dict) else None

    # Skip if slide master color is white
    if shape_color == "#FFFFFF":
        return template_json

    # Update Font color for all shapes with Slide Master Font color
    for shape in template_json.get("shapes", []):
        font_color = shape.get("FontColor", None)
        # Check if the current Font color is not grey/white or black
        if font_color and is_valid_hex_color(font_color) and not is_grey_or_white(font_color) and font_color != "#000000":
            shape["FontColor"] = shape_color

    # Iterate through all shapes in the template JSON
    for shape in template_json.get("shapes", []):
        shape_type = shape.get("ShapeType")
        current_fill_color = shape.get("FillColor")
        current_line_color = shape.get("LineColor")


        # Apply shape color only to rectangles, rounded rectangles, and ovals
        if shape_type in ["msoShapeRectangle", "msoShapeRoundedRectangle", "msoShapeOval", "Line", "msoShapeIsoscelesTriangle", "msoShapeDonut", "Picture_freeform"]:
            # Check if the current fill color is not a grey/white color
            if current_fill_color and is_valid_hex_color(current_fill_color) and not is_grey_or_white(
                    current_fill_color):
                shape["FillColor"] = shape_color

            # Check if the current line color is not grey/white
            if current_line_color and is_valid_hex_color(current_line_color) and not is_grey_or_white(
                    current_line_color):
                shape["LineColor"] = shape_color

    return template_json
