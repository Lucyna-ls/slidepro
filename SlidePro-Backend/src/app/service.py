import os
import json
import time
from fastapi import HTTPException
from src.app.recommendation_service import fetch_recommender
from src.app.utils import get_slide_category


def upload_json_service(input_json: dict, template_json: dict, category: str, slide_no: str, template_no: str):
    # Define directory paths
    input_folder = os.path.join("Data", category, 'input')
    template_folder = os.path.join("Data", category, 'template')

    # Create directories if they don't exist
    os.makedirs(input_folder, exist_ok=True)
    os.makedirs(template_folder, exist_ok=True)

    # Define file paths based on slide number
    input_file_path = os.path.join(input_folder, f'input_slide{slide_no}_{template_no}.json')
    template_file_path = os.path.join(template_folder, f'template_slide{slide_no}_{template_no}.json')

    # Save input JSON
    with open(input_file_path, 'w') as input_file:
        json.dump(input_json, input_file, indent=4, ensure_ascii=False)

    # Save template JSON
    with open(template_file_path, 'w') as template_file:
        json.dump(template_json, template_file, indent=4, ensure_ascii=False)

    return input_file_path, template_file_path


async def recommendation_service(input_slide: dict):
    start_time = time.time()
    # category = get_slide_category(input_slide)
    # print("Category:", category)
    category = "Agenda"
    recommender_class = fetch_recommender(input_slide, category)
    if recommender_class is None:
        raise HTTPException(status_code=404, detail=f"No recommender found for category: {category}")
    try:
        top_templates = recommender_class.recommend()
    except Exception as e:
        raise HTTPException(status_code=404, detail=str(e))
    print(f"Recommendation Service took {time.time() - start_time:.2f} seconds")
    return top_templates
