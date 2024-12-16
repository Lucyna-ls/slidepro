import json
import re
from typing import List
from fastapi import FastAPI, UploadFile, File, HTTPException
from dotenv import load_dotenv
from fastapi import FastAPI, HTTPException
from .service import recommendation_service, upload_json_service

load_dotenv()
app = FastAPI()

from fastapi.middleware.cors import CORSMiddleware

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # You can restrict this to specific domains
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/", tags=["Health Check"])
async def Health_Check():
    """
     Health Check Function
    """
    return {'detail': 'Health Check'}


@app.post("/upload", tags=["Upload"])
async def upload_ppt_files(category: str, files: List[UploadFile] = File(...)):
    """
    Upload multiple JSON files and save them to category folders based on slide numbers.
    """
    # Ensure only two files are uploaded: one for input and one for template
    if len(files) != 2:
        raise HTTPException(status_code=400,
                            detail="Exactly two files are required: one for input and one for template")

    input_json = None
    template_json = None
    input_slide_no = None
    template_slide_no = None
    input_template_no = None
    template_template_no = None

    # Process each file and determine if it is input or template based on filename
    for file in files:
        #print("FILE : ", file.filename)
        content = await file.read()
        json_data = json.loads(content)
        # Extract slide number from the filename using regex
        slide_template_match = re.search(r'_(\d+)_(\d+)', file.filename.lower())

        if not slide_template_match:
            raise HTTPException(status_code=400, detail=f"Invalid file name format: {file.filename}")

        slide_no = slide_template_match.group(1)
        template_no = slide_template_match.group(2)

        if 'input' in file.filename.lower():
            input_json = json_data
            input_slide_no = slide_no
            input_template_no = template_no
        elif 'template' in file.filename.lower():
            template_json = json_data
            template_slide_no = slide_no
            template_template_no = template_no

    # Ensure both input and template JSONs are provided
    if input_json is None or template_json is None:
        raise HTTPException(status_code=400, detail="Both input and template JSON files are required")

    # Check if slide and template numbers match
    if input_slide_no != template_slide_no or input_template_no != template_template_no:
        raise HTTPException(status_code=400,
                            detail="Slide and template numbers for input and template files must match")

    # Save the JSON files into their respective folders based on slide number and template number
    input_file_path, template_file_path = upload_json_service(input_json, template_json, category, input_slide_no,
                                                          input_template_no)

    return {"message": "Files successfully uploaded and saved",
            "input_file": input_file_path,
            "template_file": template_file_path}


@app.post("/recommendation", tags=["Analyse PPT"])
async def recommendation(input_slide: dict):
    """
    Analyse the input JSON and return recommendations.
    """
    if not input_slide:
        raise HTTPException(status_code=400, detail="Input JSON is required")

    return await recommendation_service(input_slide)
