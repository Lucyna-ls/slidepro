import json
import os
import re
import random

from src.app.utils import filter_jsons_by_metadata, compare_jsons_to_multiple, \
    compare_template_names, get_placeholders, generate_placeholders, fill_placeholders, apply_master_colors, \
    load_json_files_from_directory


class RecommenderBase:
    def __init__(self, input_slide, category):
        self.input_slide = input_slide
        self.category = category
        self.metaData = None
        self.top_templates = []
        self.json_placeholders = None
        self.MAX_TEMPLATES = 10

    def extract_metadata(self):
        raise NotImplementedError

    def get_category_prompt(self):
        raise NotImplementedError

    def load_json_files(self):
        """Load JSON files from directory."""
        input_jsons = load_json_files_from_directory(self.category)
        json_data_list = [item[0] for item in input_jsons]
        filenames = [item[1] for item in input_jsons]
        return json_data_list, filenames

    def filter_jsons_by_metadata(self, json_data_list, filenames):
        """Filter JSON data based on metadata."""
        json_data_list, filenames = filter_jsons_by_metadata(
            json_data_list, filenames, self.metaData
        )
        return json_data_list, filenames

    def compare_jsons(self, json_data_list, filenames):
        """Compare input_slide to multiple JSONs."""
        top_matches = compare_jsons_to_multiple(
            self.input_slide, json_data_list, filenames, self.MAX_TEMPLATES
        )
        return top_matches

    def process_top_matches(self, top_matches):
        """Process the top matching templates."""
        for match_index, closest_filename, similarity in top_matches:
            print(
                f"Top match is JSON file: {closest_filename} with Similarity Score: {similarity:.2f}"
            )
            template_json = self.load_template(closest_filename)
            if self.is_self_recommendation(template_json):
                print("\n\t====Input is a template, skipping self-recommendation===")
                continue
            placeholders = self.extract_placeholders(template_json)
            if self.json_placeholders is None:
                self.json_placeholders = self.generate_placeholders(placeholders)

            updated_template = self.apply_master_colors(template_json)
            updated_template = self.fill_placeholders(updated_template)

            # updated_template = self.fill_placeholders(template_json)
            # updated_template = self.apply_master_colors(updated_template)
            self.top_templates.append(
                {
                    "filename": closest_filename,
                    "similarity_score": similarity,
                    "updated_template": updated_template,
                }
            )
            if len(self.top_templates) > self.MAX_TEMPLATES:
                break

    def load_template(self, filename):
        """Load the template file based on filename."""
        slide_template_match = re.search(r'(\d+)_(\d+)', filename)
        slide_number = slide_template_match.group(1) if slide_template_match else "unknown"
        template_number = slide_template_match.group(2) if slide_template_match else "unknown"
        template_folder = os.path.join("Data", self.category, "template")
        template_file = os.path.join(
            template_folder, f'template_slide{slide_number}_{template_number}.json'
        )
        if not os.path.exists(template_file):
            raise Exception(
                f"Template for slide {slide_number}_{template_number} not found"
            )
        with open(template_file, 'r') as f:
            template_json = json.load(f)
        return template_json

    def is_self_recommendation(self, template_json):
        """Check if the template is the same as the input slide."""
        return compare_template_names(self.input_slide, template_json)

    def extract_placeholders(self, template_json):
        """Extract placeholders from the template."""
        template_str = json.dumps(template_json)
        placeholders = get_placeholders(template_str)
        placeholders = list(set(placeholders))
        print("\n\t====placeholders: ", placeholders)
        return placeholders

    def generate_placeholders(self, placeholders):
        """Generate placeholder values from the input slide."""
        input_str = json.dumps(self.input_slide)
        category_prompt = self.get_category_prompt()
        placeholders_json = generate_placeholders(input_str, placeholders, category_prompt)
        json_placeholders = json.loads(placeholders_json)
        print("\n\t====generate_placeholders: ", json_placeholders)
        return json_placeholders

    def fill_placeholders(self, template_json):
        """Fill placeholders in the template with values."""
        template_str = json.dumps(template_json)
        updated_template_str = fill_placeholders(template_str, self.json_placeholders)
        updated_template_str = re.sub(r'<<[^<>]+>>', '', updated_template_str)
        updated_template_str = updated_template_str.replace('<<', '').replace('>>', '')
        updated_template = json.loads(updated_template_str)
        return updated_template

    def apply_master_colors(self, template_json):
        """Apply master colors from the input slide to the template."""
        updated_template = apply_master_colors(self.input_slide, template_json)
        return updated_template

    def recommend(self):
        """Run the recommendation pipeline."""
        self.extract_metadata()
        json_data_list, filenames = self.load_json_files()
        json_data_list, filenames = self.filter_jsons_by_metadata(
            json_data_list, filenames
        )
        if not json_data_list:
            raise Exception("No matching templates found based on metadata")
        top_matches = self.compare_jsons(json_data_list, filenames)
        self.process_top_matches(top_matches)
        # randomize the top_templates
        random.shuffle(self.top_templates)


        if not self.top_templates:
            raise Exception("No matching templates found in the top results")
        return self.top_templates
