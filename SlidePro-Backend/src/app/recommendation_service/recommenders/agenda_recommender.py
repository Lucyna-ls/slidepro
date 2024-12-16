from ..base_recommender import RecommenderBase
from ...constants import METADATA_EXTRACTION_PROMPT, PLACEHOLDERS_GENERATION_PROMPT_AGENDA
from ...utils import extract_slide_metadata, filter_dict, get_llm


class AgendaRecommender(RecommenderBase):
    category_name = "Agenda"

    def __init__(self, input_slide, category):
        super().__init__(input_slide, category)
        self.MAX_TEMPLATES = 30

    def extract_metadata(self):
        """Extract metadata for Agenda category."""
        self.metaData = extract_slide_metadata(self.input_slide, self.category, METADATA_EXTRACTION_PROMPT)
        print("MetaData : ", self.metaData)

    def get_category_prompt(self):
        """Return the placeholder generation prompt specific to 'Next Steps' category."""
        return PLACEHOLDERS_GENERATION_PROMPT_AGENDA

