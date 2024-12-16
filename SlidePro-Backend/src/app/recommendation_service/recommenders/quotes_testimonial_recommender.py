from ..base_recommender import RecommenderBase
from ...constants import METADATA_EXTRACTION_PROMPT_QUOTES
from ...utils import extract_slide_metadata


class QuotesTestimonialRecommender(RecommenderBase):
    category_name = "Quotes Testimonials"

    def __init__(self, input_slide, category):
        super().__init__(input_slide, category)

    def extract_metadata(self):
        """Extract metadata for Agenda category."""
        self.metaData = extract_slide_metadata(self.input_slide, self.category, METADATA_EXTRACTION_PROMPT_QUOTES)
        print("MetaData : ", self.metaData)

