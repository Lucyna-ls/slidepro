from src.app.recommendation_service.recommenders.agenda_recommender import AgendaRecommender
from src.app.recommendation_service.recommenders.next_steps_recommender import NextStepsRecommender
from src.app.recommendation_service.recommenders.quotes_testimonial_recommender import QuotesTestimonialRecommender


def fetch_recommender(input_slide, category_name):
    if category_name == "Agenda":
        return AgendaRecommender(input_slide, category_name)
    if category_name == "Next Steps":
        return NextStepsRecommender(input_slide, category_name)
    if category_name == "Quotes Testimonials":
        return QuotesTestimonialRecommender(input_slide, category_name)
    else:
        return None
