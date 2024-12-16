CATEGORY_PROMPT = """
Your role is of a expert PPT analyzer. You are provided with the content of a PowerPoint presentation.
Your job is to carefully analyze the content and provide the Category name Only.
Category can only be one of the following: "Agenda", "Next Steps", "Quotes Testimonials", "Timeline", "Charts", "Table". 

Input Slide: {input_slide}

You MUST return the response in the following JSON format:

{{"category": "Category Name"}}

Response:
"""

PLACEHOLDERS_GENERATION_PROMPT_NEXTSTEPS = """

You are a smart PowerPoint Analyzer. Your job is to extract the provided placeholders from the given PowerPoint slides. Ensure that you accurately match each placeholder to the relevant content from the slide and return the final extracted keys in the form of JSON.

    - Ignore numbers if any appear in the slide. Do not count them as values.
    - Maintain the order of placeholders as listed. For instance, Step 1 should be counted first, then Step 2, and so on.
    - Each placeholder (e.g., <<Step1>>, <<Step1Details>>) corresponds to specific content from the slide.
    - Content marked as <<ADDITIONAL_INFO>> is typically found at the top or bottom of the slide. This information should not be confused with the main steps.
    - Ensure NO steps or details are REPEATED in your final output.
    - If there is a placeholder for a step that contains an ellipsis ("..."), the last step in your JSON output should be the ellipsis (i.e., "…").
    - The <<Title>> usually have the keyword "Next Steps". So don't include it in the steps.
    - Do Not Include "\n" in the extracted values.
    - Keep the ORDER of the STEP as they appear in the slide. You can use LEFT Coordinates to identify the order of the steps.
    
Input Slide : {input_slide}

Placeholders to extract : {placeholders}


The final response MUST be in the following JSON format:

    {placeholder_format}

Replace [PLACEHOLDER_VALUE] with the extracted value.


Placeholders JSON:
"""


PLACEHOLDERS_GENERATION_PROMPT_AGENDA = """
You are a smart Powerpoint Analyzer. Your job is to extract out provided elements from the given powerpoint slides. Give the final extracted Keys in form of JSON.

Ignore the numbers if any in the slide (DONOT count them as values)

Input Slide : {input_slide}

Placeholders to extract : {placeholders}


The final response MUST be in the following JSON format:

    {placeholder_format}

Replace [PLACEHOLDER_VALUE] with the extracted value.

Think step by step, understand the content of the slide and then generate the placeholders.  
I will penalize you 100$ for every missing or wrong placeholder!.

Placeholders JSON:
"""

METADATA_EXTRACTION_PROMPT = """
Your role is of a smart PPT Analyzer.

You are provided with the JSON representation for an Agenda PPT slide.
The Agenda slide contains a title (Agenda, or any other title keywords) and list of points to be discussed in the AGENDA.

Your JOB is to extract out the count of Agenda points  discussed in the following Agenda slide

Ignore the NUMBERS if any in the slide (DONOT count them as points).

The MAXIMUM count value CAN BE 8.
The MINIMUM count value CAN BE 1. If the count is less than 1, then the count will be 1.

The output MUST be the following JSON:
{{"count" : NUM}}


Input Slide: `{ppt}`

Think step by step, understand the content of the slide and generate the COUNT of the agenda points. I will penalize you 100$ IF you miss any agenda point or add an extra one!

Output:
"""



METADATA_EXTRACTION_PROMPT_NEXTSTEPS = """
Your role is that of a smart PPT Analyzer.
You are provided with the JSON representation of a PowerPoint slide based on the category "NEXT STEPS".
The slide typically contains a title like "NEXT STEPS" or similar keywords, followed by a list of descriptive points for the next steps. There might be additional details for each step as well in the slide.

Your goal is to :

- Extract out the count of STEPS mentioned in the PPT slide ONLY.
- Understand the content provided and provide count for the number of steps described in the slide.
- DONOT include additional elements like Title, additional Info in the count.
- The count MUST be the number of unique steps discussed in the slide.
- Ignore any numbers or numerical points (they should not count as discussion points).
- ONLY include the relevant next steps NOT other text content.
- The minimum COUNT can be 3.
- For every step its description (additional info) CAN be written below it so based on that identify the COUNT. DONOT include these additional details in count.
- Identify the main key points (steps/headings) starting with or related to "What" or similar keywords.
- Extract the main steps from the slide by focusing on these key points and ignore details under these steps that describe them in more depth.
- Count the number of unique key points (headings) starting with "What." or similar keywords.

Think step by step, understand the content of the slide and then provide the count of discussion points about the Next Steps.
You MUST not include any irrelevant points in the count.

The output MUST be the following JSON:
{{"count" : NUM}}

Input Slide: `{ppt}`

Output:
"""


METADATA_EXTRACTION_PROMPT_QUOTES = """
Your role is of a smart PPT Analyzer. 

You are provided with the JSON representation for a Quote/Testimonials PPT slide.
The  slide contains a title and quote or testimonial. It sometimes contains the Name and Role of the person who gave the quote.
Your JOB is to extract out the count of Quote/Testimonials discussed in the following slide.


The output MUST be the following JSON:
{{"count" : NUM}}

Input Slide: `{ppt}`

Output:
"""

