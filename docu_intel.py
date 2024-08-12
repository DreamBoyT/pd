import os  
import streamlit as st  
from pptx import Presentation  
from pptx.enum.shapes import MSO_SHAPE_TYPE  
from pptx.util import Inches  
from docx import Document  
from docx.shared import Inches as DocxInches  
from docx.oxml import OxmlElement, ns  
from io import BytesIO  
from langchain_openai import AzureChatOpenAI  
from langchain.prompts import PromptTemplate  
import re  
from PIL import Image  
import tempfile  
import pythoncom  
from win32com.client import Dispatch  
  
# Azure OpenAI API details  
azure_endpoint = 'https://chat-gpt-a1.openai.azure.com/openai/deployments/GPT-4o-Mini/chat/completions?api-version=2024-02-15-preview'  
azure_deployment_name = 'GPT-4o-Mini'  
azure_api_key = 'c09f91126e51468d88f57cb83a63ee36'  
azure_api_version = '2024-02-15-preview'  
  
# Initialize Azure OpenAI LLM  
llm = AzureChatOpenAI(  
    openai_api_key=azure_api_key,  
    api_version=azure_api_version,  
    azure_endpoint=azure_endpoint,  
    model="gpt-4o-mini",  
    azure_deployment=azure_deployment_name,  
    temperature=0.3  
)  
  
# Profanity words list  
patent_profanity_words = [  
    "absolute", "absolutely", "all", "always", "authoritative", "authoritatively", "best", "biggest", "black hat",  
    "black list", "blackhat", "blacklist", "broadest", "certain", "certainly", "chinese wall", "compel", "compelled",  
    "compelling", "compulsorily", "compulsory", "conclusive", "conclusively", "constantly", "critical", "critically",  
    "crucial", "crucially", "decisive", "decisively", "definitely", "definitive", "definitively", "determinative",  
    "each", "earliest", "easiest", "embodiment", "embodiments", "entire", "entirely", "entirety", "essential",  
    "essentially", "essentials", "every", "everything", "everywhere", "exactly", "exclusive", "exclusively", "exemplary",  
    "exhaustive", "farthest", "finest", "foremost", "forever", "fundamental", "furthest", "greatest", "highest",  
    "imperative", "imperatively", "important", "importantly", "indispensable", "indispensably", "inescapable",  
    "inescapably", "inevitable", "inevitably", "inextricable", "inextricably", "inherent", "inherently", "instrumental",  
    "instrumentally", "integral", "integrally", "intrinsic", "intrinsically", "invaluable", "invaluably", "invariably",  
    "invention", "inventions", "irreplaceable", "irreplaceably", "key", "largest", "latest", "least", "littlest", "longest",  
    "lowest", "major", "man hours", "mandate", "mandated", "mandatorily", "mandatory", "master", "maximize", "maximum",  
    "minimize", "minimum", "most", "must", "nearest", "necessarily", "necessary", "necessitate", "necessitated",  
    "necessitates", "necessity", "need", "needed", "needs", "never", "newest", "nothing", "nowhere", "obvious", "obviously",  
    "oldest", "only", "optimal", "ought", "overarching", "paramount", "perfect", "perfected", "perfectly", "perpetual",  
    "perpetually", "pivotal", "pivotally", "poorest", "preferred", "purest", "required", "requirement", "requires",  
    "requisites", "shall", "shortest", "should", "simplest", "slaves", "slightest", "smallest", "tribal knowledge",  
    "ultimate", "ultimately", "unavoidable", "unavoidably", "unique", "uniquely", "unrivalled", "urgent", "urgently",  
    "valuable", "very", "vital", "vitally", "white hat", "white list", "whitehat", "whitelist", "widest", "worst"  
]  
  
# Lists for tone, style, and conditional & tentative language  
tone_list = [  
    "Precision and Specificity",  
    "Formality",  
    "Complexity",  
    "Objective and Impersonal",  
    "Structured and Systematic"  
]  
  
style_list = [  
    "Formal and Objective",  
    "Structured and Systematic",  
    "Technical Jargon and Terminology",  
    "Detailed and Specific",  
    "Impersonal Tone",  
    "Instructional and Descriptive",  
    "Use of Figures and Flowcharts",  
    "Legal and Protective Language",  
    "Repetitive and Redundant",  
    "Examples and Clauses"  
]  
  
conditional_language_list = [  
    "may include", "in some aspects", "aspects of the present disclosure", "wireless communication networks",  
    "by way of example", "may be", "may further include", "may be used", "may occur", "may use", "may monitor",  
    "may periodically wake up", "may demodulate", "may consume", "can be performed", "may enter and remain",  
    "may correspond to", "may also include", "may be identified in response to", "may be further a function of",  
    "may be multiplied by", "may schedule", "may select", "may also double", "may further comprise",  
    "may be configured to", "may correspond to a duration value", "may correspond to a product of", "may be closer",  
    "may be significant", "may not be able", "may result", "may reduce", "may be operating in", "may further be configured to",  
    "may further process", "may be executed by", "may be received", "may avoid", "may indicate", "may be selected",  
    "may be proactive", "may perform", "may be necessary", "may be amplified", "may involve", "may require", "may be stored",  
    "may be accessed", "may be transferred", "may be implemented", "may include instructions to", "may depend upon",  
    "may communicate", "may be generated", "may be configured"  
]  
  
# Function to sanitize text by removing non-XML-compatible characters  
def sanitize_text(text):  
    return re.sub(r'[^\x09\x0A\x0D\x20-\x7F]', '', text)  
  
# Function to extract text and title from ppt slides  
def extract_text_and_title_from_ppt(ppt_file):  
    prs = Presentation(ppt_file)  
    slides_data = []  
    for slide_num, slide in enumerate(prs.slides, start=1):  
        slide_text = []  
        slide_title = None  
        for shape in slide.shapes:  
            if hasattr(shape, "text"):  
                if shape == slide.shapes.title:  
                    slide_title = shape.text  
                else:  
                    slide_text.append(shape.text)  
        slides_data.append((slide_num, slide_title, "\n".join(slide_text)))  
    return slides_data  
  
# Function to check if slide contains images, tables, or flowcharts  
def contains_relevant_elements(slide):  
    for shape in slide.shapes:  
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:  
            return True  
        if shape.shape_type == MSO_SHAPE_TYPE.TABLE:  
            return True  
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP and any(s.shape_type == MSO_SHAPE_TYPE.LINE for s in shape.shapes):  
            return True  
    return False  
  
# Function to save slide as image using win32com  
def save_slide_as_image(ppt_file, slide_num):  
    pythoncom.CoInitialize()  
    powerpoint = Dispatch("PowerPoint.Application")  
    powerpoint.Visible = 1  
  
    presentation = powerpoint.Presentations.Open(ppt_file)  
    slide = presentation.Slides(slide_num)  
    with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp:  
        img_path = tmp.name  
    slide.Export(img_path, 'PNG')  
    presentation.Close()  
    powerpoint.Quit()  
    return img_path  
  
# Function to generate explanation using Azure OpenAI  
def generate_explanation(slide_text):  
    prompt_template = PromptTemplate(  
        input_variables=["slide_text"],  
        template=f"""  
        Slide Content: {{slide_text}}  
          
        Aspects of the present disclosure may include insights extracted from the above slide content. The information should be delivered directly and engagingly. Avoid phrases like 'The slide presents,' 'discusses,' 'outlines,' or 'content.' The explanation should be formatted as paragraphs, without line breaks or bullet points, and must be semantically meaningful. Analyze the major points from the following text and create paragraphs accordingly. If there is one major point, create one paragraph. If there are two major points, create two paragraphs, and so on. Keep the paragraph precise and extremely short.  
          
        The text should adhere to the following style guidelines:  
        1. Remove all listed profanity words.  
        2. Use passive voice.  
        3. Use conditional and tentative language, such as "may include," "in some aspects," and "aspects of the present disclosure."  
        4. Replace "Million" with "1,000,000" and "Billion" with "1,000,000,000".  
        5. Maintain the following tone characteristics: {', '.join(tone_list)}.  
        6. Follow these style elements: {', '.join(style_list)}.  
        7. Use the following conditional and tentative language phrases: {', '.join(conditional_language_list)}.  
          
        It is crucial to strictly adhere to the above guidelines to ensure the highest quality and most accurate output.  
        """  
    )  
    prompt = prompt_template.format(slide_text=sanitize_text(slide_text))  
    response = llm(prompt)  
    return response.content  
  
# Function to add a horizontal rule in the Word document  
def add_horizontal_rule(doc):  
    p = doc.add_paragraph()  
    p_el = p._element  
    pPr = p_el.get_or_add_pPr()  
    pBdr = OxmlElement('w:pBdr')  
    bottom = OxmlElement('w:bottom')  
    bottom.set(ns.qn('w:val'), 'single')  
    bottom.set(ns.qn('w:sz'), '6')  
    bottom.set(ns.qn('w:space'), '1')  
    bottom.set(ns.qn('w:color'), 'auto')  
    pBdr.append(bottom)  
    pPr.append(pBdr)  
  
# Streamlit app  
st.title("PPT Insights Extractor with Azure OpenAI")  
uploaded_file = st.file_uploader("Upload a PPT file", type=["pptx"])  
  
if uploaded_file is not None:  
    with tempfile.NamedTemporaryFile(delete=False, suffix='.pptx') as tmp:  
        tmp.write(uploaded_file.read())  
        ppt_path = tmp.name  
  
    slides_data = extract_text_and_title_from_ppt(ppt_path)  
    prs = Presentation(ppt_path)  
  
    explanations = []  
    for slide_num, slide_title, slide_text in slides_data:  
        explanation = generate_explanation(slide_text)  
        explanations.append((slide_num, slide_title, explanation))  
  
    # Create a Word document  
    doc = Document()  
    for slide_num, slide_title, explanation in explanations:  
        sanitized_slide_title = sanitize_text(slide_title if slide_title else "Untitled Slide")  
        sanitized_explanation = sanitize_text(explanation)  
  
        # Add slide number  
        doc.add_heading(f"Slide {slide_num}", level=1)  
        doc.add_heading(sanitized_slide_title, level=2)  
        doc.add_paragraph(sanitized_explanation)  
  
    # Add a horizontal rule  
    add_horizontal_rule(doc)  
  
    # Append slide images after the horizontal rule, only for relevant slides  
    for slide_num, _, _ in explanations:  
        slide = prs.slides[slide_num - 1]  
        if contains_relevant_elements(slide):  
            img_path = save_slide_as_image(ppt_path, slide_num)  
            doc.add_heading(f"Slide {slide_num}", level=1)  
            doc.add_picture(img_path, width=DocxInches(5))  
            os.remove(img_path)  # Clean up the image file  
  
    os.remove(ppt_path)  # Clean up the temporary PPT file  
  
    # Save the Word document to a BytesIO object  
    buffer = BytesIO()  
    doc.save(buffer)  
    buffer.seek(0)  
  
    st.download_button(  
        label="Download Word Document",  
        data=buffer,  
        file_name="slides_insights.docx",  
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"  
    )  
