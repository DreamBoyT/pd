import os  
import streamlit as st  
from pptx import Presentation  
from docx import Document  
from io import BytesIO  
from langchain_openai import AzureChatOpenAI  
from langchain.prompts import PromptTemplate  
  
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
  
# Function to extract text and title from ppt slides  
def extract_text_and_title_from_ppt(ppt_file):  
    prs = Presentation(ppt_file)  
    slides_data = []  
    for slide in prs.slides:  
        slide_text = []  
        slide_title = None  
        for shape in slide.shapes:  
            if hasattr(shape, "text"):  
                if shape == slide.shapes.title:  
                    slide_title = shape.text  
                else:  
                    slide_text.append(shape.text)  
        slides_data.append((slide_title, "\n".join(slide_text)))  
    return slides_data  
  
# Function to generate explanation using Azure OpenAI  
def generate_explanation(slide_text):  
    prompt_template = PromptTemplate(  
        input_variables=["slide_text"],  
        template="""  
        Slide Content: {slide_text}  
  
        Aspects of the present disclosure may include insights extracted from the above slide content. The information should be delivered directly and engagingly. Avoid phrases like 'The slide presents,' 'discusses,' 'outlines,' or 'content.' The explanation should be formatted as paragraphs, without line breaks or bullet points, and must be semantically meaningful.  
  
        The text should adhere to the following style guidelines:  
        1. Remove all listed profanity words.  
        2. Use passive voice.  
        3. Use conditional and tentative language, such as "may include," "in some aspects," and "aspects of the present disclosure."  
        4. Replace "Million" with "1,000,000" and "Billion" with "1,000,000,000".  
        5. Avoid using adjectives, superlatives, or any terms that imply absolute certainty.  
          
        It is crucial to strictly adhere to the above guidelines to ensure the highest quality and most accurate output.  
        """  
    )  
    prompt = prompt_template.format(slide_text=slide_text)  
    response = llm(prompt)  
    return response.content  
  
# Streamlit app  
st.title("PPT Insights Extractor")  
uploaded_file = st.file_uploader("Upload a PPT file", type=["pptx"])  
  
if uploaded_file is not None:  
    slides_data = extract_text_and_title_from_ppt(uploaded_file)  
  
    explanations = []  
    for slide_title, slide_text in slides_data:  
        explanation = generate_explanation(slide_text)  
        explanations.append((slide_title, explanation))  
  
    # Create a Word document  
    doc = Document()  
    for slide_title, explanation in explanations:  
        if slide_title:  
            doc.add_heading(slide_title, level=1)  
        else:  
            doc.add_heading("Untitled Slide", level=1)  
        doc.add_paragraph(explanation)  
  
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
