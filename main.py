from langchain_core.prompts import ChatPromptTemplate, MessagesPlaceholder
from langchain_google_genai import ChatGoogleGenerativeAI
from langchain_core.tools import tool
import io
import pptx
from pptx.util import Inches
from typing import Optional
from flask import Flask, render_template, jsonify, request, send_file
import json
from dotenv import load_dotenv
import os

load_dotenv()
GOOGLE_API_KEY = os.getenv('gemini-api-key') # Store the api key in the key button present on the left sidebar shaped
MODEL_NAME = "gemini-1.5-pro-latest"
if not GOOGLE_API_KEY:
    raise ValueError("GOOGLE_API_KEY not found. Please set it in your .env file.")
os.environ['GOOGLE_API_KEY'] = GOOGLE_API_KEY
################################################# AGENT SETUP #################################################
prompt = ChatPromptTemplate.from_messages([
    ("system", """
    You are a presentation generation AI. Your only goal is to create a PowerPoint file by calling the `create_powerpoint` tool.
    To do this, you will first reformat the user's text into a specific, detailed markdown structure. You MUST follow the structure and formatting rules precisely.
    After creating the perfectly formatted markdown, call the `create_powerpoint` tool. Your job is complete only when the tool returns a success message.
    """),
    ("user", """
    Here are the precise formatting rules you MUST follow for each slide.

    **Slide 1: Title Slide**
    - The first line must be the main title, starting with `# Slide 1:`. **Generate a compelling title based on the text.**
    - The second line must be the subtitle, starting with `##`. **Generate a short, engaging subtitle.**

    ---

    **Slides 2 to (N-1): Content Slides**
    For each key idea from the text, create a slide. Every content slide **MUST** contain all four of the following elements in this exact order:
    1.  A title line starting with `# Slide [Number]:`. **Generate a descriptive title for the key idea.**
    2.  2-4 bullet points. Each bullet point must start with `-` and **summarize a specific detail from the text.**
    3.  Speaker notes, starting with the literal text `**Speaker Notes:**`. **Write brief, helpful notes for the presenter.**
    4.  A visual suggestion, starting with the literal text `**Visual Suggestion:**`. **Describe a relevant image, chart, or icon.**

    ---

    **Slide N: Conclusion Slide**
    The final slide **MUST** contain all four of the following elements in this exact order:
    1.  A title line: `# Slide [Number]: Conclusion`.
    2.  2-3 bullet points. Each must start with `-` and **summarize a main takeaway from the presentation.**
    3.  Speaker notes, starting with the literal text `**Speaker Notes:**`. **Write brief concluding remarks.**
    4.  A visual suggestion, starting with the literal text `**Visual Suggestion:**`. **Suggest a company logo or 'Thank You' image.**

    Now, create a presentation with exactly {number_of_slides} slides using the text below.

    <TEXT>
    {bulk_text}
    </TEXT>
    """),
    MessagesPlaceholder(variable_name="agent_scratchpad"),
])
llm = ChatGoogleGenerativeAI(model=MODEL_NAME)
agent_executor = llm | prompt
app = Flask(__name__)
################################################# HELPER FUNCTIONS #################################################
def create_powerpoint_in_memory(markdown_slides: str):
    """
    Creates a PowerPoint presentation from a markdown slides and returns the buffer
    """
    prs = pptx.Presentation()
    slides_content = [s.strip() for s in markdown_slides.strip().split('---') if s.strip()]

    for i, slide_markdown in enumerate(slides_content):
        lines = [line.strip() for line in slide_markdown.split('\n') if line.strip()]
        if not lines:
            continue

        if i == 0:
            slide_layout = prs.slide_layouts[0] # Title Slide
            slide = prs.slides.add_slide(slide_layout)
            slide.shapes.title.text = lines[0].replace('#', '').split(':', 1)[-1].strip()
            if len(lines) > 1:
                slide.placeholders[1].text = lines[1].replace('##', '').strip()
        else:
            slide_layout = prs.slide_layouts[1] # Title and Content
            slide = prs.slides.add_slide(slide_layout)
            slide.shapes.title.text = lines[0].replace('#', '').split(':', 1)[-1].strip()
            tf = slide.placeholders[1].text_frame
            tf.clear()
            for line in lines[1:]:
                if line.startswith('-'):
                    p = tf.add_paragraph()
                    p.text = line.lstrip('- ').strip()
                    p.level = 0
    
    # Save the presentation to an in-memory buffer
    buffer = io.BytesIO()
    prs.save(buffer)
    buffer.seek(0) # Rewind the buffer to the beginning
    return buffer
################################################# ROUTES #################################################
@app.route("/")
def index():
    return render_template("index.html")
@app.route("/create_ppt",methods=["POST"])
def create_ppt():
    global agent_executor
    data  = request.get_json()
    bulk_text = data.get("bulk_text")
    number_of_slides  = int(data.get("number_of_slides"))
    filename = data.get("filename", "presentation.pptx")
    result = agent_executor.invoke({
        "bulk_text": bulk_text,
        "number_of_slides": number_of_slides,
    })
    markdown_plan = result.content
    ppt_buffer = create_powerpoint_in_memory(markdown_plan)
    return send_file(
        ppt_buffer,
        as_attachment=True,
        download_name = filename,
        mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
    )
if __name__ == "__main__":
    app.run(debug=True)

