import os
import io
import pptx
from flask import Flask, render_template, jsonify, request, send_file
from langchain_core.prompts import ChatPromptTemplate
from langchain_google_genai import ChatGoogleGenerativeAI
from dotenv import load_dotenv

# --- 1. Setup and Configuration ---
load_dotenv()
FALLBACK_GOOGLE_API_KEY = os.getenv("GOOGLE_API_KEY")
MODEL_NAME = "gemini-1.5-pro-latest"

# --- 2. The AI Chain (Re-written for Research and Content Generation) ---
# This prompt now instructs the AI to research a topic and create content.
prompt = ChatPromptTemplate.from_messages([
    ("system", """
    You are an expert research assistant and presentation designer. Your goal is to generate a detailed markdown string for a presentation on a given topic.
    You will research the topic to gather concise, accurate information and then structure it into the requested number of slides.
    Your final output must ONLY be the markdown string, following the specified format with `---` separators.
    """),
    ("user", """
    Here are the precise formatting rules you MUST follow for each slide.

    **Slide 1: Title Slide**
    - A title line starting with `#`. The title should be compelling and relevant to the topic.
    - A subtitle line starting with `##`. The subtitle should be a short, engaging summary.
    ---
    **Content Slides**
    - Each slide MUST have a title line starting with `#`.
    - Each slide MUST have 2-4 concise bullet points summarizing key information about the topic.
    ---
    **Conclusion Slide**
    - The final slide MUST have a title line starting with `#`.
    - The final slide MUST have 2-3 summary bullet points of the main takeaways.

    **Optional Guidance for tone and structure:** {guidance}

    Now, please research the following topic and create the markdown for a presentation with exactly {number_of_slides} slides.

    **TOPIC:** {topic}
    """),
])

# --- 3. Style Extraction and PowerPoint Generation ---
def get_layout_from_template(prs, layout_name):
    """Finds a slide layout by its name in the presentation."""
    for master in prs.slide_masters:
        for layout in master.slide_layouts:
            if layout.name == layout_name:
                return layout
    # Fallback if a named layout isn't found
    if "title" in layout_name.lower():
        return prs.slide_layouts[0]
    return prs.slide_layouts[1]

def create_ppt_with_template(markdown_slides: str, template_file):
    """Creates a PowerPoint presentation applying styles from a template file."""
    prs = pptx.Presentation(template_file)
    
    # Remove all existing slides from the template
    while len(prs.slides) > 0:
        rId = prs.slides._sldIdLst[0].rId
        prs.part.drop_rel(rId)
        del prs.slides._sldIdLst[0]

    slides_content = [s.strip() for s in markdown_slides.strip().split('---') if s.strip()]

    title_layout = get_layout_from_template(prs, 'Title Slide')
    content_layout = get_layout_from_template(prs, 'Title and Content')

    for i, slide_markdown in enumerate(slides_content):
        lines = [line.strip() for line in slide_markdown.split('\n') if line.strip()]
        if not lines:
            continue

        if i == 0:
            slide = prs.slides.add_slide(title_layout)
            title = slide.shapes.title
            subtitle = slide.placeholders[1] if len(slide.placeholders) > 1 else None
            title.text = lines[0].replace('#', '').strip()
            if subtitle and len(lines) > 1:
                subtitle.text = lines[1].replace('##', '').strip()
        else:
            slide = prs.slides.add_slide(content_layout)
            title = slide.shapes.title
            body = slide.placeholders[1] if len(slide.placeholders) > 1 else None
            title.text = lines[0].replace('#', '').strip()
            if body:
                tf = body.text_frame
                tf.clear()
                for line in lines[1:]:
                    if line.startswith('-'):
                        p = tf.add_paragraph()
                        p.text = line.lstrip('- ').strip()
                        p.level = 0
    
    buffer = io.BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer

# --- 4. Flask Web Application ---
app = Flask(__name__)

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/generate_presentation", methods=['POST'])
def generate_presentation():
    """Single endpoint to handle all inputs and generate the presentation."""
    if 'template_file' not in request.files:
        return jsonify({"error": "No template file provided."}), 400

    template_file = request.files['template_file']
    topic = request.form.get('topic', '')
    guidance = request.form.get('guidance', 'A standard professional presentation.')
    user_api_key = request.form.get('api_key', '')
    number_of_slides = int(request.form.get('number_of_slides', 3))

    api_key_to_use = user_api_key or FALLBACK_GOOGLE_API_KEY
    if not api_key_to_use:
        return jsonify({"error": "No API key provided or configured on the server."}), 400

    try:
        llm = ChatGoogleGenerativeAI(model=MODEL_NAME, google_api_key=api_key_to_use)
        chain = prompt | llm

        response = chain.invoke({
            "topic": topic,
            "guidance": guidance,
            "number_of_slides": number_of_slides
        })
        markdown_plan = response.content.strip().replace("```markdown", "").replace("```", "")

        ppt_buffer = create_ppt_with_template(markdown_plan, template_file)
        
        return send_file(
            ppt_buffer,
            as_attachment=True,
            download_name="generated_presentation.pptx",
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )
    except Exception as e:
        print(f"An error occurred: {e}")
        if "API key not valid" in str(e):
             return jsonify({"error": "The provided API key is not valid. Please check it and try again."}), 401
        return jsonify({"error": "An error occurred while generating the presentation."}), 500

if __name__ == '__main__':
    app.run(debug=True)
