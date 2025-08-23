import os
import io
import pptx
from flask import Flask, render_template, jsonify, request, send_file
from langchain_core.prompts import ChatPromptTemplate
from langchain_google_genai import ChatGoogleGenerativeAI
from dotenv import load_dotenv

# --- 1. Setup and Configuration ---
load_dotenv()
GOOGLE_API_KEY = os.getenv("GOOGLE_API_KEY")
MODEL_NAME = "gemini-1.5-pro-latest"

if not GOOGLE_API_KEY:
    raise ValueError("GOOGLE_API_KEY not found. Please set it in your .env file.")

os.environ['GOOGLE_API_KEY'] = GOOGLE_API_KEY

# --- 2. The AI Chain (for Markdown Generation) ---
# This prompt is highly detailed to ensure the AI generates a complete and well-structured plan.
prompt = ChatPromptTemplate.from_messages([
    ("system", """
    You are an expert presentation designer AI. Your only goal is to generate a detailed markdown string for a presentation based on the user's text and slide count.
    You MUST follow the structure and formatting rules precisely, especially the `---` slide separators.
    Your final output must ONLY be the markdown string. Do not add any other conversation or explanation.
    """),
    ("user", """
    Here are the precise formatting rules you MUST follow for each slide.

    **Slide 1: Title Slide**
    - The first line must be the main title, starting with `# Slide 1:`. Generate a compelling title based on the text.
    - The second line must be the subtitle, starting with `##`. Generate a short, engaging subtitle.
    ---
    **Slides 2 to (N-1): Content Slides**
    - Each content slide MUST contain a title line starting with `# Slide [Number]:`.
    - Each content slide MUST have 2-4 bullet points starting with `-`.
    - Each content slide MUST have Speaker Notes starting with `**Speaker Notes:**`.
    - Each content slide MUST have a Visual Suggestion starting with `**Visual Suggestion:**`.
    ---
    **Slide N: Conclusion Slide**
    - The final slide MUST have a title line: `# Slide [Number]: Conclusion`.
    - The final slide MUST have 2-3 bullet points summarizing takeaways.
    - The final slide MUST have Speaker Notes and a Visual Suggestion.

    Now, create a presentation with exactly {number_of_slides} slides using the text below.

    <TEXT>
    {bulk_text}
    </TEXT>
    """),
])

llm = ChatGoogleGenerativeAI(model=MODEL_NAME)
# A simple chain: the user input is formatted by the prompt and then sent to the LLM.
chain = prompt | llm

# --- 3. PowerPoint Helper Function ---
def create_powerpoint_in_memory(markdown_slides: str):
    """Creates a PowerPoint presentation from a markdown string and returns the in-memory buffer."""
    prs = pptx.Presentation()
    slides_content = [s.strip() for s in markdown_slides.strip().split('---') if s.strip()]

    for i, slide_markdown in enumerate(slides_content):
        lines = [line.strip() for line in slide_markdown.split('\n') if line.strip()]
        if not lines:
            continue

        if i == 0:
            slide_layout = prs.slide_layouts[0]  # Title Slide
            slide = prs.slides.add_slide(slide_layout)
            slide.shapes.title.text = lines[0].replace('#', '').split(':', 1)[-1].strip()
            if len(lines) > 1:
                slide.placeholders[1].text = lines[1].replace('##', '').strip()
        else:
            slide_layout = prs.slide_layouts[1]  # Title and Content
            slide = prs.slides.add_slide(slide_layout)
            slide.shapes.title.text = lines[0].replace('#', '').split(':', 1)[-1].strip()
            tf = slide.placeholders[1].text_frame
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

@app.route("/generate_plan", methods=['POST'])
def generate_plan():
    """Endpoint to generate the markdown plan from the AI."""
    data = request.get_json()
    if not data or "bulk_text" not in data or "number_of_slides" not in data:
        return jsonify({"error": "Invalid request"}), 400

    try:
        response = chain.invoke({
            "bulk_text": data["bulk_text"],
            "number_of_slides": data["number_of_slides"]
        })
        return jsonify({"markdown_plan": response.content})
    except Exception as e:
        print(f"Error during AI plan generation: {e}")
        return jsonify({"error": "Failed to generate AI plan."}), 500

@app.route("/create_file", methods=['POST'])
def create_file():
    """Endpoint to create the PPT file from the markdown plan."""
    data = request.get_json()
    if not data or "markdown_plan" not in data:
        return jsonify({"error": "Invalid request"}), 400

    try:
        ppt_buffer = create_powerpoint_in_memory(data["markdown_plan"])
        filename = data.get("filename", "presentation.pptx")
        return send_file(
            ppt_buffer,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )
    except Exception as e:
        print(f"Error during file creation: {e}")
        return jsonify({"error": "Failed to create presentation file."}), 500

if __name__ == '__main__':
    app.run(debug=True)
