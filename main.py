from langchain_core.prompts import ChatPromptTemplate, MessagesPlaceholder
from langchain_google_genai import ChatGoogleGenerativeAI
from langchain_core.tools import tool
from langchain.agents import create_tool_calling_agent, AgentExecutor
import pptx
from pptx.util import Inches
from typing import Optional
from flask import Flask, render_template
from dotenv import load_dotenv
load_dotenv()
GOOGLE_API_KEY = userdata.get('gemini-api-key') # Store the api key in the key button present on the left sidebar shaped
MODEL_NAME = "gemini-1.5-pro-latest"
if not GOOGLE_API_KEY:
  raise NoneTypeException("GOOGLE_API_KEY not found. Please set the api key as gemini-api-key in left sidebar")
print("GOOGLE_API_KEY is loaded..")
os.environ['GOOGLE_API_KEY'] = GOOGLE_API_KEY
print("GOOGLE_API_KEY is set")

!pip install langchain_core langchain_google_genai langchain_experimental google-generativeai

!pip install python-pptx


@tool
def create_powerpoint(markdown_slides: str, filename: Optional[str] = None) -> str:
    """
    Creates a PowerPoint presentation from a markdown string.
    """
    if filename is None:
        filename = "presentation.pptx"

    prs = pptx.Presentation()
    slides = markdown_slides.strip().split('---')

    for i, slide_content in enumerate(slides):
        lines = [line for line in slide_content.strip().split('\n') if line.strip()]
        if not lines:
            continue

        # Use 'Title Slide' layout for the first slide, 'Title and Content' for the rest
        if i == 0:
            slide_layout = prs.slide_layouts[0] # Title Slide
            slide = prs.slides.add_slide(slide_layout)
            slide.shapes.title.text = lines[0].replace('#', '').strip()
            if len(lines) > 1:
                slide.placeholders[1].text = '\n'.join(line.replace('##', '').replace('###', '').strip() for line in lines[1:])
        else:
            slide_layout = prs.slide_layouts[1] # Title and Content
            slide = prs.slides.add_slide(slide_layout)
            slide.shapes.title.text = lines[0].replace('#', '').strip()
            tf = slide.placeholders[1].text_frame
            tf.clear()
            for line in lines[1:]:
                if line.strip().startswith('-'):
                    p = tf.add_paragraph()
                    p.text = line.strip().lstrip('-').strip()
                    p.level = 0

    prs.save(filename)
    return f"Success! The presentation has been saved as {filename}"
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
tools = [create_powerpoint]
agent = create_tool_calling_agent(llm, tools, prompt)
agent_executor = AgentExecutor(agent=agent, tools=tools, verbose=True)

def test_create_powerpoint():
    global agent_executor
    user_text = """The fundamental shift to cloud computing represents one of the most significant advancements in modern technology. The primary benefit for businesses is a massive reduction in capital expenditure, as they can avoid purchasing and maintaining expensive hardware. Another key advantage is scalability and flexibility. Cloud services allow businesses to scale their resources up or down almost instantly based on demand, preventing wasted resources. Finally, cloud computing enhances collaboration and accessibility. With data and applications hosted in the cloud, teams can access their work from anywhere in the world, significantly improving productivity."""
    slide_count = 4

    result = agent_executor.invoke({
        "bulk_text": user_text,
        "number_of_slides": slide_count
    })
    print("\n--- Final Output ---")
    print(result['output'])
test_create_powerpoint()

prs = pptx.Presentation()
prs.slide_layouts[7].name
slide = prs.slides.add_slide(prs.slide_layouts[1])
slide.placeholders[1].text_frame
