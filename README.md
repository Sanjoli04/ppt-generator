# AI Presentation Generator

An intelligent web application that transforms raw, unstructured text into a fully formatted and downloadable PowerPoint (.pptx) presentation. The application uses Google's **Gemini Pro** model via **LangChain** to first generate a structured plan, which is then converted into a presentation file, providing an interactive and intelligent user experience.

<!-- TODO: Add a screenshot of your application here -->

---

✨ **Features**

* **Dynamic Content Planning**: The AI first generates a logical plan for the presentation, which is displayed to the user in real-time with a typewriter effect, making the process feel interactive and intelligent.
* **Custom Slide Count**: Users can specify the exact number of slides they want in the final presentation, giving them full control over the output.
* **In-Memory File Generation**: PowerPoint files are created entirely in memory and streamed directly to the user. This is highly efficient and secure as no files are ever stored on the server.
* **Interactive UI**: A clean, multi-step user interface guides the user through the process, from input to final download.
* **Ready for Deployment**: Comes with a Dockerfile and `render.yaml` for easy, one-click deployment on a platform like Render.

---

🛠️ **Tech Stack**

* **Backend**: Flask, Gunicorn
* **AI/LLM**: Google Gemini Pro, LangChain
* **File Generation**: python-pptx
* **Frontend**: HTML, Tailwind CSS
* **Deployment**: Docker, Render

---

📂 **Project Structure**

```
├── .env                # For storing secret API keys locally
├── Dockerfile          # Blueprint for building the production container
├── main.py             # The core Flask application logic and API endpoints
├── README.md           # This file
├── render.yaml         # Infrastructure-as-Code for deployment on Render
├── requirements.txt    # Python dependencies
└── templates/
    └── index.html      # The main HTML file for the user interface
```

---

🚀 **Getting Started**

Follow these instructions to get the project running on your local machine for development and testing.

### Prerequisites

* Python 3.9+
* A Google Gemini API Key

### Local Installation

1. Clone the repository:

   ```bash
   git clone https://github.com/Sanjoli04/ppt-generator.git
   cd ppt-generator
   ```

2. Create a virtual environment:

   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows, use `venv\Scripts\activate`
   ```

3. Install the dependencies:

   ```bash
   pip install -r requirements.txt
   ```

4. Set up your environment variables:

   * Create a file named `.env` in the root of the project and add your API key:

     ```bash
     GOOGLE_API_KEY="your_google_gemini_api_key_here"
     ```

5. Run the Flask application:

   ```bash
   flask run
   ```

6. The application will be available at:
   👉 [http://127.0.0.1:5000](http://127.0.0.1:5000)

---

⚙️ **API Endpoints**

The application uses a two-endpoint process to create an interactive experience:

1. **`/generate_plan`**

   * **Method**: POST
   * **Description**: Takes the user's raw text and slide count, calls the AI to generate a markdown plan for the presentation.
   * **Request Body (JSON)**:

     ```json
     {
       "bulk_text": "The text for the presentation...",
       "number_of_slides": 5
     }
     ```
   * **Response (JSON)**:

     ```json
     {
       "markdown_plan": "# Slide 1: Title..."
     }
     ```

2. **`/create_file`**

   * **Method**: POST
   * **Description**: Takes the AI-generated markdown plan and converts it into a `.pptx` file.
   * **Request Body (JSON)**:

     ```json
     {
       "markdown_plan": "# Slide 1: Title...",
       "filename": "Generated-Presentation.pptx"
     }
     ```
   * **Response**: Streams a `.pptx` file directly to the user, triggering a download.

---

☁️ **Deployment to Render**

This project is configured for easy deployment on Render using the provided `render.yaml` and `Dockerfile`.

1. Push your code to a GitHub/GitLab repository.
2. Create a new **"Blueprint" service** on Render:

   * On the Render dashboard, click **New → Blueprint**.
   * Connect the repository you just created. Render will automatically detect and use your `render.yaml` file.
3. Add your API Key as a Secret File:

   * In your service's dashboard, go to the **Environment** tab.
   * Under **Secret Files**, click **Add Secret File**.
   * **Filename**: `gemini-api-key-secret` (This must match the name in `render.yaml`).
   * **Contents**: Paste your Google Gemini API key.
   * Click **Save Changes**.
4. **Deploy**: Trigger a manual deploy or push a new commit to your main branch. Render will automatically build the Docker image and deploy your application.
