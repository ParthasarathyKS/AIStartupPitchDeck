
import logging
import re
import uuid
from io import BytesIO

import requests
from flask import Flask, render_template, request, send_file, url_for
from pptx import Presentation
from pptx.util import Inches, Pt
from docx import Document

app = Flask(__name__)
logging.basicConfig(level=logging.INFO)

OLLAMA_URL = "http://127.0.0.1:11434/api/generate"
MODEL_NAME = "gemma2:2b"
OLLAMA_TIMEOUT = 300

# Temporary storage for generated results.
# Results are cleared when the application restarts.
GENERATED_RESULTS = {}


def ask_ai(prompt):
    """Generate text using the local Ollama model."""
    response = requests.post(
        OLLAMA_URL,
        json={
            "model": MODEL_NAME,
            "prompt": prompt,
            "stream": False,
        },
        timeout=OLLAMA_TIMEOUT,
    )
    response.raise_for_status()

    data = response.json()
    answer = data.get("response", "").strip()

    if not answer:
        raise ValueError("Ollama returned an empty response.")

    return answer


def create_pitch_prompt(name, idea, audience, business_model, funding):
    return f"""
You are an expert startup pitch deck writer.

Create exactly seven slides for this startup.

Startup name: {name}
Startup idea: {idea}
Target audience: {audience}
Business model: {business_model}
Funding goal: {funding or "Not specified"}

Use these exact headings, each on a separate line:
SLIDE 1: Startup Introduction
SLIDE 2: Problem Statement
SLIDE 3: Proposed Solution
SLIDE 4: Target Market
SLIDE 5: Business Model
SLIDE 6: Competitive Advantage
SLIDE 7: Growth Strategy and Funding

Under each heading, provide 3-5 concise bullet points.
Make the content specific to the supplied startup.
Do not invent real customers, partnerships, revenue, or statistics.
Label assumptions clearly when facts are unavailable.
Write plain text, not a table or JSON.
"""


def create_video_prompt(name, idea, audience, duration):
    return f"""
You are a professional explainer-video scriptwriter.

Startup name: {name}
Startup idea: {idea}
Target audience: {audience}
Requested video duration: {duration}

Write an explainer video script suitable for the requested duration.

Include:
1. Opening hook
2. Scene-by-scene visual descriptions
3. Narration for every scene
4. Suggested on-screen text
5. Closing call to action

Use clear, natural language.
Keep the script appropriate for the requested duration.
Do not claim unverified benefits are proven.
Use readable headings and plain text.
"""


def split_slides(text):
    """Split an AI response into slide titles and content."""
    slides = []
    current_title = "Startup Pitch"
    current_lines = []

    heading_pattern = re.compile(
        r"^\s*(?:SLIDE\s*)?(\d+)\s*[:.)-]\s*(.+?)\s*$",
        re.IGNORECASE,
    )

    for line in text.splitlines():
        match = heading_pattern.match(line)

        if match:
            if current_lines:
                slides.append({
                    "title": current_title,
                    "content": "\n".join(current_lines).strip(),
                })

            current_title = f"Slide {match.group(1)}: {match.group(2)}"
            current_lines = []
        elif line.strip():
            current_lines.append(line.strip())

    if current_lines:
        slides.append({
            "title": current_title,
            "content": "\n".join(current_lines).strip(),
        })

    if not slides:
        slides = [{
            "title": "AI-Generated Startup Pitch",
            "content": text.strip(),
        }]

    return slides


@app.route("/")
def home():
    return render_template("index.html")


@app.route("/generate", methods=["POST"])
def generate():
    name = request.form.get("startup_name", "").strip()
    idea = request.form.get("idea", "").strip()
    audience = request.form.get("target_audience", "").strip()
    business_model = request.form.get("business_model", "").strip()
    funding = request.form.get("funding_goal", "").strip()
    duration = request.form.get("video_duration", "60 seconds").strip()
    include_video = request.form.get("include_video") == "yes"

    if not all([name, idea, audience, business_model]):
        return render_template(
            "index.html",
            error="Please complete all required fields.",
        ), 400

    try:
        app.logger.info("Generating pitch deck for %s", name)

        pitch_text = ask_ai(
            create_pitch_prompt(
                name, idea, audience, business_model, funding
            )
        )
        pitch_deck = split_slides(pitch_text)

        video_script = None
        if include_video:
            app.logger.info("Generating explainer script")
            video_script = ask_ai(
                create_video_prompt(name, idea, audience, duration)
            )

        result_id = str(uuid.uuid4())

        GENERATED_RESULTS[result_id] = {
            "startup_name": name,
            "pitch_deck": pitch_deck,
            "video_script": video_script,
            "video_duration": duration,
        }

        # Keep this beginner project from growing indefinitely.
        while len(GENERATED_RESULTS) > 20:
            oldest_id = next(iter(GENERATED_RESULTS))
            GENERATED_RESULTS.pop(oldest_id)

        return render_template(
            "result.html",
            result_id=result_id,
            **GENERATED_RESULTS[result_id],
        )

    except requests.exceptions.ConnectionError:
        app.logger.exception("Cannot connect to Ollama")
        return render_template(
            "index.html",
            error=(
                "Cannot connect to Ollama. Check that Ollama is running "
                "at http://127.0.0.1:11434."
            ),
        ), 503

    except requests.exceptions.Timeout:
        app.logger.exception("Ollama timed out")
        return render_template(
            "index.html",
            error="The AI took too long to respond. Please try again.",
        ), 504

    except requests.exceptions.RequestException:
        app.logger.exception("Ollama request failed")
        return render_template(
            "index.html",
            error="The AI service returned an error. Check the VS Code terminal.",
        ), 502

    except Exception as e:
        app.logger.exception("Startup pitch generation failed")
        return render_template(
            "index.html",
            error=f"Generation failed: {type(e).__name__}: {e}",
        ), 500


@app.route("/download/pptx/<result_id>")
def download_pptx(result_id):
    result = GENERATED_RESULTS.get(result_id)

    if not result:
        return "Result not found. Please generate your pitch again.", 404

    presentation = Presentation()

    # Title slide
    slide = presentation.slides.add_slide(
        presentation.slide_layouts[0]
    )
    slide.shapes.title.text = result["startup_name"]
    slide.placeholders[1].text = "AI-Generated Startup Pitch Deck"

    # Content slides
    for item in result["pitch_deck"]:
        slide = presentation.slides.add_slide(
            presentation.slide_layouts[1]
        )
        slide.shapes.title.text = item["title"]

        text_frame = slide.placeholders[1].text_frame
        text_frame.clear()

        for index, line in enumerate(item["content"].splitlines()):
            line = line.strip()
            if not line:
                continue

            paragraph = (
                text_frame.paragraphs[0]
                if index == 0 and not text_frame.paragraphs[0].text
                else text_frame.add_paragraph()
            )
            paragraph.text = line.lstrip("•-* \t")
            paragraph.font.size = Pt(20)

    if result["video_script"]:
        slide = presentation.slides.add_slide(
            presentation.slide_layouts[1]
        )
        slide.shapes.title.text = "Explainer Video Script"
        slide.placeholders[1].text = result["video_script"]

    buffer = BytesIO()
    presentation.save(buffer)
    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name="startup_pitch_deck.pptx",
        mimetype=(
            "application/vnd.openxmlformats-officedocument."
            "presentationml.presentation"
        ),
    )


@app.route("/download/docx/<result_id>")
def download_docx(result_id):
    result = GENERATED_RESULTS.get(result_id)

    if not result:
        return "Result not found. Please generate your pitch again.", 404

    document = Document()
    document.add_heading(
        f"{result['startup_name']} - Explainer Video Script", level=0
    )
    document.add_paragraph(
        f"Requested duration: {result['video_duration']}"
    )

    if result["video_script"]:
        for line in result["video_script"].splitlines():
            line = line.strip()
            if line:
                document.add_paragraph(line)
    else:
        document.add_paragraph(
            "No video script was generated for this pitch."
        )

    buffer = BytesIO()
    document.save(buffer)
    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name="explainer_video_script.docx",
        mimetype=(
            "application/vnd.openxmlformats-officedocument."
            "wordprocessingml.document"
        ),
    )


if __name__ == "__main__":
    app.run(debug=True)