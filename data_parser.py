import os
import json
import mimetypes
from dotenv import load_dotenv
import google.generativeai as genai

# Load environment variables (such as GEMINI_API_KEY)
load_dotenv()

api_key = os.getenv("GEMINI_API_KEY")
if api_key:
    genai.configure(api_key=api_key)

def parse_blood_test(file_content: bytes, filename: str, model_name: str = "gemini-3.5-flash") -> dict:
    """
    Parses a blood test file (PDF, image, or text) using Gemini.
    If the requested model fails with a 429 or quota limit, it systematically
    falls back to alternative models.
    """
    if not api_key:
        raise ValueError("GEMINI_API_KEY is not configured in the environment.")

    # Guess MIME type based on filename
    mime_type, _ = mimetypes.guess_type(filename)
    if not mime_type:
        # Fallbacks
        if filename.endswith(".pdf"):
            mime_type = "application/pdf"
        elif filename.endswith(".txt"):
            mime_type = "text/plain"
        elif filename.endswith((".png", ".jpg", ".jpeg", ".webp")):
            mime_type = "image/png"
        else:
            mime_type = "application/octet-stream"

    prompt = """
    You are an expert medical data extraction assistant.
    Analyze this blood test result document. Extract:
    1. The date and time when the test was collected/performed.
       Format the date/time as: 'YYYY-MM-DD HH:MM AM/PM' (for example: '2023-03-21 09:15 AM').
       If only a date is found without a time, format it as 'YYYY-MM-DD 08:00 AM' (defaulting to 8:00 AM).
       If no date is found, use the current date or leave it empty, but look very carefully for any collection date or report date.
    2. All blood test markers/names and their corresponding numeric/text values.
       For example: 'WBC' -> '5.2', 'Hemoglobin' -> '14.5'.
       Do not include reference ranges in the values, just the reported value for the patient.

    Return the output strictly as a JSON object matching this structure:
    {
      "test_date": "YYYY-MM-DD HH:MM AM/PM",
      "results": [
        {
          "test_name": "Name of the test (e.g. WBC, RBC, Glucose, Estradiol)",
          "value": "Patient value (e.g. 5.2, 14.5, Negative)"
        }
      ]
    }

    Respond ONLY with the raw JSON. Do not include markdown code block formatting (like ```json) or any other text.
    """

    # If it's a plain text file, we can pass it as a string
    if mime_type.startswith("text/"):
        text_content = file_content.decode("utf-8", errors="ignore")
        contents = [prompt, text_content]
    else:
        # Pass binary content as a part
        contents = [
            prompt,
            {
                "mime_type": mime_type,
                "data": file_content
            }
        ]

    # Model attempts queue
    models_to_try = [model_name]
    alternatives = ["gemini-3.5-flash", "gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash", "gemini-flash-latest"]
    for alt in alternatives:
        if alt not in models_to_try:
            models_to_try.append(alt)

    last_error = None
    for current_model_name in models_to_try:
        print(f"Attempting parse with model: {current_model_name}")
        try:
            model = genai.GenerativeModel(current_model_name)
            response = model.generate_content(contents)
            response_text = response.text.strip()
            
            # Clean the response in case it contains markdown formatting
            if response_text.startswith("```"):
                lines = response_text.splitlines()
                if lines[0].startswith("```"):
                    lines = lines[1:]
                if lines[-1].startswith("```"):
                    lines = lines[:-1]
                response_text = "\n".join(lines).strip()

            data = json.loads(response_text)
            return data
            
        except Exception as e:
            last_error = e
            err_msg = str(e).lower()
            print(f"Warning: Model {current_model_name} failed: {e}")
            # If rate limited (429), quota exceeded, or model not found, try the next one
            if "429" in err_msg or "quota" in err_msg or "not found" in err_msg or "not_found" in err_msg:
                continue
            else:
                # Other errors (like JSON parsing issues, etc.), try to fallback or raise
                if current_model_name != models_to_try[-1]:
                    continue
                raise e

    raise last_error

if __name__ == "__main__":
    # Small test block
    test_file = r"c:\Users\admin\Desktop\emma-ai-capstone\images-for-excel\Blood_Tests_Tracker_P1.png"
    if os.path.exists(test_file):
        print(f"Testing parser with {test_file}...")
        with open(test_file, "rb") as f:
            content = f.read()
        res = parse_blood_test(content, "Blood_Tests_Tracker_P1.png")
        print(json.dumps(res, indent=2))
    else:
        print("Test file not found.")
