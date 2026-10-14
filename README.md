# 📄 ATS Resume Checker

A Streamlit app that scores a resume for ATS (Applicant Tracking System) compatibility and tells you exactly how to improve it. Powered by Google's Gemini Flash model.

## Features

- Upload a **PDF, DOCX or TXT** resume (up to 5 MB)
- Optional **job description** for a tailored keyword match and match percentage
- **ATS score out of 100** with a weighted breakdown:

  | Category | Weight |
  |---|---|
  | Keywords & Skills | 30% |
  | Content & Impact | 25% |
  | Formatting & ATS Parseability | 20% |
  | Structure & Sections | 15% |
  | Language & Readability | 10% |

- Missing keywords, strengths, weaknesses
- Prioritised improvements (High / Medium / Low)
- Before/after bullet-point rewrites
- Rule-based checks (contact info, section headings, length, quantified results)
- Download the full report as Markdown
- Shows the extracted text so you can see what an ATS actually reads

## How the score works

Gemini rates each category from 0 to 100. The app then calculates the overall score itself using the weights above, so the maths is always consistent. The rule-based checks are shown separately and do not change the score.

> ATS scores are **estimates**. Every company's ATS works differently and none publishes its algorithm. Use the score as guidance, not a guarantee.

## Project structure

```
.
├── app.py             # the whole application
├── requirements.txt   # Python dependencies
└── README.md
```

## Run locally

1. **Get a free Gemini API key** at <https://aistudio.google.com/apikey>.

2. **Install dependencies** (Python 3.10+ recommended):

   ```bash
   python -m venv .venv
   source .venv/bin/activate        # Windows: .venv\Scripts\activate
   pip install -r requirements.txt
   ```

3. **Add your API key** - choose one:

   - Create `.streamlit/secrets.toml`:

     ```toml
     GEMINI_API_KEY = "your-api-key-here"
     ```

   - or set an environment variable:

     ```bash
     export GEMINI_API_KEY="your-api-key-here"      # Windows PowerShell: $env:GEMINI_API_KEY="..."
     ```

   - or simply paste the key into the app's sidebar at runtime.

4. **Start the app**:

   ```bash
   streamlit run app.py
   ```

> ⚠️ **Never commit `.streamlit/secrets.toml` or your API key to GitHub.** If you use Git from the command line, add `.streamlit/secrets.toml` to a `.gitignore` file.

## Configuration

| Setting | Where | Default |
|---|---|---|
| `GEMINI_API_KEY` | Streamlit secrets or environment variable | none (required) |
| `GEMINI_MODEL` | Streamlit secrets or environment variable, or sidebar → Advanced | `gemini-3.6-flash` |

Google retires and renames models regularly. If the configured model is not found, the app automatically falls back to `gemini-flash-latest`. You can also change the model name in the sidebar (Advanced) without touching the code. Current names are listed at <https://ai.google.dev/gemini-api/docs/models>.

## Deploy on Streamlit Community Cloud

1. Push this project to a GitHub repository (`app.py`, `requirements.txt` and `README.md` in the repo root).
2. Go to <https://share.streamlit.io> and sign in with GitHub.
3. Click **Create app** → **Deploy a public app from GitHub**.
4. Choose your repository, branch `main`, and main file path `app.py`.
5. Open **Advanced settings**, pick a Python version (3.11 or 3.12), and paste into **Secrets**:

   ```toml
   GEMINI_API_KEY = "your-api-key-here"
   ```

6. Click **Deploy**. Every future commit to the branch redeploys the app automatically.

If the app is public, anyone with the link can use **your** API key quota. Keep an eye on your usage in Google AI Studio, or leave the secret empty so visitors must enter their own key in the sidebar.

## Privacy

Resume text is sent to Google's Gemini API for analysis. This app does not store uploaded files or results; they live only in your browser session. Check Google's current Gemini API terms regarding data use before uploading sensitive documents, especially on a free-tier key.

## Limitations

- **Scanned / image-only resumes** can't be read (an ATS can't read them either). Use a text-based PDF or DOCX.
- Formatting is judged from the **extracted text**, not the visual layout. The model can spot signs of problems (merged columns, tables, garbled text) but can't see your design.
- Only the first 10 pages of a PDF and about 30,000 characters are analysed.
- AI feedback can be wrong. Never add skills or achievements to your resume that aren't true.

## Troubleshooting

| Problem | Fix |
|---|---|
| "API key was rejected" | Re-copy the key from Google AI Studio; make sure there are no extra spaces. |
| "Model not found" | Sidebar → Advanced → enter a current model name from the Gemini models page. |
| Rate limit / quota message | Wait a minute and retry, or check your quota in Google AI Studio. |
| "Could barely extract any text" | Your file is probably a scan or image. Export a text-based PDF/DOCX. |
| Build fails on Streamlit Cloud | Make sure `requirements.txt` is in the repo root and the main file is `app.py`. |
