# AI Question Paper Analyzer — V4

A deployment-ready Streamlit prototype for analysing school Physics question papers.

## Features

- PDF, DOCX, TXT and image upload
- Fixed 12-chapter Physics syllabus
- Question, sub-question and sub-sub-question extraction
- Exact chapter name displayed with every question
- Teacher-editable chapter classification
- Dynamic chapter-wise marks and pie-chart weightage
- Printed marks vs maximum exam marks
- Main concepts to focus on
- Question type and difficulty
- On-demand AI answer generation
- Word export: sorted paper
- Word export: sorted paper + AI answers
- Start New Paper button
- API key read from Streamlit Secrets or environment variable
- No API key stored in source code

## GitHub files

Commit these files to your repository:

- `app.py`
- `requirements.txt`
- `README.md`
- `.gitignore`

Do NOT commit `.streamlit/secrets.toml`.

## Local setup

```bash
pip install -r requirements.txt
```

Set your API key:

### Windows PowerShell

```powershell
$env:OPENAI_API_KEY="your_api_key"
streamlit run app.py
```

### Streamlit Cloud

In the deployed app settings, add:

```toml
OPENAI_API_KEY = "your_api_key"
```

Optional model:

```toml
OPENAI_MODEL = "gpt-5.6-luna"
```

## Important

AI-generated classifications and answers should be reviewed before being used
as official teaching or marking material.

## Recommended next features

- Preserve original diagrams in the exported Word document
- Detect and retain section/choice relationships at sub-question level
- Add chapter priority scoring
- Add repeated-concept analysis across multiple papers
- Add multi-paper/year comparison
- Add authentication if the app is shared beyond a small group
