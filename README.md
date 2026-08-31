# AI Question Paper Analyzer — V6

## V6 changes

### Correct paper-weightage model

For the benchmark paper:
- Section A = 40 compulsory marks
- Section B = 6 printed questions × 10 marks = 60 printed marks
- Complete printed paper = **100 marks**
- Student attempts 4 Section B questions = 40 marks
- Actual examination maximum = **80 marks**

Therefore the **chapter-weightage pie chart is calculated out of the complete
100-mark printed paper**. The app separately displays the 80-mark actual exam
maximum.

This is intentional: the tool analyses the question paper setter's chapter
weightage, not one student's optional-question combination.

### Figures / diagrams

For PDF uploads, V6 records the source page and normalized figure coordinates.
It crops the original PDF visual and shows it below the sorted question and in
the Word export.

## Files

- `app.py`
- `requirements.txt`
- `README.md`
- `.gitignore`

Do not commit `.streamlit/secrets.toml` or API keys.

## Run

```bash
pip install -r requirements.txt
streamlit run app.py
```

Streamlit Secrets:

```toml
OPENAI_API_KEY = "your_api_key"
OPENAI_MODEL = "gpt-5.6-luna"
```
