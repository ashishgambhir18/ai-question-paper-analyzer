# AI Question Paper Analyzer — V16

Streamlit app for analysing CISCE question papers against the official syllabus.

## V16 changes

- **Live sorted paper:** questions are first extracted, then classified **one at a time**. As soon as a question is classified, it is displayed under its chapter instead of waiting for the complete paper analysis.
- **Original MCQs preserved:** the extractor is explicitly instructed to keep the complete MCQ stem and all options. Candidate-facing instructions such as “write the correct answers only” are ignored by the analyzer.
- **AI answers remain separate:** an AI answer is generated only when **View AI Answer** is clicked. It is never substituted for the original question in the sorted paper.
- **Original PDF text mirror:** PyMuPDF text is supplied alongside the PDF during extraction to improve fidelity of question wording and options.
- Retains V15 automatic CISCE 2027 syllabus retrieval, chapter navigation, Bloom analysis, weightage, figure detection, automatic figure preview, manual crop, and Word export.

## Run

```bash
pip install -r requirements.txt
streamlit run app.py
```

Add `OPENAI_API_KEY` to Streamlit Secrets or the environment.

## Current automatic syllabus retrieval

Configured for CISCE Examination Year 2027:

- ICSE: https://cisce.org/regulations-and-syllabus-icse-2027/
- ISC: https://cisce.org/regulations-and-syllabus-isc-2027/

Older years remain selectable in the interface but automatic retrieval currently stops with a clear message rather than applying a modern syllabus to an older paper.
