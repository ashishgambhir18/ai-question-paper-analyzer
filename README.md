# AI Question Paper Analyzer V21

Streamlit app for analysing school question papers against a syllabus supplied by the user.

## V21 changes
- Automatic CISCE syllabus retrieval has been completely removed.
- Upload the syllabus PDF yourself.
- The uploaded syllabus is the source of truth for chapter/unit classification.
- Paper Year is retained only as reference information; it does not alter the uploaded syllabus.
- Live question-by-question sorting is retained.
- Original MCQs remain intact with the complete stem and all options.
- AI answers are generated separately when requested.
- Figure detection, automatic figure preview, manual crop, Bloom taxonomy, chapter weightage and Word export are retained.

## Run
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Required uploads
1. Syllabus PDF
2. Question Paper PDF

The syllabus should match the Board/Class/Subject selected in the app.
