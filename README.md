# AI Question Paper Analyzer — V18

Streamlit app for analysing CISCE question papers against the official **2027 CISCE syllabus**.

## V18 changes

- **2027 syllabus is always used for classification:** the year of the uploaded paper is now only a reference value. A 2017, 2020, 2026 or 2027 paper is mapped against the official CISCE 2027 syllabus.
- **Paper Year** is retained in the interface for identification and historical analysis, but it does **not** select the syllabus.
- The analysis header clearly shows both the paper year and the syllabus year used.
- **Live sorted paper:** questions are extracted and classified one at a time; each question appears under its chapter as soon as classification is completed.
- **Original MCQs preserved:** the complete MCQ stem and all options remain visible. The analyzer does not replace an MCQ with its answer.
- **AI answers remain separate:** answers are generated only through **View AI Answer** and are not inserted into the sorted question text.
- Retains chapter navigation, Bloom analysis, weightage, figure detection, automatic figure preview, manual crop, and Word export.

## Run

```bash
pip install -r requirements.txt
streamlit run app.py
```

Add `OPENAI_API_KEY` to Streamlit Secrets or the environment.

## Official CISCE 2027 syllabus pages

- ICSE 2027: https://cisce.org/regulations-and-syllabus-icse-2027/
- ISC 2027: https://cisce.org/regulations-and-syllabus-isc-2027/

## Syllabus rule

The analyzer intentionally uses the **2027 syllabus as the classification source of truth for every uploaded paper**, regardless of the paper's year. This allows historical papers to be studied in relation to the current syllabus.
