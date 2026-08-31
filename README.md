# AI Question Paper Analyzer V9

Syllabus-driven ICSE/ISC prototype.

Initial supported configurations:
- ICSE Class X Physics, 2026
- ISC Class XII Physics, 2027

The syllabus is stored as data and selected by Board → Examination Year →
Class → Subject. The app maps uploaded questions to those official syllabus
units, calculates weightage, identifies focus concepts, provides AI answers,
and exports a sorted Word document.

API key belongs in Streamlit Secrets:
OPENAI_API_KEY = "..."
OPENAI_MODEL = "gpt-5.6-luna"
