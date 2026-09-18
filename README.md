# AI Question Paper Analyzer — V20

V20 fixes the CISCE HTTP 403 problem for Physics papers.

## Important behaviour
- Every uploaded paper, regardless of paper year, is analysed against the **CISCE 2027 syllabus**.
- The paper year is reference information only.
- For ICSE Physics, a bundled 2027 syllabus unit map is used if the CISCE website blocks Streamlit server requests. This prevents the app from failing because of HTTP 403 and does not substitute an older syllabus.
- ISC Physics also has a bundled 2027 unit map fallback.
- Live question-by-question sorting is retained.
- Original MCQ stems and all options are retained. AI answers are never inserted into the question text.
- Figure detection/cropping, Bloom analysis, chapter weightage and Word export are retained.
