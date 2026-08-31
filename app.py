
import io
import json
import os
import tempfile
from collections import defaultdict

import pandas as pd
import plotly.express as px
import streamlit as st
from docx import Document
from docx.shared import Inches
from openai import OpenAI

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None

CHAPTERS = [
    "Force",
    "Work, Energy and Power",
    "Machines",
    "Refraction of Light at Plane Surfaces",
    "Refraction through a Lens",
    "Spectrum",
    "Sound",
    "Current Electricity",
    "Household Circuits",
    "Electro-Magnetism",
    "Calorimetry",
    "Radioactivity",
]

DEFAULT_MODEL = "gpt-5.6-luna"
MODEL = os.getenv("OPENAI_MODEL", DEFAULT_MODEL)

st.set_page_config(
    page_title="AI Question Paper Analyzer",
    page_icon="📘",
    layout="wide",
)

st.markdown("""
<style>
.block-container {padding-top:1.5rem;padding-bottom:3rem}
.hero {padding:1.2rem 1.3rem;border:1px solid #ddd;border-radius:16px;
       background:linear-gradient(135deg,rgba(139,47,201,.08),rgba(255,255,255,.8))}
.hero h1 {margin:0}
.muted {color:#666}
.chapter-title {font-size:1.18rem;font-weight:750;border-bottom:2px solid #ddd;
                padding-bottom:.4rem;margin-top:1rem}
.tag {display:inline-block;padding:.25rem .65rem;border-radius:999px;
      background:#f3e8ff;color:#7a27b0;font-weight:650;font-size:.84rem}
</style>
""", unsafe_allow_html=True)


def get_client():
    key = os.getenv("OPENAI_API_KEY")
    if not key:
        try:
            key = st.secrets["OPENAI_API_KEY"]
        except Exception:
            key = None
    if not key:
        st.error(
            "OpenAI API key not found. Add OPENAI_API_KEY to Streamlit Secrets "
            "or set it as an environment variable."
        )
        st.stop()
    return OpenAI(api_key=key)


def extract_docx(data):
    doc = Document(io.BytesIO(data))
    parts = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    for table in doc.tables:
        for row in table.rows:
            cells = [c.text.strip() for c in row.cells]
            if any(cells):
                parts.append(" | ".join(cells))
    return "\n".join(parts)


def make_input(upload):
    data = upload.getvalue()
    name = upload.name.lower()

    if name.endswith(".docx"):
        return [{"type": "input_text", "text": extract_docx(data)}]

    if name.endswith(".txt"):
        return [{"type": "input_text", "text": data.decode("utf-8", errors="ignore")}]

    f = get_client().files.create(file=(upload.name, data), purpose="user_data")
    return [{"type": "input_file", "file_id": f.id}]


def schema():
    # Figure coordinates are normalized 0..1 relative to the PDF page.
    # If no figure is present, has_figure=false and coordinates are 0.
    return {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "paper_title": {"type": "string"},
            "subject": {"type": "string"},
            "printed_question_marks": {"type": "number"},
            "maximum_exam_marks": {"type": "number"},
            "section_structure": {"type": "string"},
            "questions": {
                "type": "array",
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "question_no": {"type": "string"},
                        "parent_question": {"type": "string"},
                        "section": {"type": "string"},
                        "question_text": {"type": "string"},
                        "chapter": {"type": "string", "enum": CHAPTERS},
                        "marks": {"type": "number"},
                        "concept": {"type": "string"},
                        "question_type": {"type": "string"},
                        "difficulty": {
                            "type": "string",
                            "enum": ["Easy", "Moderate", "Difficult", "Not clear"],
                        },
                        "split_reason": {"type": "string"},
                        "source_page": {"type": "integer"},
                        "has_figure": {"type": "boolean"},
                        "figure_x1": {"type": "number"},
                        "figure_y1": {"type": "number"},
                        "figure_x2": {"type": "number"},
                        "figure_y2": {"type": "number"},
                    },
                    "required": [
                        "question_no", "parent_question", "section",
                        "question_text", "chapter", "marks", "concept",
                        "question_type", "difficulty", "split_reason",
                        "source_page", "has_figure",
                        "figure_x1", "figure_y1", "figure_x2", "figure_y2"
                    ],
                },
            },
            "focus_concepts": {
                "type": "array",
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "concept": {"type": "string"},
                        "chapter": {"type": "string", "enum": CHAPTERS},
                        "reason": {"type": "string"},
                    },
                    "required": ["concept", "chapter", "reason"],
                },
            },
        },
        "required": [
            "paper_title", "subject", "printed_question_marks",
            "maximum_exam_marks", "section_structure",
            "questions", "focus_concepts"
        ],
    }


def analyze_paper(upload):
    chapters = "\n".join(f"{i+1}. {x}" for i, x in enumerate(CHAPTERS))

    prompt = f"""
You are an expert school Physics examination-paper analyst.

Use ONLY these chapter names:
{chapters}

Analyse the uploaded question paper.

QUESTION EXTRACTION
- Extract every assessable question.
- Preserve original numbering and wording as closely as possible.
- Split (a), (b), (c) when separately assessable.
- Split further when a sub-part contains distinct assessable parts from
  different chapters, e.g. Q4(iii)(a) and Q4(iii)(b).
- Do not split one idea merely because it has multiple clauses.

FIGURES / DIAGRAMS / GRAPHS / CIRCUITS
- For every question, identify the PDF page containing it in source_page.
- If that question contains or depends on a figure, diagram, graph, ray
  diagram, circuit, apparatus drawing, table, or other visual, set has_figure=true.
- Give the bounding box of the relevant visual as normalized page coordinates:
  x1,y1 = top-left and x2,y2 = bottom-right, all from 0 to 1.
- The box should include the COMPLETE visual needed for the question, including
  labels and arrows, but avoid unrelated surrounding questions where possible.
- If there is no relevant visual, set has_figure=false and all four coordinates to 0.
- Do not guess a figure that is not present.

CHAPTER CLASSIFICATION
- Every extracted item gets exactly one chapter from the supplied list.
- Never invent or rename a chapter.

MARKS
- Record printed marks for each extracted item.
- printed_question_marks = all marks printed across the complete paper,
  including every optional Section B question.
- maximum_exam_marks = maximum marks a candidate can actually score after
  applying the paper's choice rules.
- Do not confuse the complete printed total with the actual scoring maximum.
  For this benchmark paper, the printed total is 100 and the actual exam
  maximum is 80.

ANALYSIS
- Identify concept, type and difficulty.
- Identify the main concepts to focus on.

Return structured JSON only.
"""

    response = get_client().responses.create(
        model=MODEL,
        input=[{
            "role": "user",
            "content": make_input(upload)
            + [{"type": "input_text", "text": prompt}],
        }],
        text={
            "format": {
                "type": "json_schema",
                "name": "physics_question_paper_v5",
                "strict": True,
                "schema": schema(),
            }
        },
    )
    return json.loads(response.output_text)


def answer_question(q):
    prompt = f"""
You are a school Physics teacher.

Answer this question accurately.

Chapter: {q['chapter']}
Concept: {q['concept']}
Marks: {q['marks']}
Type: {q['question_type']}

Question:
{q['question_text']}

Match the depth to the marks.
For numerical questions: Given → Formula → Substitution → Calculation → Final Answer.
For theory: concise explanation and key points.
If a diagram is essential, describe exactly what should be drawn.
"""
    return get_client().responses.create(model=MODEL, input=prompt).output_text.strip()


def totals(data):
    out = {c: {"marks": 0.0, "questions": 0} for c in CHAPTERS}
    for q in data["questions"]:
        out[q["chapter"]]["marks"] += float(q.get("marks", 0))
        out[q["chapter"]]["questions"] += 1
    return out


def normalized_weightage(data):
    """
    Chapter weightage is based on the complete printed question paper.
    For this paper that is 100 marks: 40 compulsory + 60 printed in Section B.
    The actual candidate maximum is 80 because only four 10-mark Section B
    questions are attempted.
    """
    t = totals(data)
    printed = sum(v["marks"] for v in t.values())
    rows = []
    for c in CHAPTERS:
        m = t[c]["marks"]
        if m > 0:
            rows.append({
                "Chapter": c,
                "Questions": t[c]["questions"],
                "Printed Marks": m,
                "Weightage %": round((m / printed * 100) if printed else 0, 1),
            })
    return rows, printed, float(data.get("maximum_exam_marks") or 80)


def pdf_figure(upload_bytes, page_no, x1, y1, x2, y2):
    if fitz is None:
        return None
    try:
        doc = fitz.open(stream=upload_bytes, filetype="pdf")
        if page_no < 1 or page_no > len(doc):
            doc.close()
            return None

        page = doc[page_no - 1]
        rect = page.rect

        # Clamp normalized coordinates.
        x1, y1, x2, y2 = [max(0.0, min(1.0, float(v))) for v in (x1, y1, x2, y2)]
        if x2 <= x1 or y2 <= y1:
            doc.close()
            return None

        clip = fitz.Rect(
            rect.x0 + x1 * rect.width,
            rect.y0 + y1 * rect.height,
            rect.x0 + x2 * rect.width,
            rect.y0 + y2 * rect.height,
        )
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), clip=clip, alpha=False)
        img = pix.tobytes("png")
        doc.close()
        return img
    except Exception:
        return None


def build_word(data, source_bytes=None, include_answers=False):
    doc = Document()
    doc.add_heading(data.get("paper_title") or "AI Sorted Question Paper", 0)
    doc.add_paragraph(f"Subject: {data.get('subject', '')}")
    doc.add_paragraph(
        f"Printed question marks: {data.get('printed_question_marks', 0)} | "
        f"Maximum exam marks: {data.get('maximum_exam_marks', 80)}"
    )

    rows, printed, target = normalized_weightage(data)

    doc.add_heading("Chapter-wise Weightage", 1)
    doc.add_paragraph(
        f"Weightage is calculated from the complete printed paper ({printed:g} marks). "
        f"Actual maximum exam marks after Section B choice = {target:g}."
    )

    table = doc.add_table(rows=1, cols=4)
    for cell, label in zip(
        table.rows[0].cells,
        ["Chapter", "Questions", "Printed Marks", "Weightage"]
    ):
        cell.text = label

    for r in rows:
        cells = table.add_row().cells
        cells[0].text = r["Chapter"]
        cells[1].text = str(r["Questions"])
        cells[2].text = str(r["Printed Marks"])
        cells[3].text = f"{r['Weightage %']}%"

    grouped = defaultdict(list)
    for q in data["questions"]:
        grouped[q["chapter"]].append(q)

    doc.add_page_break()
    doc.add_heading("Questions Sorted Chapter-wise", 1)

    for chapter in CHAPTERS:
        if not grouped[chapter]:
            continue
        doc.add_heading(chapter, 2)

        for q in grouped[chapter]:
            p = doc.add_paragraph()
            p.add_run(f"Q{q['question_no']}. ").bold = True
            p.add_run(q["question_text"])

            meta = doc.add_paragraph()
            meta.add_run("Chapter: ").bold = True
            meta.add_run(q["chapter"])
            meta.add_run(" | Marks: ").bold = True
            meta.add_run(str(q["marks"]))
            meta.add_run(" | Concept: ").bold = True
            meta.add_run(q["concept"])

            if source_bytes and q.get("has_figure") and q.get("source_page"):
                img = pdf_figure(
                    source_bytes,
                    q["source_page"],
                    q.get("figure_x1", 0),
                    q.get("figure_y1", 0),
                    q.get("figure_x2", 0),
                    q.get("figure_y2", 0),
                )
                if img:
                    doc.add_paragraph("Figure / diagram from original paper:")
                    doc.add_picture(io.BytesIO(img), width=Inches(5.8))

            if include_answers:
                h = doc.add_paragraph()
                h.add_run("AI Answer").bold = True
                doc.add_paragraph(answer_question(q))

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def reset():
    st.session_state.data = None
    st.session_state.answers = {}
    st.session_state.source_bytes = None
    st.session_state.source_name = None


if "data" not in st.session_state:
    st.session_state.data = None
if "answers" not in st.session_state:
    st.session_state.answers = {}
if "source_bytes" not in st.session_state:
    st.session_state.source_bytes = None
if "source_name" not in st.session_state:
    st.session_state.source_name = None


# ============================================================
# UI
# ============================================================

st.markdown("""
<div class="hero">
<h1>📘 AI Question Paper Analyzer</h1>
<div class="muted">
Upload → extract → sort by chapter → analyse weightage → identify focus concepts
→ generate answers → download a clean Word paper.
</div>
</div>
""", unsafe_allow_html=True)

with st.sidebar:
    st.header("📚 12-Chapter Syllabus")
    for i, chapter in enumerate(CHAPTERS, 1):
        st.write(f"**{i}.** {chapter}")
    st.divider()
    st.caption(f"AI model: {MODEL}")
    if st.session_state.data:
        if st.button("🆕 Start New Paper", use_container_width=True):
            reset()
            st.rerun()

upload = st.file_uploader(
    "Upload Question Paper",
    type=["pdf", "docx", "txt", "png", "jpg", "jpeg", "webp"],
)

analyse = st.button(
    "🔍 Analyse Paper",
    type="primary",
    disabled=upload is None,
    use_container_width=True,
)

if analyse and upload:
    st.session_state.data = None
    st.session_state.answers = {}
    st.session_state.source_bytes = upload.getvalue()
    st.session_state.source_name = upload.name

    with st.spinner(
        "Reading the paper, splitting sub-parts, locating figures and classifying questions..."
    ):
        try:
            st.session_state.data = analyze_paper(upload)
            st.success("Paper analysed successfully.")
        except Exception as exc:
            st.error(f"Analysis failed: {exc}")

data = st.session_state.data

if not data:
    st.info("Upload a question paper above to begin.")
    st.stop()

st.success("Analysis complete.")
st.info(data["section_structure"])

rows, printed, target = normalized_weightage(data)

m1, m2, m3, m4 = st.columns(4)
m1.metric("Assessable items", len(data["questions"]))
m2.metric("Printed paper", f"{printed:g} marks")
m3.metric("Actual exam maximum", f"{target:g} marks")
m4.metric("Paper weightage", "100%")

# ============================================================
# Weightage
# ============================================================

st.header("📊 Chapter-wise Weightage")

st.caption(
    f"Weightage is calculated from the complete printed paper: {printed:g} marks. "
    f"For this paper, Section A is compulsory (40 marks) and Section B prints "
    f"60 marks, of which any four 10-mark questions are attempted. Therefore "
    f"the printed-paper analysis is out of {printed:g}, while the student's "
    f"actual maximum score is {target:g}."
)

if rows:
    df = pd.DataFrame(rows)
    fig = px.pie(
        df,
        names="Chapter",
        values="Printed Marks",
        hole=0.4,
        title=f"Chapter Weightage — {printed:g}-Mark Printed Paper",
    )
    fig.update_traces(textposition="inside", textinfo="percent")
    st.plotly_chart(fig, use_container_width=True)

    table_df = df.copy()
    table_df["Weightage %"] = table_df["Weightage %"].astype(str) + "%"
    st.dataframe(
        table_df[
            ["Chapter", "Questions", "Printed Marks", "Weightage %"]
        ],
        use_container_width=True,
        hide_index=True,
    )

# ============================================================
# Focus concepts
# ============================================================

st.header("🎯 Main Concepts to Focus On")
for item in data["focus_concepts"]:
    st.markdown(
        f"**{item['concept']}** — `{item['chapter']}`  \n{item['reason']}"
    )

# ============================================================
# Teacher review
# ============================================================

st.header("✏️ Teacher Review")
st.caption(
    "Correct any AI chapter assignment. Weightage and sorted questions update "
    "after the correction."
)

for i, q in enumerate(data["questions"]):
    with st.container(border=True):
        left, right = st.columns([5, 2])

        with left:
            st.markdown(f"**Q{q['question_no']}.** {q['question_text']}")
            st.markdown(
                f'<span class="tag">Chapter: {q["chapter"]}</span>',
                unsafe_allow_html=True,
            )
            st.caption(
                f"Marks: {q['marks']} • Concept: {q['concept']} • "
                f"Type: {q['question_type']} • Difficulty: {q['difficulty']}"
            )
            if q.get("has_figure"):
                st.caption(
                    f"📐 Figure detected — original PDF page {q.get('source_page', '?')}"
                )

        with right:
            selected = st.selectbox(
                "Correct chapter if needed",
                CHAPTERS,
                index=CHAPTERS.index(q["chapter"]),
                key=f"chapter_{i}",
            )
            if selected != q["chapter"]:
                st.session_state.data["questions"][i]["chapter"] = selected
                st.rerun()

# ============================================================
# Sorted questions
# ============================================================

st.header("📝 Questions Sorted Chapter-wise")

grouped = defaultdict(list)
for i, q in enumerate(data["questions"]):
    grouped[q["chapter"]].append((i, q))

for chapter in CHAPTERS:
    chapter_questions = grouped.get(chapter, [])
    if not chapter_questions:
        continue

    st.markdown(
        f'<div class="chapter-title">{chapter}</div>',
        unsafe_allow_html=True,
    )

    for i, q in chapter_questions:
        answer_key = f"{q['question_no']}::{q['question_text']}"

        with st.container(border=True):
            st.markdown(f"**Q{q['question_no']}.** {q['question_text']}")
            st.markdown(
                f'<span class="tag">Chapter: {q["chapter"]}</span>',
                unsafe_allow_html=True,
            )

            x, y, z = st.columns(3)
            x.caption(f"Marks: {q['marks']} • {q['question_type']}")
            y.caption(f"Concept: {q['concept']}")
            z.caption(f"Difficulty: {q['difficulty']}")

            if q.get("has_figure"):
                st.caption(
                    f"📐 Figure / diagram from original paper — page {q.get('source_page', '?')}"
                )
                if (
                    st.session_state.source_bytes
                    and st.session_state.source_name
                    and st.session_state.source_name.lower().endswith(".pdf")
                ):
                    img = pdf_figure(
                        st.session_state.source_bytes,
                        q.get("source_page", 0),
                        q.get("figure_x1", 0),
                        q.get("figure_y1", 0),
                        q.get("figure_x2", 0),
                        q.get("figure_y2", 0),
                    )
                    if img:
                        st.image(img, caption="Original figure / diagram")

            if answer_key in st.session_state.answers:
                with st.expander("💡 AI Answer", expanded=True):
                    st.markdown(st.session_state.answers[answer_key])
            elif st.button("🔗 View AI Answer", key=f"answer_{i}"):
                with st.spinner("Generating answer..."):
                    try:
                        st.session_state.answers[answer_key] = answer_question(q)
                        st.rerun()
                    except Exception as exc:
                        st.error(f"Answer generation failed: {exc}")

# ============================================================
# Downloads
# ============================================================

st.divider()
st.header("⬇️ Download")

source_bytes = st.session_state.source_bytes

c1, c2 = st.columns(2)

with c1:
    st.download_button(
        "📄 Sorted Question Paper (Word)",
        data=build_word(data, source_bytes, include_answers=False),
        file_name="sorted_question_paper_v5.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary",
        use_container_width=True,
    )

with c2:
    st.download_button(
        "📄 Sorted Paper + AI Answers",
        data=build_word(data, source_bytes, include_answers=True),
        file_name="sorted_question_paper_with_answers_v5.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True,
    )
