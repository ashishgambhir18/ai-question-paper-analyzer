
import io
import json
import os
from collections import defaultdict

import pandas as pd
import plotly.express as px
import streamlit as st
from docx import Document
from openai import OpenAI

# ============================================================
# AI QUESTION PAPER ANALYZER — V4
# ============================================================

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
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
.block-container {padding-top: 1.5rem; padding-bottom: 3rem;}
.hero {padding: 1.2rem 1.3rem; border: 1px solid #ddd; border-radius: 16px;
       background: linear-gradient(135deg, rgba(139,47,201,.08), rgba(255,255,255,.75));}
.hero h1 {margin:0;}
.muted {color:#666;}
.chapter-title {font-size:1.18rem;font-weight:750;border-bottom:2px solid #ddd;
                padding-bottom:.4rem;margin-top:1rem;}
.tag {display:inline-block;padding:.25rem .65rem;border-radius:999px;
      background:#f3e8ff;color:#7a27b0;font-weight:650;font-size:.84rem;}
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
    chunks = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    for table in doc.tables:
        for row in table.rows:
            vals = [c.text.strip() for c in row.cells]
            if any(vals):
                chunks.append(" | ".join(vals))
    return "\n".join(chunks)


def make_input(upload):
    data = upload.getvalue()
    name = upload.name.lower()

    if name.endswith(".docx"):
        return [{"type": "input_text", "text": extract_docx(data)}]

    if name.endswith(".txt"):
        return [{"type": "input_text", "text": data.decode("utf-8", errors="ignore")}]

    f = get_client().files.create(
        file=(upload.name, data),
        purpose="user_data",
    )
    return [{"type": "input_file", "file_id": f.id}]


def schema():
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
                    },
                    "required": [
                        "question_no", "parent_question", "section",
                        "question_text", "chapter", "marks", "concept",
                        "question_type", "difficulty", "split_reason"
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

Use ONLY this chapter list:
{chapters}

Analyse the uploaded question paper.

EXTRACTION:
- Extract every assessable question.
- Preserve original numbering and wording as closely as possible.
- Split (a), (b), (c) parts when they are separately assessable.
- Split further when a sub-part contains genuinely distinct assessable
  parts from different chapters. Example: Q4(iii)(a), Q4(iii)(b).
- Do not split a single idea merely because it contains multiple clauses.
- Preserve useful information from diagrams, graphs and circuits in text form
  when possible.

CLASSIFICATION:
- Every extracted item gets exactly ONE chapter from the supplied list.
- Never invent, rename or merge chapters.
- If an item was split, explain the reason briefly in split_reason.

MARKS:
- Record printed marks for each extracted item.
- printed_question_marks = total marks printed in the paper, including
  optional questions.
- maximum_exam_marks = maximum marks a candidate can actually score after
  applying the paper's attempt/choice rules.
- Do not double count optional questions.
- If a mark is genuinely unavailable, use 0.

ANALYSIS:
- Identify the main concept, question type and difficulty.
- Identify the most important concepts to focus on using marks, frequency,
  recurrence and conceptual importance.
- Do not silently alter the source paper.

Return ONLY structured JSON.
"""
    response = get_client().responses.create(
        model=MODEL,
        input=[{
            "role": "user",
            "content": make_input(upload) + [{"type": "input_text", "text": prompt}],
        }],
        text={
            "format": {
                "type": "json_schema",
                "name": "physics_question_paper_v4",
                "strict": True,
                "schema": schema(),
            }
        },
    )
    return json.loads(response.output_text)


def answer_question(q):
    prompt = f"""
You are a school Physics teacher.

Answer the following question accurately.

Chapter: {q['chapter']}
Concept: {q['concept']}
Marks: {q['marks']}
Question type: {q['question_type']}
Question:
{q['question_text']}

Rules:
- Match the depth to the marks.
- Numerical: Given → Formula → Substitution → Calculation → Final Answer.
- Theory: concise explanation and key points.
- Preserve units and equations.
- If a diagram is required, state clearly what should be drawn.
"""
    return get_client().responses.create(model=MODEL, input=prompt).output_text.strip()


def totals_from_questions(data):
    totals = {c: {"marks": 0.0, "questions": 0} for c in CHAPTERS}
    for q in data["questions"]:
        c = q["chapter"]
        totals[c]["marks"] += float(q.get("marks", 0))
        totals[c]["questions"] += 1
    return totals


def build_word(data, include_answers=False):
    doc = Document()
    doc.add_heading(data.get("paper_title") or "AI Sorted Question Paper", 0)
    doc.add_paragraph(f"Subject: {data.get('subject', '')}")
    doc.add_paragraph(
        f"Printed question marks: {data.get('printed_question_marks', 0)} | "
        f"Maximum exam marks: {data.get('maximum_exam_marks', 0)}"
    )

    totals = totals_from_questions(data)
    total_printed = sum(x["marks"] for x in totals.values())

    doc.add_heading("Chapter-wise Weightage", 1)
    table = doc.add_table(rows=1, cols=4)
    for cell, label in zip(
        table.rows[0].cells,
        ["Chapter", "Questions", "Marks", "Weightage"]
    ):
        cell.text = label

    for c in CHAPTERS:
        m = totals[c]["marks"]
        if m <= 0:
            continue
        cells = table.add_row().cells
        cells[0].text = c
        cells[1].text = str(totals[c]["questions"])
        cells[2].text = str(m)
        cells[3].text = f"{m/total_printed*100:.1f}%" if total_printed else "0%"

    grouped = defaultdict(list)
    for q in data["questions"]:
        grouped[q["chapter"]].append(q)

    doc.add_page_break()
    doc.add_heading("Questions Sorted Chapter-wise", 1)

    for c in CHAPTERS:
        if not grouped[c]:
            continue
        doc.add_heading(c, 2)
        for q in grouped[c]:
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


if "data" not in st.session_state:
    st.session_state.data = None
if "answers" not in st.session_state:
    st.session_state.answers = {}

# ============================================================
# Header
# ============================================================

st.markdown("""
<div class="hero">
<h1>📘 AI Question Paper Analyzer</h1>
<div class="muted">
Upload a Physics question paper → classify every question → analyse chapter
weightage → identify focus concepts → generate answers → download Word.
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
    help="Scanned PDFs and images are supported through the model's file input.",
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
    with st.spinner("Reading paper, splitting sub-parts and classifying questions..."):
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

totals = totals_from_questions(data)
questions = data["questions"]
printed = sum(x["marks"] for x in totals.values())
max_exam = float(data["maximum_exam_marks"])

m1, m2, m3, m4 = st.columns(4)
m1.metric("Assessable items", len(questions))
m2.metric("Printed marks", int(printed) if printed.is_integer() else printed)
m3.metric("Maximum exam marks", int(max_exam) if max_exam.is_integer() else max_exam)
m4.metric("Chapters used", sum(1 for c in CHAPTERS if totals[c]["marks"] > 0))

# ============================================================
# Weightage
# ============================================================

st.header("📊 Chapter-wise Weightage")

rows = [
    {
        "Chapter": c,
        "Questions": totals[c]["questions"],
        "Marks": totals[c]["marks"],
    }
    for c in CHAPTERS if totals[c]["marks"] > 0
]

if rows:
    df = pd.DataFrame(rows)
    fig = px.pie(
        df,
        names="Chapter",
        values="Marks",
        hole=0.4,
        title="Printed Question Weightage",
    )
    fig.update_traces(textposition="inside", textinfo="percent")
    st.plotly_chart(fig, use_container_width=True)

    table_df = df.copy()
    table_df["Weightage"] = (
        table_df["Marks"] / table_df["Marks"].sum() * 100
    ).round(1).astype(str) + "%"
    st.dataframe(table_df, use_container_width=True, hide_index=True)

# ============================================================
# Focus concepts
# ============================================================

st.header("🎯 Main Concepts to Focus On")
for item in data["focus_concepts"]:
    st.markdown(
        f"**{item['concept']}** — `{item['chapter']}`  \n"
        f"{item['reason']}"
    )

# ============================================================
# Teacher review / correction
# ============================================================

st.header("✏️ Teacher Review")
st.caption(
    "AI chapter assignments can be corrected here. The chart and Word export "
    "use the corrected assignments."
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
            if q.get("split_reason"):
                st.caption(f"Classification/split note: {q['split_reason']}")

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
# Export
# ============================================================

st.divider()
st.header("⬇️ Download")

c1, c2 = st.columns(2)

with c1:
    st.download_button(
        "📄 Sorted Question Paper (Word)",
        data=build_word(data, include_answers=False),
        file_name="sorted_question_paper_v4.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary",
        use_container_width=True,
    )

with c2:
    st.download_button(
        "📄 Sorted Paper + AI Answers",
        data=build_word(data, include_answers=True),
        file_name="sorted_question_paper_with_answers_v4.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True,
    )
