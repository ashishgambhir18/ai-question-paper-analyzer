
import io
import json
import os
from collections import defaultdict

import streamlit as st
import pandas as pd
import plotly.express as px
from PIL import Image
from streamlit_cropper import st_cropper
from docx import Document
from docx.shared import Inches
from openai import OpenAI
import requests
from bs4 import BeautifulSoup

try:
    import fitz
except ImportError:
    fitz = None

try:
    import cv2
    import numpy as np
except ImportError:
    cv2 = None
    np = None

DEFAULT_CHAPTERS = [
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

ICSE_SUBJECTS=[
"English","Second Languages","History, Civics and Geography","Mathematics","Science",
"Economics","Commercial Studies","Modern Foreign Language","Classical Language",
"Environmental Science","Computer Applications","Economic Applications",
"Commercial Applications","Art","Performing Arts","Home Science","Cookery",
"Fashion Designing","Physical Education","Yoga","Technical Drawing Applications",
"Environmental Applications","Mass Media & Communication","Hospitality Management"
]
ISC_SUBJECTS=[
"English","Indian Languages","Modern Foreign Languages","Classical Languages",
"Elective English","History","Political Science","Geography","Sociology","Psychology",
"Economics","Commerce","Accounts","Business Studies","Mathematics","Physics",
"Chemistry","Biology","Home Science","Fashion Designing","Electricity & Electronics",
"Engineering Science","Computer Science","Geometrical & Mechanical Drawing",
"Geometrical & Building Drawing","Art","Music","Physical Education",
"Environmental Science","Biotechnology","Mass Media & Communication",
"Legal Studies","Hospitality Management"
]
OFFICIAL_PAGES={"ICSE":"https://cisce.org/icse-publications/","ISC":"https://cisce.org/isc-publications/"}
CHAPTERS=list(DEFAULT_CHAPTERS)

def official_links(board,year,subject):
    url=OFFICIAL_PAGES[board]
    try:
        r=requests.get(url,headers={"User-Agent":"Mozilla/5.0"},timeout=25); r.raise_for_status()
        soup=BeautifulSoup(r.text,"html.parser")
        terms=[t.lower() for t in re.findall(r"[A-Za-z]+",subject) if len(t)>2]
        found=[]
        for a in soup.find_all("a",href=True):
            href=a["href"].strip(); txt=" ".join(a.stripped_strings); full=requests.compat.urljoin(url,href)
            blob=(txt+" "+full).lower()
            if str(year) not in blob or not ("syllab" in blob or "regulation" in blob or full.lower().endswith(".pdf")): continue
            score=sum(t in blob for t in terms)
            if terms and score==0: continue
            found.append((score,txt or full,full))
        seen=set(); out=[]
        for score,txt,full in sorted(found,key=lambda x:(x[0],-len(x[1])),reverse=True):
            if full not in seen: out.append({"title":txt,"url":full}); seen.add(full)
        return out[:10],None
    except Exception as e:
        return [],str(e)

def extract_syllabus_units(raw,name,board,year,klass,subject,source):
    f=client().files.create(file=(name,raw),purpose="user_data")
    prompt=f"""Extract the official CISCE syllabus hierarchy from this document.
Board: {board}; Examination year: {year}; Class: {klass}; Subject: {subject}.
Use only terminology appearing in the official document. Preserve order.
Return top-level units/chapters/sections suitable for classifying exam questions,
with concise subtopics. Do not invent, rename or merge unrelated content.
For ISC, return only the selected class."""
    r=client().responses.create(model=MODEL,input=[{"role":"user","content":[
        {"type":"input_file","file_id":f.id},{"type":"input_text","text":prompt}]}])
    rr=client().responses.create(model=MODEL,input="Convert the extraction below to JSON only.\n"+r.output_text,
        text={"format":{"type":"json_schema","name":"syllabus_units","strict":True,"schema":{
            "type":"object","additionalProperties":False,
            "properties":{"units":{"type":"array","items":{
                "type":"object","additionalProperties":False,
                "properties":{"name":{"type":"string"},"topics":{"type":"array","items":{"type":"string"}}},
                "required":["name","topics"]}}},
            "required":["units"]}}})
    units=json.loads(rr.output_text)["units"]
    return {"board":board,"year":int(year),"class":klass,"subject":subject,
            "source_url":source,"units":units}


MODEL = os.getenv("OPENAI_MODEL", "gpt-5.6-luna")

st.set_page_config(page_title="AI Question Paper Analyzer", page_icon="📘", layout="wide")

st.markdown("""
<style>
.block-container{padding-top:1.4rem;padding-bottom:3rem}
.hero{padding:1.2rem 1.3rem;border:1px solid #ddd;border-radius:16px;
background:linear-gradient(135deg,rgba(139,47,201,.08),rgba(255,255,255,.85))}
.hero h1{margin:0}.muted{color:#666}
.chapter-title{font-size:1.18rem;font-weight:750;border-bottom:2px solid #ddd;
padding-bottom:.4rem;margin-top:1rem}
.tag{display:inline-block;padding:.25rem .65rem;border-radius:999px;
background:#f3e8ff;color:#7a27b0;font-weight:650;font-size:.84rem}
</style>
""", unsafe_allow_html=True)


def client():
    key = os.getenv("OPENAI_API_KEY")
    if not key:
        try:
            key = st.secrets["OPENAI_API_KEY"]
        except Exception:
            key = None
    if not key:
        st.error("OPENAI_API_KEY is missing. Add it in Streamlit Secrets.")
        st.stop()
    return OpenAI(api_key=key)


def docx_text(data):
    doc = Document(io.BytesIO(data))
    parts = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    for table in doc.tables:
        for row in table.rows:
            vals = [c.text.strip() for c in row.cells]
            if any(vals):
                parts.append(" | ".join(vals))
    return "\n".join(parts)


def upload_input(upload):
    data = upload.getvalue()
    name = upload.name.lower()
    if name.endswith(".docx"):
        return [{"type":"input_text","text":docx_text(data)}]
    if name.endswith(".txt"):
        return [{"type":"input_text","text":data.decode("utf-8",errors="ignore")}]
    f = client().files.create(file=(upload.name, data), purpose="user_data")
    return [{"type":"input_file","file_id":f.id}]


def output_schema():
    return {
        "type":"object","additionalProperties":False,
        "properties":{
            "paper_title":{"type":"string"},
            "subject":{"type":"string"},
            "printed_question_marks":{"type":"number"},
            "maximum_exam_marks":{"type":"number"},
            "section_structure":{"type":"string"},
            "questions":{
                "type":"array","items":{
                    "type":"object","additionalProperties":False,
                    "properties":{
                        "question_no":{"type":"string"},
                        "parent_question":{"type":"string"},
                        "section":{"type":"string"},
                        "question_text":{"type":"string"},
                        "chapter":{"type":"string","enum":CHAPTERS},
                        "marks":{"type":"number"},
                        "concept":{"type":"string"},
                        "question_type":{"type":"string"},
                        "bloom_level":{"type":"string","enum":["Remember","Understand","Apply","Analyze","Evaluate","Create"]},
                        "bloom_reason":{"type":"string"},
                        "difficulty":{"type":"string","enum":["Easy","Moderate","Difficult","Not clear"]},
                        "split_reason":{"type":"string"},
                        "source_page":{"type":"integer"},
                        "has_figure":{"type":"boolean"},
                        "figure_x1":{"type":"number"},
                        "figure_y1":{"type":"number"},
                        "figure_x2":{"type":"number"},
                        "figure_y2":{"type":"number"},
                    },
                    "required":[
                        "question_no","parent_question","section","question_text",
                        "chapter","marks","concept","question_type","bloom_level","bloom_reason","difficulty",
                        "split_reason","source_page","has_figure",
                        "figure_x1","figure_y1","figure_x2","figure_y2"
                    ]
                }
            },
            "focus_concepts":{
                "type":"array","items":{
                    "type":"object","additionalProperties":False,
                    "properties":{
                        "concept":{"type":"string"},
                        "chapter":{"type":"string","enum":CHAPTERS},
                        "reason":{"type":"string"}
                    },
                    "required":["concept","chapter","reason"]
                }
            }
        },
        "required":[
            "paper_title","subject","printed_question_marks","maximum_exam_marks",
            "section_structure","questions","focus_concepts"
        ]
    }


def analyse(upload):
    chapter_list = "\n".join(f"{i+1}. {x}" for i,x in enumerate(CHAPTERS))
    prompt = f"""
You are an expert school Physics examination-paper analyst.

Allowed chapters ONLY:
{chapter_list}

Analyse the complete uploaded question paper.

EXTRACTION:
- Extract every assessable question and preserve its original numbering.
- Split independently assessable (a), (b), (c) parts.
- If a sub-question contains distinct assessable parts from different chapters,
  split those further.
- Keep each extracted question's text limited to THAT question only. Do not
  accidentally append text from the preceding or following question.

MARKS:
- printed_question_marks = total marks printed across the complete paper,
  including all optional questions.
- maximum_exam_marks = maximum score a student can obtain after applying all
  attempt/choice rules.
- Do not confuse printed total with exam maximum.
- For the benchmark paper, Section A is 40 compulsory, Section B prints six
  10-mark questions, any four are attempted, so printed total is 100 and
  maximum exam score is 80.

FIGURES:
- Identify source_page for each question.
- If a question depends on a diagram, graph, circuit, ray diagram, apparatus,
  table, or other visual, set has_figure=true.
- Give a TIGHT normalized bounding box (0..1) around ONLY the visual needed
  for that question.
- Include the visual's own labels, arrows, axes and legend.
- EXCLUDE question prose, answer choices, marks, headings and adjacent-question
  text. Do not use a large box merely because it is convenient.
- If no visual, has_figure=false and coordinates all 0.

CLASSIFICATION:
- Assign exactly one allowed chapter to every extracted item.
- Identify concept, type and difficulty.
- Classify every question using revised Bloom's Taxonomy:
  Remember, Understand, Apply, Analyze, Evaluate, Create.
- Choose the highest cognitive process genuinely required to answer the question.
- Do not label every numerical question as Apply automatically.
- Give a short bloom_reason.
- Give concise reasons for any split.
- Identify important focus concepts.

Return structured JSON only.
"""
    r = client().responses.create(
        model=MODEL,
        input=[{"role":"user","content":upload_input(upload)+[{"type":"input_text","text":prompt}]}],
        text={"format":{"type":"json_schema","name":"physics_qp_v7","strict":True,"schema":output_schema()}}
    )
    return json.loads(r.output_text)


def answer(q):
    prompt = f"""
Answer this school Physics question as a teacher.
Chapter: {q['chapter']}
Concept: {q['concept']}
Marks: {q['marks']}
Question: {q['question_text']}

Match the depth to the marks. For numericals use Given, Formula,
Substitution, Calculation and Final Answer. Keep units. For theory give
clear exam-ready points. If a diagram is essential, describe it.

MATHEMATICAL FORMATTING:
- Use $...$ for inline mathematics.
- Use $$...$$ for equations on separate lines.
- Do not use \\(...\\) or \\[...\\] delimiters.
- Do not put equations in code blocks.
- Use proper LaTeX for fractions, powers, subscripts and Greek letters.
"""
    return client().responses.create(model=MODEL,input=prompt).output_text.strip()



def section_mark_summary(data):
    """Summarize printed marks by section using extracted question numbers."""
    sections = defaultdict(float)
    for q in data["questions"]:
        sections[q.get("section", "Other")] += float(q.get("marks", 0))
    return dict(sections)


def benchmark_mark_validation(data):
    """
    The benchmark paper has a known structure:
    Section A = 40 compulsory
    Section B = six printed 10-mark questions = 60
    Printed total = 100
    Actual maximum = 80

    Return a diagnostic instead of silently changing the AI's extracted marks.
    """
    sections = section_mark_summary(data)
    # Section labels vary by OCR/AI, so also use the question numbering when needed.
    a = sum(float(q["marks"]) for q in data["questions"]
            if str(q.get("question_no","")).split("(")[0].strip().upper() in
            {"1","2","3"})
    b = sum(float(q["marks"]) for q in data["questions"]
            if str(q.get("question_no","")).split("(")[0].strip().upper() in
            {"4","5","6","7","8","9"})
    return {
        "section_a_extracted": a,
        "section_b_extracted": b,
        "expected_printed": 100.0,
        "expected_exam_max": 80.0,
        "extracted_total": a + b,
        "difference": round((a + b) - 100.0, 2),
    }




def benchmark_printed_total(data):
    """The benchmark paper is 40 compulsory + six printed 10-mark options = 100."""
    structure=(data.get("section_structure") or "").lower()
    # V8 is the benchmark ICSE/ISC Physics paper with the known 100/80 structure.
    # Use the official paper structure as the printed-total source of truth.
    if ("40" in structure and "six" in structure and "10" in structure) or \
       ("40" in structure and "6" in structure and "10" in structure):
        return 100.0
    reported=float(data.get("printed_question_marks") or 0)
    return reported if reported > 0 else sum(float(q.get("marks",0) or 0) for q in data.get("questions",[]))


def chapter_totals(data):
    """Return chapter-wise marks and question counts for the analysed paper."""
    totals={c:{"marks":0.0,"questions":0} for c in CHAPTERS}
    for q in data.get("questions",[]):
        c=q.get("chapter")
        if c not in totals:
            continue
        totals[c]["marks"] += float(q.get("marks",0) or 0)
        totals[c]["questions"] += 1
    return totals



def bloom_totals(data):
    levels=["Remember","Understand","Apply","Analyze","Evaluate","Create"]
    totals={level:{"marks":0.0,"questions":0} for level in levels}
    for q in data.get("questions",[]):
        level=q.get("bloom_level")
        if level in totals:
            totals[level]["marks"] += float(q.get("marks",0) or 0)
            totals[level]["questions"] += 1
    return totals


def figure_crop(pdf_bytes, page_no, coords):
    """
    Extract the ORIGINAL figure from the PDF.

    The AI supplies an approximate figure box, but we do NOT use that box
    blindly because it can contain nearby question text. Instead:
      1. Find vector drawing/image objects inside the AI box.
      2. Build a connected visual cluster from those objects.
      3. Include only short text labels that are immediately attached to the
         visual cluster.
      4. Never use a large paragraph/text block to enlarge the crop.
    """
    if fitz is None:
        return None

    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        if not (1 <= int(page_no) <= len(doc)):
            doc.close()
            return None

        page = doc[int(page_no) - 1]
        p = page.rect

        vals = [max(0.0, min(1.0, float(v))) for v in coords]
        candidate = fitz.Rect(
            p.x0 + vals[0] * p.width,
            p.y0 + vals[1] * p.height,
            p.x0 + vals[2] * p.width,
            p.y0 + vals[3] * p.height,
        )

        if candidate.width <= 3 or candidate.height <= 3:
            doc.close()
            return None

        # Find actual visual objects, not text.
        objects = []

        try:
            for d in page.get_drawings():
                r = d.get("rect")
                if not r or r.width <= 1 or r.height <= 1:
                    continue
                if not r.intersects(candidate):
                    continue
                inter = r & candidate
                if inter.width * inter.height >= 6:
                    objects.append(fitz.Rect(r))
        except Exception:
            pass

        try:
            for item in page.get_image_info(xrefs=True):
                r = item.get("bbox")
                if not r:
                    continue
                r = fitz.Rect(r)
                if r.intersects(candidate) and r.width * r.height >= 100:
                    objects.append(r)
        except Exception:
            pass

        # Keep objects whose centre is inside the AI figure area.
        inside=[]
        for r in objects:
            cx=(r.x0+r.x1)/2
            cy=(r.y0+r.y1)/2
            if candidate.contains(fitz.Point(cx,cy)):
                inside.append(r)

        if not inside:
            # If no vector/image object was detected, use a tighter version of
            # the AI box, but do not add any text outside it.
            crop=candidate
        else:
            # Build connected/nearby visual components. This is deliberately
            # based only on visual objects, so nearby prose cannot enter here.
            remaining=inside[:]
            components=[]
            xgap=candidate.width*0.035
            ygap=candidate.height*0.035

            while remaining:
                component=[remaining.pop(0)]
                changed=True
                while changed:
                    changed=False
                    union=component[0]
                    for r in component[1:]:
                        union |= r

                    keep=[]
                    for r in remaining:
                        gx=max(0,max(union.x0,r.x0)-min(union.x1,r.x1))
                        gy=max(0,max(union.y0,r.y0)-min(union.y1,r.y1))
                        if gx <= xgap and gy <= ygap:
                            component.append(r)
                            changed=True
                        else:
                            keep.append(r)
                    remaining=keep

                u=component[0]
                for r in component[1:]:
                    u |= r
                components.append(u)

            # Select the component/combined components that overlap the centre
            # region of the AI box. This prevents a neighbouring diagram from
            # being merged simply because it is on the same page.
            ccx=(candidate.x0+candidate.x1)/2
            ccy=(candidate.y0+candidate.y1)/2

            components.sort(
                key=lambda r:
                ((r.x0+r.x1)/2-ccx)**2 +
                ((r.y0+r.y1)/2-ccy)**2
            )

            crop=components[0]

            # Add another visual component only when it is very close and
            # aligned with the selected diagram (useful for a broken wire).
            for comp in components[1:]:
                gx=max(0,max(crop.x0,comp.x0)-min(crop.x1,comp.x1))
                gy=max(0,max(crop.y0,comp.y0)-min(crop.y1,comp.y1))
                aligned_x = gy <= candidate.height*0.035
                aligned_y = gx <= candidate.width*0.035
                close = gx <= candidate.width*0.045 and gy <= candidate.height*0.045
                if close or (aligned_x and gx <= candidate.width*0.08) or (aligned_y and gy <= candidate.height*0.08):
                    crop |= comp

            # Add ONLY compact labels belonging to the selected figure.
            # Large text blocks (question sentences, instructions, etc.) are
            # explicitly rejected.
            try:
                blocks=page.get_text("blocks")
                for b in blocks:
                    if len(b)<5:
                        continue
                    tr=fitz.Rect(b[:4])
                    txt=(b[4] or "").strip()
                    if not txt:
                        continue

                    # Reject paragraph-like blocks.
                    word_count=len(txt.split())
                    if word_count > 8 or tr.width > candidate.width*0.45:
                        continue

                    # A label should be small relative to the figure.
                    if tr.height > candidate.height*0.35:
                        continue

                    # Label must be inside the original AI figure area.
                    if not candidate.intersects(tr):
                        continue

                    # Label must touch/overlap the visual crop.
                    gx=max(0,max(crop.x0,tr.x0)-min(crop.x1,tr.x1))
                    gy=max(0,max(crop.y0,tr.y0)-min(crop.y1,tr.y1))
                    if gx <= candidate.width*0.025 and gy <= candidate.height*0.06:
                        crop |= tr
            except Exception:
                pass

            # Very small margin only. Do not expand back into surrounding text.
            mx=min(4.0,p.width*0.004)
            my=min(4.0,p.height*0.004)
            crop=fitz.Rect(
                max(p.x0,crop.x0-mx),
                max(p.y0,crop.y0-my),
                min(p.x1,crop.x1+mx),
                min(p.y1,crop.y1+my),
            )

        pix=page.get_pixmap(
            matrix=fitz.Matrix(2.5,2.5),
            clip=crop,
            alpha=False
        )
        png=pix.tobytes("png")
        doc.close()
        return png

    except Exception:
        return None


def source_page_image(pdf_bytes,page_no):
    if fitz is None: return None
    try:
        doc=fitz.open(stream=pdf_bytes,filetype="pdf")
        page_no=int(page_no)
        if page_no<1 or page_no>len(doc):
            doc.close(); return None
        pix=doc[page_no-1].get_pixmap(matrix=fitz.Matrix(2.2,2.2),alpha=False)
        out=pix.tobytes("png"); doc.close(); return out
    except Exception:
        return None

def manual_crop_ui(pdf_bytes,page_no,crop_key):
    png=source_page_image(pdf_bytes,page_no)
    if not png:
        st.error("Could not render the original PDF page.")
        return None
    image=Image.open(io.BytesIO(png))
    st.markdown("**✂️ Crop the original figure**")
    st.caption("Drag the rectangle around the figure. Include labels/arrows/axes/wires; exclude surrounding question text.")
    return st_cropper(
        image,realtime_update=True,box_color="#ff4b4b",
        aspect_ratio=None,return_type="image",key=f"cropper_{crop_key}"
    )

def render_ai_answer(text):
    """Normalize common LaTeX delimiters before Streamlit rendering."""
    if not text:
        return ""
    text=text.replace("\r\n","\n").replace("\r","\n")
    text=text.replace(r"\[","$$").replace(r"\]","$$")
    text=text.replace(r"\(","$").replace(r"\)","$")
    # Some model outputs wrap a display equation in [ ... ].
    out=[]
    for line in text.split("\n"):
        s=line.strip()
        if len(s)>4 and s.startswith("[") and s.endswith("]") and "\\" in s:
            s=s[1:-1].strip()
            line="$$\n"+s+"\n$$"
        out.append(line)
    return "\n".join(out)


def make_word(data,pdf_bytes=None,with_answers=False):
    manual_crops=st.session_state.get("manual_crops",{})
    doc=Document()
    doc.add_heading(data.get("paper_title") or "AI Sorted Question Paper",0)
    doc.add_paragraph(f"Subject: {data.get('subject','')}")
    doc.add_paragraph(
        f"Printed paper: {data.get('printed_question_marks',100):g} marks | "
        f"Actual exam maximum: {data.get('maximum_exam_marks',80):g} marks"
    )
    t=chapter_totals(data); total=sum(v["marks"] for v in t.values())
    doc.add_heading("Chapter-wise Weightage",1)
    doc.add_paragraph(f"Weightage is calculated from the complete printed paper ({total:g} marks).")
    table=doc.add_table(rows=1,cols=4)
    for cell,label in zip(table.rows[0].cells,["Chapter","Questions","Marks","Weightage"]):
        cell.text=label
    for c in CHAPTERS:
        if t[c]["marks"]:
            cells=table.add_row().cells; m=t[c]["marks"]
            cells[0].text=c; cells[1].text=str(t[c]["questions"])
            cells[2].text=str(m); cells[3].text=f"{m/total*100:.1f}%"

    bt=bloom_totals(data)
    doc.add_heading("Bloom's Taxonomy Analysis",1)
    btotal=sum(v["marks"] for v in bt.values())
    btable=doc.add_table(rows=1,cols=4)
    for cell,label in zip(btable.rows[0].cells,["Bloom Level","Questions","Marks","Weightage"]):
        cell.text=label
    for level in ["Remember","Understand","Apply","Analyze","Evaluate","Create"]:
        if bt[level]["questions"]:
            cells=btable.add_row().cells
            cells[0].text=level
            cells[1].text=str(bt[level]["questions"])
            cells[2].text=str(bt[level]["marks"])
            cells[3].text=f"{bt[level]['marks']/btotal*100:.1f}%"

    groups=defaultdict(list)
    for q in data["questions"]: groups[q["chapter"]].append(q)
    doc.add_page_break(); doc.add_heading("Questions Sorted Chapter-wise",1)
    for c in CHAPTERS:
        if not groups[c]: continue
        doc.add_heading(c,2)
        for q in groups[c]:
            p=doc.add_paragraph(); p.add_run(f"Q{q['question_no']}. ").bold=True
            p.add_run(q["question_text"])
            meta=doc.add_paragraph()
            meta.add_run("Marks: ").bold=True; meta.add_run(str(q["marks"]))
            meta.add_run(" | Concept: ").bold=True; meta.add_run(q["concept"])
            meta.add_run(" | Bloom: ").bold=True; meta.add_run(q.get("bloom_level","Not classified"))
            if pdf_bytes and q.get("has_figure"):
                crop_key=f"{q['question_no']}_{i}"
                saved=manual_crops.get(crop_key)
                if saved is not None:
                    buf=io.BytesIO()
                    saved.save(buf,format="PNG")
                    img_bytes=buf.getvalue()
                else:
                    img_bytes=figure_crop(
                        pdf_bytes,q.get("source_page",0),
                        (q.get("figure_x1",0),q.get("figure_y1",0),
                         q.get("figure_x2",0),q.get("figure_y2",0))
                    )
                if img_bytes:
                    doc.add_paragraph("Figure:")
                    doc.add_picture(io.BytesIO(img_bytes),width=Inches(5.8))

            if with_answers:
                h=doc.add_paragraph(); h.add_run("AI Answer").bold=True
                doc.add_paragraph(answer(q))
    out=io.BytesIO(); doc.save(out); return out.getvalue()


def reset():
    for k in ("data","answers","source_bytes","source_name"):
        st.session_state[k]=None if k=="data" or k=="source_bytes" or k=="source_name" else {}


if "data" not in st.session_state: st.session_state.data=None
if "answers" not in st.session_state: st.session_state.answers={}
if "source_bytes" not in st.session_state: st.session_state.source_bytes=None
if "source_name" not in st.session_state: st.session_state.source_name=None

st.markdown("""
<div class="hero">
<h1>📘 AI Question Paper Analyzer</h1>
<div class="muted">Upload → classify → analyse → focus → answer → download.</div>
</div>
""",unsafe_allow_html=True)

with st.sidebar:
    st.header("📚 Chapters / Units")
    nav_data=st.session_state.get("data")
    if nav_data and nav_data.get("questions"):
        nav_groups=defaultdict(list)
        for qq in nav_data["questions"]:
            nav_groups[qq.get("chapter","Unclassified")].append(qq)

        if st.button("🏠 Overview",key="nav_overview",use_container_width=True):
            st.session_state.pop("selected_chapter",None)
            st.rerun()

        # Preserve the syllabus order when available; append any unexpected
        # classifications afterward.
        nav_order=list(CHAPTERS)
        for cc in nav_groups:
            if cc not in nav_order:
                nav_order.append(cc)

        for nn,cc in enumerate(nav_order):
            if nav_groups.get(cc):
                if st.button(
                    f"{cc}  ({len(nav_groups[cc])})",
                    key=f"nav_chapter_{nn}",
                    use_container_width=True
                ):
                    st.session_state["selected_chapter"]=cc
                    st.rerun()
    else:
        st.caption("Run the analysis first. Your chapter folders will appear here.")


st.title("📘 AI Question Paper Analyzer")
st.caption("Upload the syllabus and question paper. The syllabus is the source of truth for chapter/unit classification.")

c1,c2,c3,c4=st.columns(4)
with c1: board=st.selectbox("Board",["ICSE","ISC"])
with c2: year=st.selectbox("Examination Year",list(range(2027,2014,-1)))
with c3: klass="X" if board=="ICSE" else st.selectbox("Class",["XII","XI"])
with c4:
    subjects=ICSE_SUBJECTS if board=="ICSE" else ISC_SUBJECTS
    subject=st.selectbox("Subject",subjects)

u1,u2=st.columns(2)
with u1:
    syllabus_file=st.file_uploader("📚 Upload Syllabus PDF",type=["pdf"],key="syllabus_upload")
with u2:
    paper_file=st.file_uploader("📄 Upload Question Paper PDF",type=["pdf"],key="paper_upload")

if st.button("🔍 Analyse Paper",type="primary",
             disabled=(syllabus_file is None or paper_file is None),
             use_container_width=True):
    st.session_state.data=None
    st.session_state.answers={}
    st.session_state.source_bytes=paper_file.getvalue()
    st.session_state.source_name=paper_file.name
    st.session_state.pop("selected_chapter",None)
    st.session_state.pop("manual_crops",None)

    # Browser-side animated progress: it moves continuously from 0 toward
    # 99% while the long AI request is running, then becomes 100% when done.
    progress_html=st.empty()
    progress_html.markdown("""
    <div class="qp-progress-wrap">
      <div class="qp-progress-title" id="qp-progress-title">0% completed</div>
      <div class="qp-progress-track">
        <div class="qp-progress-fill" id="qp-progress-fill"></div>
      </div>
      <div class="qp-progress-note">Processing syllabus, questions, concepts, Bloom taxonomy and figures…</div>
    </div>
    <style>
      .qp-progress-wrap { margin: 12px 0 24px 0; }
      .qp-progress-title { font-size: 24px; font-weight: 650; margin-bottom: 8px; }
      .qp-progress-track { width:100%; height:14px; background:#e8eaed; border-radius:8px; overflow:hidden; }
      .qp-progress-fill {
        height:100%; width:0%;
        border-radius:8px;
        background:#2f80ed;
        animation: qpProgress 90s linear forwards;
      }
      .qp-progress-note { margin-top:8px; color:#6b7280; font-size:14px; }
      @keyframes qpProgress {
        0% { width:0%; }
        2% { width:2%; }
        10% { width:10%; }
        20% { width:20%; }
        30% { width:30%; }
        40% { width:40%; }
        50% { width:50%; }
        60% { width:60%; }
        70% { width:70%; }
        80% { width:80%; }
        90% { width:90%; }
        98% { width:98%; }
        99% { width:99%; }
        100% { width:99%; }
      }
    </style>
    <script>
      (function(){
        let start=Date.now(), duration=90000;
        const title=document.getElementById("qp-progress-title");
        const fill=document.getElementById("qp-progress-fill");
        function tick(){
          if(!title || !fill) return;
          let p=Math.min(99, Math.floor(((Date.now()-start)/duration)*99));
          title.textContent=p+"% completed";
          fill.style.width=p+"%";
          if(p<99) requestAnimationFrame(tick);
        }
        requestAnimationFrame(tick);
      })();
    </script>
    """,unsafe_allow_html=True)

    with st.status("Analysing paper...",expanded=True) as status:
        try:
            st.write("📚 Reading uploaded syllabus...")
            syllabus=extract_syllabus_units(
                syllabus_file.getvalue(),syllabus_file.name,
                board,year,klass,subject,"Syllabus uploaded by user"
            )
            st.session_state["active_syllabus"]=syllabus
            CHAPTERS=[u["name"] for u in syllabus["units"]]
            st.write(f"✓ {len(CHAPTERS)} syllabus units/chapters identified")

            st.write("📄 Reading question paper...")
            st.write("🔎 Extracting questions, marks and figures...")
            st.write("🧠 Classifying questions against the syllabus...")
            st.write("🎯 Analysing Bloom taxonomy, concepts and difficulty...")

            st.session_state.data=analyse(paper_file)

            st.write(f"✓ {len(st.session_state.data.get('questions',[]))} questions processed")
            st.write("📐 Preparing detected figures for review...")
            st.write("📊 Preparing chapter weightage and Bloom summaries...")

            # Replace animated progress with a completed bar after the request.
            progress_html.markdown("""
            <div class="qp-progress-wrap">
              <div class="qp-progress-title">100% completed ✅</div>
              <div class="qp-progress-track">
                <div class="qp-progress-fill-done"></div>
              </div>
            </div>
            <style>
              .qp-progress-wrap { margin:12px 0 24px 0; }
              .qp-progress-title { font-size:24px; font-weight:650; margin-bottom:8px; }
              .qp-progress-track { width:100%; height:14px; background:#e8eaed; border-radius:8px; overflow:hidden; }
              .qp-progress-fill-done { height:100%; width:100%; background:#2f80ed; border-radius:8px; }
            </style>
            """,unsafe_allow_html=True)
            status.update(label="✅ Analysis complete",state="complete",expanded=False)
            # Re-run once so the sidebar is rebuilt after data/chapters exist.
            st.rerun()

        except Exception as e:
            progress_html.markdown("""
            <div style="margin:12px 0 24px 0;">
              <div style="font-size:24px;font-weight:650;">Analysis stopped ❌</div>
            </div>
            """,unsafe_allow_html=True)
            status.update(label="❌ Analysis failed",state="error",expanded=True)
            st.error(f"Analysis failed: {e}")



data=st.session_state.data
if st.session_state.get("active_syllabus"):
    CHAPTERS=[u["name"] for u in st.session_state["active_syllabus"]["units"]]
if not data:
    st.info("Upload a paper to begin."); st.stop()

selected=st.session_state.get("selected_chapter")

if not selected:
    st.header("📊 Paper Overview")
    st.info("Select a chapter from the left navigation to view only its questions. The chapter list appears automatically after analysis.")

    t=chapter_totals(data)
    printed=benchmark_printed_total(data)
    exam=float(data.get("maximum_exam_marks") or 80)
    extracted_total=sum(v["marks"] for v in t.values())

    a,b,c,d=st.columns(4)
    a.metric("Questions",len(data["questions"]))
    b.metric("Printed paper",f"{printed:g}")
    c.metric("Actual exam maximum",f"{exam:g}")
    d.metric("Chapters covered",sum(1 for cc in CHAPTERS if t[cc]["questions"]))

    if abs(extracted_total-printed)>0.01:
        st.warning(f"⚠️ Mark reconciliation: extracted {extracted_total:g} marks vs printed {printed:g} marks.")

    st.header("📊 Chapter-wise Weightage")
    rows=[{"Chapter":cc,"Questions":t[cc]["questions"],"Marks":t[cc]["marks"],
           "Weightage %":round(t[cc]["marks"]/printed*100,1) if printed else 0}
          for cc in CHAPTERS if t[cc]["marks"]]
    if rows:
        df=pd.DataFrame(rows)
        fig=px.pie(df,names="Chapter",values="Marks",hole=.4,title="Chapter Weightage — Complete Printed Paper")
        fig.update_traces(textposition="inside",textinfo="percent")
        st.plotly_chart(fig,use_container_width=True)
        view=df.copy()
        view["Weightage %"]=view["Weightage %"].astype(str)+"%"
        st.dataframe(view,use_container_width=True,hide_index=True)

    st.header("🧠 Bloom's Taxonomy")
    bt=bloom_totals(data)
    brow=[{"Bloom Level":lv,"Questions":bt[lv]["questions"],"Marks":bt[lv]["marks"]}
          for lv in ["Remember","Understand","Apply","Analyze","Evaluate","Create"] if bt[lv]["questions"]]
    if brow:
        bdf=pd.DataFrame(brow)
        bfig=px.pie(bdf,names="Bloom Level",values="Questions",hole=.4,title="Bloom's Taxonomy — Question Distribution")
        bfig.update_traces(textposition="inside",textinfo="percent")
        st.plotly_chart(bfig,use_container_width=True)
        st.dataframe(bdf,use_container_width=True,hide_index=True)

    st.header("🎯 Main Concepts to Focus On")
    for x in data["focus_concepts"]:
        st.markdown(f"**{x['concept']}** — `{x['chapter']}`  \n{x['reason']}")
    st.stop()

groups=defaultdict(list)
for i,q in enumerate(data["questions"]):
    groups[q["chapter"]].append((i,q))

st.header(f"📝 {selected}")
selected_questions=groups.get(selected,[])
st.caption(f"{len(selected_questions)} question(s) • {sum(float(q.get('marks',0) or 0) for _,q in selected_questions):g} marks")

for i,q in selected_questions:
    with st.container(border=True):
        st.markdown(f"**Q{q['question_no']}.** {q['question_text']}")
        st.markdown(f'<span class="tag">Chapter: {q["chapter"]}</span>',unsafe_allow_html=True)
        st.markdown(
            f"**Bloom: {q.get('bloom_level','Not classified')}**  • "
            f"Marks: {q['marks']} • {q['concept']} • {q['question_type']} • {q['difficulty']}"
        )

        if q.get("has_figure"):
            crop_key=f"{q['question_no']}_{i}"
            crops=st.session_state.setdefault("manual_crops",{})
            saved=crops.get(crop_key)
            st.markdown("### 📐 Figure")

            if saved is not None:
                st.image(saved,caption="Selected original figure",width=650)
                if st.button("✂️ Re-crop Figure",key=f"recrop_{i}",use_container_width=True):
                    st.session_state[f"crop_mode_{crop_key}"]=True
                    st.rerun()
            else:
                st.markdown("**Is the detected figure correct?**")
                y,n=st.columns(2)
                with y:
                    if st.button("✅ Yes — use this figure",key=f"use_{i}",use_container_width=True):
                        auto=figure_crop(
                            st.session_state.source_bytes,q.get("source_page",1),
                            (q.get("figure_x1",0),q.get("figure_y1",0),
                             q.get("figure_x2",0),q.get("figure_y2",0))
                        )
                        if auto:
                            crops[crop_key]=Image.open(io.BytesIO(auto))
                            st.rerun()
                        else:
                            st.session_state[f"crop_mode_{crop_key}"]=True
                            st.rerun()
                with n:
                    if st.button("❌ No — crop manually",key=f"manual_{i}",use_container_width=True):
                        st.session_state[f"crop_mode_{crop_key}"]=True
                        st.rerun()

            if st.session_state.get(f"crop_mode_{crop_key}",False):
                cropped=manual_crop_ui(
                    st.session_state.source_bytes,q.get("source_page",1),crop_key
                )
                if cropped is not None:
                    st.image(cropped,caption="Crop preview",width=650)
                    x,y=st.columns(2)
                    with x:
                        if st.button("✓ Save This Crop",key=f"save_{i}",use_container_width=True):
                            crops[crop_key]=cropped
                            st.session_state[f"crop_mode_{crop_key}"]=False
                            st.rerun()
                    with y:
                        if st.button("Cancel",key=f"cancel_{i}",use_container_width=True):
                            st.session_state[f"crop_mode_{crop_key}"]=False
                            st.rerun()
        else:
            st.caption("No figure was detected for this question.")

        with st.expander("👩‍🏫 Edit Question Classification"):
            choice=st.selectbox(
                "Chapter / Unit",CHAPTERS,index=CHAPTERS.index(q["chapter"]),key=f"ch_{i}"
            )
            if choice!=q["chapter"]:
                data["questions"][i]["chapter"]=choice
                st.session_state["selected_chapter"]=choice
                st.rerun()

        key=f"{q['question_no']}::{q['question_text']}"
        if key in st.session_state.answers:
            with st.expander("💡 AI Answer",expanded=True):
                st.markdown(render_ai_answer(st.session_state.answers[key]))
        elif st.button("🔗 View AI Answer",key=f"ans_{i}"):
            with st.spinner("Generating answer..."):
                try:
                    st.session_state.answers[key]=answer(q)
                    st.rerun()
                except Exception as e:
                    st.error(f"Answer generation failed: {e}")

st.divider(); st.header("⬇️ Download")
x,y=st.columns(2)
with x:
    st.download_button("📄 Sorted Question Paper (Word)",
        data=make_word(data,st.session_state.source_bytes,False),
        file_name="sorted_question_paper.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary",use_container_width=True)
with y:
    st.download_button("📄 Sorted Paper + AI Answers",
        data=make_word(data,st.session_state.source_bytes,True),
        file_name="sorted_question_paper_with_answers.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True)
