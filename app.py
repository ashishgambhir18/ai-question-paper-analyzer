
import io, json, os
from pathlib import Path
from collections import defaultdict
import streamlit as st
import pandas as pd
import plotly.express as px
from docx import Document
from docx.shared import Inches
from openai import OpenAI

try:
    import fitz
except Exception:
    fitz = None

ROOT = Path(__file__).parent
MODEL = os.getenv("OPENAI_MODEL","gpt-5.6-luna")

st.set_page_config(page_title="AI Question Paper Analyzer",page_icon="📘",layout="wide")

def get_client():
    key=os.getenv("OPENAI_API_KEY")
    if not key:
        try: key=st.secrets["OPENAI_API_KEY"]
        except Exception: key=None
    if not key:
        st.error("OPENAI_API_KEY is missing. Add it in Streamlit Secrets.")
        st.stop()
    return OpenAI(api_key=key)

def syllabus_options():
    out=[]
    for p in (ROOT/"syllabus").glob("*/*/*/*.json"):
        try:
            d=json.loads(p.read_text(encoding="utf-8"))
            out.append((d["board"],str(d["year"]),d["class"],d["subject"],p,d))
        except Exception:
            pass
    return out

OPTIONS=syllabus_options()
BOARDS=sorted({x[0] for x in OPTIONS})
if not OPTIONS:
    st.error("No syllabus files found.")
    st.stop()

def load_syllabus(board,year,klass,subject):
    for b,y,k,s,p,d in OPTIONS:
        if (b,y,k,s)==(board,year,klass,subject):
            return d
    return None

def doc_text(data):
    doc=Document(io.BytesIO(data))
    return "\n".join(p.text for p in doc.paragraphs if p.text.strip())

def input_content(upload):
    data=upload.getvalue()
    if upload.name.lower().endswith(".txt"):
        return [{"type":"input_text","text":data.decode("utf-8",errors="ignore")}]
    if upload.name.lower().endswith(".docx"):
        return [{"type":"input_text","text":doc_text(data)}]
    f=get_client().files.create(file=(upload.name,data),purpose="user_data")
    return [{"type":"input_file","file_id":f.id}]

def schema(units):
    allowed=[u["name"] for u in units]
    return {"type":"object","additionalProperties":False,"properties":{
        "printed_question_marks":{"type":"number"},
        "maximum_exam_marks":{"type":"number"},
        "section_structure":{"type":"string"},
        "questions":{"type":"array","items":{"type":"object","additionalProperties":False,
            "properties":{
                "question_no":{"type":"string"},"section":{"type":"string"},
                "question_text":{"type":"string"},"syllabus_unit":{"type":"string","enum":allowed},
                "marks":{"type":"number"},"concept":{"type":"string"},
                "question_type":{"type":"string"},"difficulty":{"type":"string"},
                "source_page":{"type":"integer"},"has_figure":{"type":"boolean"},
                "figure_x1":{"type":"number"},"figure_y1":{"type":"number"},
                "figure_x2":{"type":"number"},"figure_y2":{"type":"number"}
            },
            "required":["question_no","section","question_text","syllabus_unit","marks",
                        "concept","question_type","difficulty","source_page","has_figure",
                        "figure_x1","figure_y1","figure_x2","figure_y2"]}},
        "focus_concepts":{"type":"array","items":{"type":"object","additionalProperties":False,
            "properties":{"concept":{"type":"string"},"unit":{"type":"string","enum":allowed},
                        "reason":{"type":"string"}},
            "required":["concept","unit","reason"]}}
    },"required":["printed_question_marks","maximum_exam_marks","section_structure",
                  "questions","focus_concepts"]}

def analyse(upload,syll):
    units=syll["units"]
    unit_text="\n".join(f"- {u['name']}" for u in units)
    prompt=f"""
You are an examination-paper analyst. Match the uploaded {syll['board']} {syll['class']}
{syll['subject']} paper to the official syllabus structure below.

Official syllabus: {syll['board']} {syll['year']} {syll['class']} {syll['subject']}
Source: {syll['source_url']}

Allowed syllabus units:
{unit_text}

Rules:
- Extract every assessable question/sub-question without merging neighbouring questions.
- Assign exactly one official syllabus unit to each item.
- Keep question text limited to that item.
- Preserve marks and numbering.
- printed_question_marks means all printed marks including optional questions.
- maximum_exam_marks means the maximum a candidate can actually score after choices.
- Identify figures and give a tight normalized box containing ONLY the visual,
  its own labels/arrows/axes, excluding surrounding prose.
- source_page is the PDF page number, 1-based.
- Identify concepts and important focus concepts.
- Do not invent syllabus units.
"""
    r=get_client().responses.create(
        model=MODEL,
        input=[{"role":"user","content":input_content(upload)+[{"type":"input_text","text":prompt}]}],
        text={"format":{"type":"json_schema","name":"syllabus_qp","strict":True,"schema":schema(units)}}
    )
    return json.loads(r.output_text)

def answer(q,syll):
    prompt=f"""Answer this {syll['board']} {syll['class']} {syll['subject']} exam question.
Official syllabus unit: {q['syllabus_unit']}
Marks: {q['marks']}
Question: {q['question_text']}
Give an exam-ready answer appropriate to the marks."""
    return get_client().responses.create(model=MODEL,input=prompt).output_text.strip()

def figure_crop(pdf_bytes,q):
    if fitz is None: return None
    try:
        doc=fitz.open(stream=pdf_bytes,filetype="pdf")
        page=doc[q["source_page"]-1]; p=page.rect
        c=[max(0,min(1,float(q[k]))) for k in ("figure_x1","figure_y1","figure_x2","figure_y2")]
        cand=fitz.Rect(p.x0+c[0]*p.width,p.y0+c[1]*p.height,
                       p.x0+c[2]*p.width,p.y0+c[3]*p.height)
        objects=[]
        for d in page.get_drawings():
            r=d.get("rect")
            if r and r.intersects(cand) and r.width*r.height>10: objects.append(r)
        for im in page.get_image_info(xrefs=True):
            r=im.get("bbox")
            if r:
                r=fitz.Rect(r)
                if r.intersects(cand) and r.width*r.height>100: objects.append(r)
        if objects:
            inside=[r for r in objects if cand.contains(fitz.Point((r.x0+r.x1)/2,(r.y0+r.y1)/2))]
            if inside:
                crop=inside[0]
                for r in inside[1:]:
                    if (max(0,max(crop.x0,r.x0)-min(crop.x1,r.x1))<cand.width*.12 and
                        max(0,max(crop.y0,r.y0)-min(crop.y1,r.y1))<cand.height*.18):
                        crop|=r
                cand=crop
        pix=page.get_pixmap(matrix=fitz.Matrix(2.5,2.5),clip=cand,alpha=False)
        out=pix.tobytes("png"); doc.close(); return out
    except Exception:
        return None

st.title("📘 AI Question Paper Analyzer — V9")
st.caption("Syllabus-driven ICSE / ISC analysis")

with st.sidebar:
    st.header("1. Select syllabus")
    board=st.selectbox("Board",BOARDS)
    years=sorted({x[1] for x in OPTIONS if x[0]==board},reverse=True)
    year=st.selectbox("Examination Year",years)
    classes=sorted({x[2] for x in OPTIONS if x[0]==board and x[1]==year})
    klass=st.selectbox("Class",classes)
    subjects=sorted({x[3] for x in OPTIONS if x[:3]==(board,year,klass)})
    subject=st.selectbox("Subject",subjects)
    syll=load_syllabus(board,year,klass,subject)
    if syll:
        st.success(f"Syllabus loaded: {len(syll['units'])} units")
        st.caption(f"Official source: {syll['source_url']}")

upload=st.file_uploader("2. Upload Question Paper",type=["pdf","docx","txt"])
if st.button("🔍 Analyse Paper",type="primary",disabled=upload is None,use_container_width=True):
    with st.spinner("Matching questions to the official syllabus..."):
        try:
            st.session_state["data"]=analyse(upload,syll)
            st.session_state["bytes"]=upload.getvalue()
            st.session_state["name"]=upload.name
            st.session_state["answers"]={}
        except Exception as e: st.error(f"Analysis failed: {e}")

data=st.session_state.get("data")
if not data:
    st.info("Select Board → Year → Class → Subject, then upload a paper.")
    st.stop()

units=[u["name"] for u in syll["units"]]
tot=defaultdict(float)
for q in data["questions"]: tot[q["syllabus_unit"]]+=float(q["marks"])
printed=sum(tot.values())
exam=float(data.get("maximum_exam_marks") or 0)

a,b,c=st.columns(3)
a.metric("Printed paper",f"{printed:g} marks")
b.metric("Actual exam maximum",f"{exam:g} marks")
c.metric("Syllabus units used",sum(v>0 for v in tot.values()))

st.subheader("📊 Syllabus Weightage")
rows=[{"Syllabus Unit":u,"Marks":tot[u],"Weightage %":round(tot[u]/printed*100,1)}
      for u in units if tot[u]]
if rows:
    df=pd.DataFrame(rows)
    st.plotly_chart(px.pie(df,names="Syllabus Unit",values="Marks",hole=.4,
                           title=f"{board} {klass} {subject} — {year}"),
                    use_container_width=True)
    st.dataframe(df.assign(**{"Weightage %":df["Weightage %"].astype(str)+"%"}),
                 use_container_width=True,hide_index=True)

st.subheader("🎯 Main Concepts to Focus On")
for x in data["focus_concepts"]:
    st.markdown(f"**{x['concept']}** — `{x['unit']}`  \n{x['reason']}")

st.subheader("✏️ Teacher Review")
for i,q in enumerate(data["questions"]):
    with st.container(border=True):
        st.markdown(f"**Q{q['question_no']}.** {q['question_text']}")
        new=st.selectbox("Syllabus Unit",units,index=units.index(q["syllabus_unit"]),key=f"unit_{i}")
        if new!=q["syllabus_unit"]:
            data["questions"][i]["syllabus_unit"]=new; st.rerun()

st.subheader("📚 Sorted Question Paper")
groups=defaultdict(list)
for q in data["questions"]: groups[q["syllabus_unit"]].append(q)
for u in units:
    if not groups[u]: continue
    st.markdown(f"### {u}")
    for q in groups[u]:
        with st.container(border=True):
            st.markdown(f"**Q{q['question_no']}.** {q['question_text']}")
            st.caption(f"Marks: {q['marks']} • Concept: {q['concept']} • {q['difficulty']}")
            if q["has_figure"] and st.session_state.get("name","").lower().endswith(".pdf"):
                img=figure_crop(st.session_state["bytes"],q)
                if img: st.image(img,caption="Original figure")
            key=q["question_no"]+"|"+q["question_text"]
            if st.button("🔗 View AI Answer",key="ansbtn_"+key):
                with st.spinner("Generating answer..."):
                    st.session_state["answers"][key]=answer(q,syll)
            if key in st.session_state["answers"]:
                with st.expander("💡 AI Answer",expanded=True):
                    st.markdown(st.session_state["answers"][key])

st.subheader("⬇️ Download")
doc=Document()
doc.add_heading("Sorted Question Paper",0)
doc.add_paragraph(f"{board} | {klass} | {subject} | Examination Year {year}")
doc.add_paragraph(f"Official syllabus source: {syll['source_url']}")
doc.add_heading("Syllabus Weightage",1)
for r in rows: doc.add_paragraph(f"{r['Syllabus Unit']}: {r['Marks']} marks ({r['Weightage %']}%)")
for u in units:
    if not groups[u]: continue
    doc.add_heading(u,1)
    for q in groups[u]:
        doc.add_paragraph(f"Q{q['question_no']}. {q['question_text']}")
out=io.BytesIO(); doc.save(out)
st.download_button("📄 Download Sorted Question Paper (Word)",out.getvalue(),
                   "sorted_question_paper_v9.docx",
                   "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                   type="primary",use_container_width=True)
