
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

def extract_uploaded_syllabus(upload,board,year,klass,subject):
    f=get_client().files.create(file=(upload.name,upload.getvalue()),purpose="user_data")
    prompt=f"""Extract the official syllabus hierarchy from this CISCE document.
Board={board}; Year={year}; Class={klass}; Subject={subject}.
Preserve official terminology and order. Return only units/sections useful
for classifying questions; do not invent or rename."""
    r=get_client().responses.create(model=MODEL,input=[
        {"role":"user","content":[{"type":"input_file","file_id":f.id},
                                  {"type":"input_text","text":prompt}]}])
    rr=get_client().responses.create(
        model=MODEL,
        input="Convert the following official syllabus extraction to JSON units. "
              "Do not invent or rename.\n"+r.output_text,
        text={"format":{"type":"json_schema","name":"syll_units","strict":True,
          "schema":{"type":"object","additionalProperties":False,
            "properties":{"units":{"type":"array","items":{
              "type":"object","additionalProperties":False,
              "properties":{"name":{"type":"string"},
                           "topics":{"type":"array","items":{"type":"string"}}},
              "required":["name","topics"]}}},
            "required":["units"]}}})
    return {"board":board,"year":int(year),"class":klass,"subject":subject,
            "source_url":"Uploaded official CISCE syllabus: "+upload.name,
            "units":json.loads(rr.output_text)["units"]}


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
    prompt=f"""
Answer this {syll['board']} {syll['class']} {syll['subject']} exam question.

Official syllabus unit: {q['syllabus_unit']}
Marks: {q['marks']}
Question:
{q['question_text']}

Give an exam-ready answer appropriate to the marks.

For numerical questions use Given, Formula, Substitution, Calculation and Final Answer.
For theory questions use concise, clear points.

MATHEMATICAL FORMATTING:
- Use $...$ for inline mathematics.
- Use $$...$$ for equations on separate lines.
- Do not use \\(...\\) or \\[...\\] delimiters.
- Do not put equations in code blocks.
- Use proper LaTeX for fractions, powers, subscripts, Greek letters and units.
- Put an important final result on a separate display-math line.
"""
    return get_client().responses.create(model=MODEL,input=prompt).output_text.strip()

def figure_crop(pdf_bytes,q):
    """Extract the original figure while retaining outer wires and labels."""
    if fitz is None:
        return None
    try:
        doc=fitz.open(stream=pdf_bytes,filetype="pdf")
        page=doc[int(q["source_page"])-1]; p=page.rect
        c=[max(0,min(1,float(q[k]))) for k in
           ("figure_x1","figure_y1","figure_x2","figure_y2")]
        cand=fitz.Rect(p.x0+c[0]*p.width,p.y0+c[1]*p.height,
                       p.x0+c[2]*p.width,p.y0+c[3]*p.height)
        expanded=fitz.Rect(max(p.x0,cand.x0-cand.width*.15),
                           max(p.y0,cand.y0-cand.height*.10),
                           min(p.x1,cand.x1+cand.width*.15),
                           min(p.y1,cand.y1+cand.height*.10))

        visual=[]
        try:
            for d in page.get_drawings():
                r=d.get("rect")
                if r and r.width>1 and r.height>1 and r.intersects(expanded):
                    visual.append(r)
        except Exception: pass
        try:
            for im in page.get_image_info(xrefs=True):
                r=im.get("bbox")
                if r:
                    r=fitz.Rect(r)
                    if r.intersects(expanded): visual.append(r)
        except Exception: pass

        crop=cand
        if visual:
            cx=(cand.x0+cand.x1)/2; cy=(cand.y0+cand.y1)/2
            visual.sort(key=lambda r:((r.x0+r.x1)/2-cx)**2+((r.y0+r.y1)/2-cy)**2)
            selected=[visual[0]]
            changed=True
            while changed:
                changed=False
                u=selected[0]
                for r in selected[1:]: u|=r
                for r in visual:
                    if r in selected: continue
                    gx=max(0,max(u.x0,r.x0)-min(u.x1,r.x1))
                    gy=max(0,max(u.y0,r.y0)-min(u.y1,r.y1))
                    if ((gx <= expanded.width*.20 and r.y0 <= u.y1 and r.y1 >= u.y0) or
                        (gy <= expanded.height*.20 and r.x0 <= u.x1 and r.x1 >= u.x0)):
                        selected.append(r); changed=True
            crop=selected[0]
            for r in selected[1:]: crop|=r

            # Add compact nearby text labels belonging to the diagram.
            try:
                for b in page.get_text("blocks"):
                    if len(b)<5: continue
                    tr=fitz.Rect(b[:4]); txt=(b[4] or "").strip()
                    if not txt or not tr.intersects(expanded): continue
                    if tr.width>expanded.width*.75 and tr.height<expanded.height*.18:
                        continue
                    dx=max(0,max(crop.x0,tr.x0)-min(crop.x1,tr.x1))
                    dy=max(0,max(crop.y0,tr.y0)-min(crop.y1,tr.y1))
                    if dx<expanded.width*.10 and dy<expanded.height*.15:
                        crop|=tr
            except Exception: pass

            mx=min(12,p.width*.012); my=min(12,p.height*.012)
            crop=fitz.Rect(max(p.x0,crop.x0-mx),max(p.y0,crop.y0-my),
                           min(p.x1,crop.x1+mx),min(p.y1,crop.y1+my))

        pix=page.get_pixmap(matrix=fitz.Matrix(2.5,2.5),clip=crop,alpha=False)
        out=pix.tobytes("png"); doc.close(); return out
    except Exception:
        return None


def render_ai_answer(text):
    """Normalize common AI math delimiters for Streamlit's LaTeX renderer."""
    if not text: return ""
    text=text.replace("\r\n","\n").replace("\r","\n")
    text=text.replace(r"\[","$$").replace(r"\]","$$")
    text=text.replace(r"\(","$").replace(r"\)", "$")
    lines=[]
    for line in text.split("\n"):
        s=line.strip()
        if len(s)>=4 and s.startswith("[") and s.endswith("]") and "\\" in s:
            s=s[1:-1].strip()
            line=s if s.startswith("$") else "$$\n"+s+"\n$$"
        lines.append(line)
    return "\n".join(lines)

st.title("📘 AI Question Paper Analyzer — V9")
st.caption("Syllabus-driven ICSE / ISC analysis")

with st.sidebar:
    st.header("1. Select syllabus")
    board=st.selectbox("Board",["ICSE","ISC"])
    year=st.selectbox("Examination Year",list(range(2027,2014,-1)))

    ICSE_SUBJECTS=[
        "English","Second Languages","History, Civics and Geography","Mathematics",
        "Science","Economics","Commercial Studies","Modern Foreign Language",
        "Classical Language","Environmental Science","Computer Applications",
        "Economic Applications","Commercial Applications","Art","Performing Arts",
        "Home Science","Cookery","Fashion Designing","Physical Education","Yoga",
        "Technical Drawing Applications","Environmental Applications",
        "Mass Media & Communication","Hospitality Management",
        "Robotics and Artificial Intelligence"
    ]
    ISC_SUBJECTS=[
        "English","Indian Languages","Modern Foreign Languages","Classical Languages",
        "Elective English","History","Political Science","Geography","Sociology",
        "Psychology","Economics","Commerce","Accounts","Business Studies","Mathematics",
        "Physics","Chemistry","Biology","Home Science","Fashion Designing",
        "Electricity & Electronics","Engineering Science","Computer Science",
        "Geometrical & Mechanical Drawing","Geometrical & Building Drawing","Art",
        "Music","Physical Education","Environmental Science","Biotechnology",
        "Mass Media & Communication","Legal Studies","Hospitality Management"
    ]
    subjects=ICSE_SUBJECTS if board=="ICSE" else ISC_SUBJECTS
    subject=st.selectbox("Subject",subjects)
    klass="X" if board=="ICSE" else st.selectbox("Class",["XII","XI"])
    syll=load_syllabus(board,str(year),klass,subject)
    if syll:
        st.success(f"Official syllabus loaded: {len(syll['units'])} units")
        st.caption(f"Source: {syll['source_url']}")
    else:
        st.warning("This syllabus is not bundled yet.")
        st.caption("Upload the official CISCE syllabus PDF for this selection.")
        syllabus_upload=st.file_uploader("Upload official CISCE syllabus PDF",
                                          type=["pdf"],key="syllabus_pdf")
        if syllabus_upload is not None:
            st.session_state["uploaded_syllabus_bytes"]=syllabus_upload.getvalue()
            st.session_state["uploaded_syllabus_name"]=syllabus_upload.name

upload=st.file_uploader("2. Upload Question Paper",type=["pdf","docx","txt"])
if st.button("🔍 Analyse Paper",type="primary",disabled=upload is None,use_container_width=True):
    with st.spinner("Preparing syllabus and analysing the paper..."):
        try:
            if syll is None:
                raw=st.session_state.get("uploaded_syllabus_bytes")
                name=st.session_state.get("uploaded_syllabus_name")
                if not raw:
                    st.error("Upload the official CISCE syllabus PDF for this selection first.")
                    st.stop()
                class Uploaded:
                    def __init__(self,b,n): self.b=b; self.name=n
                    def getvalue(self): return self.b
                syll=extract_uploaded_syllabus(Uploaded(raw,name),board,str(year),klass,subject)
            st.session_state["active_syllabus"]=syll
            st.session_state["data"]=analyse(upload,syll)
            st.session_state["bytes"]=upload.getvalue()
            st.session_state["name"]=upload.name
            st.session_state["answers"]={}
        except Exception as e: st.error(f"Analysis failed: {e}")

data=st.session_state.get("data")
syll=st.session_state.get("active_syllabus", syll)
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
                    st.markdown(render_ai_answer(st.session_state["answers"][key]))

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
