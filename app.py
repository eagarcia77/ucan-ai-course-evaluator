import os, io, re, json, uuid
from pathlib import Path
from typing import List
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from openai import OpenAI
from pypdf import PdfReader
from docx import Document
from pptx import Presentation
from openpyxl import load_workbook
from bs4 import BeautifulSoup

BASE=Path(__file__).resolve().parent
REPORTS=BASE/"reports"; REPORTS.mkdir(exist_ok=True)
CRITERIA=json.loads((BASE/"institutional_criteria.json").read_text(encoding="utf-8"))
MAX_CHARS=int(os.getenv("MAX_EXTRACTED_CHARS","350000"))
MODEL=os.getenv("OPENAI_MODEL","gpt-5")
client=OpenAI(api_key=os.getenv("OPENAI_API_KEY")) if os.getenv("OPENAI_API_KEY") else None
app=FastAPI(title="UCAN AI Course Evaluator")

def extract(name,raw):
    ext=Path(name).suffix.lower()
    if ext==".pdf":
        r=PdfReader(io.BytesIO(raw)); return "\n".join((p.extract_text() or "") for p in r.pages)
    if ext==".docx":
        d=Document(io.BytesIO(raw)); out=[p.text for p in d.paragraphs if p.text.strip()]
        for t in d.tables:
            for row in t.rows: out.append(" | ".join(c.text for c in row.cells))
        return "\n".join(out)
    if ext==".pptx":
        prs=Presentation(io.BytesIO(raw)); out=[]
        for i,s in enumerate(prs.slides,1):
            out.append(f"[SLIDE {i}]")
            for sh in s.shapes:
                if hasattr(sh,"text") and sh.text.strip(): out.append(sh.text)
        return "\n".join(out)
    if ext==".xlsx":
        wb=load_workbook(io.BytesIO(raw),data_only=True,read_only=True); out=[]
        for ws in wb.worksheets:
            out.append(f"[SHEET {ws.title}]")
            for row in ws.iter_rows(values_only=True):
                vals=[str(v) for v in row if v is not None]
                if vals: out.append(" | ".join(vals))
        return "\n".join(out)
    if ext in {".html",".htm"}: return BeautifulSoup(raw.decode("utf-8","ignore"),"html.parser").get_text("\n")
    if ext in {".txt",".md",".csv"}: return raw.decode("utf-8","ignore")
    return f"[FORMATO NO EXTRAÍDO AUTOMÁTICAMENTE: {name}]"

def base_schema():
    return {"summary":{"overall_score":0,"meets":0,"partially_meets":0,"does_not_meet":0,"executive_summary":"","recommendation":""},
    "syllabus_review":[],"institutional_review":[],"qm_review":[],"udl_review":[],"five_e_review":[],"objectives_review":[],
    "citation_review":[],"compliance_review":[],"module_review":[],"action_plan":[]}

def build_prompt(meta,docs):
    return f"""Actúa como un panel experto de evaluación curricular de educación superior en línea.
Debes ser conservador y basarte SOLO en la evidencia de los archivos. No inventes cumplimiento.

CURSO:
{json.dumps(meta,ensure_ascii=False,indent=2)}

CRITERIOS INSTITUCIONALES CARGADOS:
{json.dumps(CRITERIA,ensure_ascii=False,indent=2)}

EVALÚA:
1) Prontuario oficial, descripción, créditos, objetivos, contenido, actividades, evaluación y plan de trabajo.
2) Alineación prontuario → módulos → actividades → evaluaciones.
3) Criterios institucionales VAEL/UIPR.
4) Quality Matters a NIVEL ALTO: overview, objetivos, assessment, materiales, interacción, tecnología, learner support, accessibility/usability. No reproduzcas una rúbrica propietaria completa.
5) DUA/UDL: engagement, representation, action/expression.
6) Modelo instruccional 5E: Engage, Explore, Explain, Elaborate y Evaluate. Evalúa si cada módulo evidencia una secuencia pedagógica coherente: Engage activa conocimientos previos y motivación; Explore promueve indagación o experiencia activa antes de la explicación formal; Explain desarrolla conceptos y permite al estudiante explicar/comunicar comprensión; Elaborate extiende y transfiere el aprendizaje a nuevas situaciones; Evaluate recoge evidencia formativa y/o sumativa del logro. No marques cumplimiento solo porque existan actividades con nombres similares; verifica la función pedagógica y la secuencia.
7) Objetivos con Bloom revisado: Remember, Understand, Apply, Analyze, Evaluate, Create. Evalúa verbo observable, claridad, condición/situación y criterio/adecuacidad.
8) Webb DOK 1-4. No clasifiques por el verbo solamente: usa complejidad de la tarea.
9) Estilo de citación seleccionado: {meta["citation_style"]}. Para APA 7 verifica citas narrativas/parentéticas, correspondencia citas-referencias, autores/fechas, títulos, cursivas, DOI/URL, sitios web, videos, figuras e imágenes.
10) Accesibilidad: encabezados, alt text, captions/transcripts, contraste, tablas, enlaces descriptivos, navegación.
11) Derechos de autor, propiedad intelectual y atribución.
12) Rúbricas, claridad de instrucciones, interacción regular sustantiva, evaluaciones auténticas/custodiadas.
13) Identifica exactamente DÓNDE corregir: archivo, módulo, sección y página/slide si la extracción lo indica.

Para cada hallazgo devuelve:
criterion, status (Meets | Partially Meets | Does Not Meet), priority (High | Medium | Low), location, evidence, comment, recommendation, example.
En objetivos, evidence debe contener el objetivo original y comment debe explicar Bloom + DOK.
En citas/referencias, señala dónde falta cita o qué referencia corregir.
Cuando algo requiere prueba técnica externa (WAVE/Ally, enlaces en vivo, Turnitin/SafeAssign), indícalo como pendiente de validación humana/técnica.
La regla institucional cargada exige cumplimiento total para certificación; la recomendación final debe respetar esa regla.

Devuelve SOLO JSON válido con esta estructura:
{json.dumps(base_schema(),ensure_ascii=False,indent=2)}

ARCHIVOS:
{docs[:MAX_CHARS]}"""

@app.get("/api/health")
def health(): return {"ok":True,"ai_ready":client is not None,"model":MODEL}

@app.post("/api/evaluate")
async def evaluate(course_code:str=Form(...),course_name:str=Form(...),credits:str=Form(""),
                   citation_style:str=Form("APA 7"),course_description:str=Form(""),
                   frameworks:str=Form("[]"),files:List[UploadFile]=File(...)):
    if not client: raise HTTPException(503,"Configure OPENAI_API_KEY en el servidor.")
    chunks=[]; names=[]
    for f in files:
        raw=await f.read()
        if len(raw)>35*1024*1024: raise HTTPException(400,f"{f.filename} excede 35 MB.")
        try: txt=extract(f.filename,raw)
        except Exception as e: txt=f"[ERROR: {e}]"
        chunks.append(f"\n===== ARCHIVO: {f.filename} =====\n{txt}")
        names.append(f.filename)
    meta={"course_code":course_code,"course_name":course_name,"credits":credits,"citation_style":citation_style,
          "course_description":course_description,"frameworks":json.loads(frameworks or "[]"),"files":names}
    try:
        resp=client.responses.create(model=MODEL,input=[
            {"role":"system","content":"Eres un evaluador institucional riguroso. Devuelve exclusivamente JSON válido."},
            {"role":"user","content":build_prompt(meta,"".join(chunks))}
        ])
        text=resp.output_text.strip()
        text=re.sub(r"^```(?:json)?\s*|\s*```$","",text,flags=re.S)
        data=json.loads(text)
    except Exception as e: raise HTTPException(500,f"Error de evaluación IA: {e}")
    eid=str(uuid.uuid4()); data["evaluation_id"]=eid
    data["course"]={"code":course_code,"name":course_name,"citation_style":citation_style}
    (REPORTS/f"{eid}.json").write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding="utf-8")
    return data

@app.get("/api/report/{eid}/docx")
def docx_report(eid:str):
    p=REPORTS/f"{eid}.json"
    if not p.exists(): raise HTTPException(404,"Informe no encontrado.")
    d=json.loads(p.read_text(encoding="utf-8")); out=REPORTS/f"{eid}.docx"
    doc=Document(); c=d.get("course",{}); s=d.get("summary",{})
    doc.add_heading("UCAN AI Course Evaluator",0); doc.add_paragraph(f'{c.get("code","")} – {c.get("name","")}')
    doc.add_heading("Resumen ejecutivo",1); doc.add_paragraph(s.get("executive_summary",""))
    doc.add_paragraph(f'Puntuación: {s.get("overall_score","")}/100'); doc.add_paragraph("Recomendación: "+s.get("recommendation",""))
    sections=[("Prontuario y alineación","syllabus_review"),("Normas institucionales","institutional_review"),
    ("Quality Matters — alto nivel","qm_review"),("DUA / UDL","udl_review"),("Modelo 5E","five_e_review"),("Bloom + Webb DOK","objectives_review"),
    ("Citas y referencias","citation_review"),("Accesibilidad y derechos de autor","compliance_review"),("Módulos","module_review")]
    for title,key in sections:
        doc.add_heading(title,1)
        for x in d.get(key,[]):
            doc.add_heading(x.get("criterion","Hallazgo"),2)
            doc.add_paragraph(f'{x.get("status","")} | Prioridad: {x.get("priority","")} | Ubicación: {x.get("location","")}')
            if x.get("evidence"): doc.add_paragraph("Evidencia: "+x["evidence"])
            if x.get("comment"): doc.add_paragraph(x["comment"])
            if x.get("recommendation"): doc.add_paragraph("Recomendación: "+x["recommendation"])
            if x.get("example"): doc.add_paragraph("Ejemplo: "+x["example"])
    doc.add_heading("Plan de mejoramiento",1)
    for x in d.get("action_plan",[]): doc.add_paragraph(f'{x.get("priority","")} — {x.get("area","")}: {x.get("action","")} | {x.get("location","")}')
    doc.save(out); return FileResponse(out,filename=f'Informe_{course_safe(c.get("code","curso"))}.docx')

def course_safe(s): return re.sub(r"[^A-Za-z0-9_-]+","_",s)

app.mount("/",StaticFiles(directory=BASE,html=True),name="static")
