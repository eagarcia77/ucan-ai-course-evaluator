# UCAN AI Course Evaluator

## Funciones
Sube múltiples PDF, DOCX, PPTX, XLSX, TXT, Markdown, HTML y CSV. La IA evalúa:
- prontuario, plan de trabajo, módulos y alineación curricular;
- criterios institucionales VAEL/UIPR;
- Quality Matters a nivel de categorías;
- DUA/UDL;
- Modelo instruccional 5E (Engage, Explore, Explain, Elaborate, Evaluate);
- Bloom revisado;
- Norman Webb DOK 1–4;
- APA 7 por defecto (también permite MLA 9 o Chicago);
- accesibilidad/WCAG;
- derechos de autor;
- actividades, rúbricas, interacción y evaluaciones;
- reporte detallado con ubicación, evidencia, comentario, recomendación y ejemplo;
- exportación a JSON y Word.

## Instalación
1. Python 3.11+
2. `pip install -r requirements.txt`
3. Configure `OPENAI_API_KEY` como variable de entorno.
4. `uvicorn app:app --reload --host 0.0.0.0 --port 8000`
5. Abra `http://localhost:8000`

## Render
Build: `pip install -r requirements.txt`
Start: `uvicorn app:app --host 0.0.0.0 --port $PORT`
Configure `OPENAI_API_KEY` en Environment.

## Seguridad
Nunca coloque la API key en HTML o JavaScript del navegador. Este proyecto usa el backend para protegerla.

## Quality Matters
La versión incluida usa categorías de alto nivel. Para una evaluación oficial con criterios licenciados, incorpore al backend/configuración la rúbrica autorizada por la institución.

## Límites
La evaluación por IA no sustituye revisión humana, WAVE/Ally, validación de enlaces en vivo, SafeAssign/Turnitin ni la certificación institucional.

## Modelo 5E
La evaluación 5E revisa cada módulo para identificar evidencia real de:
- **Engage**: activación, motivación y conocimientos previos.
- **Explore**: indagación o experiencia activa.
- **Explain**: explicación conceptual y construcción de significado.
- **Elaborate**: aplicación, transferencia y profundización.
- **Evaluate**: evidencia formativa o sumativa del aprendizaje.

El sistema no marca cumplimiento solo por encontrar los nombres de las fases. Evalúa la **función pedagógica, secuencia y alineación** de las actividades.
