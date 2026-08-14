# Publicación en GitHub

## 1. Repositorio
Este proyecto está preparado para `ucan-ai-course-evaluator`.

## 2. No suba la API key
El archivo `.env` está excluido mediante `.gitignore`.
Nunca coloque `OPENAI_API_KEY` en `index.html`, `app.js` ni en archivos públicos.

## 3. GitHub Pages
GitHub Pages por sí solo no puede ejecutar FastAPI ni proteger la API key.
Recomendación: GitHub para el código + Render para desplegar la aplicación completa.

## 4. Render
Build Command:
`pip install -r requirements.txt`

Start Command:
`uvicorn app:app --host 0.0.0.0 --port $PORT`

Variables:
- `OPENAI_API_KEY`
- `OPENAI_MODEL` (opcional)

## 5. Modelo 5E
La aplicación incluye evaluación de Engage, Explore, Explain, Elaborate y Evaluate por módulo.
