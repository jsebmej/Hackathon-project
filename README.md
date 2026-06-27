# DataLex 1581 — ejecución local

La aplicación usa dos servicios locales:

- Frontend: `http://localhost:5173`
- API FastAPI: `http://localhost:8000`

## Inicio rápido

Desde PowerShell, en la raíz del proyecto:

```powershell
python -m pip install -r backend\requirements.txt
.\start_local.ps1
```

También puede iniciarse manualmente:

```powershell
# Terminal 1
Set-Location backend
python -m uvicorn main:app --reload --host 127.0.0.1 --port 8000

# Terminal 2, desde la raíz del proyecto
python -m http.server 5173 --bind 127.0.0.1 --directory frontend
```

Ollama debe estar activo en la URL indicada por `OLLAMA_BASE_URL` y debe tener disponible el modelo de `OLLAMA_MODEL`. Si Ollama no responde o no entrega JSON válido, DataLex usa los criterios y la clasificación de respaldo para mantener la evaluación activa.

## Inicio de sesión

- El formulario solo solicita `usuario` y `contraseña`.
- Si se configuran `LOCAL_AUTH_USERNAME` y `LOCAL_AUTH_PASSWORD` en `backend/.env`, esas son las credenciales válidas.
- Si no se configuran, se habilita el modo local de demostración y se acepta cualquier par no vacío.
- El botón de Google usa `GOOGLE_CLIENT_ID`. En Google Cloud debe estar autorizado `http://localhost:5173` como origen JavaScript.

## Resultados

Al finalizar, la aplicación siempre intenta:

1. generar un Excel con las hojas `evaluacion` y `resumen`;
2. generar el respaldo JSON;
3. guardar las mismas filas en `hackaton.evaluacion` y `hackaton.resumen` de PostgreSQL.

Si PostgreSQL falla, el endpoint conserva una respuesta exitosa, muestra una advertencia y permite descargar los archivos. Los archivos se crean de forma predeterminada en `backend/auditorias/`.
