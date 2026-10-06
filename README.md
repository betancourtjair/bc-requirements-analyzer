# BC Requirements Analyzer

App Streamlit que analiza requerimientos de Microsoft Dynamics 365 Business Central con Gemini (Consultor Funcional · Arquitecto AL · QA Tester) y exporta el reporte a Word.

## Estructura del repositorio

```
bc-analyzer/
├── app.py
├── requirements.txt
├── .gitignore
├── README.md
└── .streamlit/
    ├── config.toml
    └── secrets.toml.example
```

## Prueba local (opcional)

```bash
python -m venv .venv
.venv\Scripts\activate          # Windows  (macOS/Linux: source .venv/bin/activate)
pip install -r requirements.txt
copy .streamlit\secrets.toml.example .streamlit\secrets.toml   # y pon tu API Key
streamlit run app.py
```

## Despliegue gratuito en Streamlit Community Cloud

### Paso 1 — Subir a GitHub
1. Entra a https://github.com/new
2. Nombre: `bc-requirements-analyzer`. Visibilidad: **Private** recomendado (Streamlit Cloud permite apps desde repos privados).
3. Clic en **Create repository**.
4. En la página del repo vacío: **uploading an existing file** → arrastra `app.py`, `requirements.txt`, `.gitignore`, `README.md` y la carpeta `.streamlit` (sin `secrets.toml`) → **Commit changes**.

   O por terminal:
   ```bash
   git init && git add . && git commit -m "BC Requirements Analyzer"
   git branch -M main
   git remote add origin https://github.com/<tu-usuario>/bc-requirements-analyzer.git
   git push -u origin main
   ```

### Paso 2 — Conectar con Streamlit Community Cloud
1. Entra a https://share.streamlit.io y elige **Continue with GitHub**; autoriza el acceso (incluye repos privados si el tuyo lo es).
2. Clic en **Create app** → **Deploy a public app from GitHub**.
3. Repository: `<tu-usuario>/bc-requirements-analyzer` · Branch: `main` · Main file path: `app.py`.
4. **App URL**: escribe el subdominio deseado, p. ej. `fpt-bc-analyzer` → `https://fpt-bc-analyzer.streamlit.app`.

### Paso 3 — Configurar la API Key como Secret
1. En la misma pantalla abre **Advanced settings**.
2. Python version: **3.12**.
3. En **Secrets** pega:
   ```toml
   GEMINI_API_KEY = "AIza...tu_api_key..."
   ```
4. Clic en **Deploy**. La primera instalación tarda 2–4 minutos.

Para cambiarla después: menú ⋮ de la app → **Settings → Secrets** → guardar (la app se reinicia sola). La API Key se obtiene en https://aistudio.google.com/apikey.

### Paso 4 — Obtener y compartir la URL
- Al terminar el despliegue la app abre en `https://<subdominio>.streamlit.app`; esa es la URL pública.
- Para limitar quién entra: ⋮ → **Settings → Sharing** → *Only specific people can view this app* y agrega sus correos (solo disponible si el repo es privado).
- Cada `git push` a `main` redespliega automáticamente.

## Notas de operación
- Si Google retira `gemini-2.5-flash`, cambia el modelo en **Opciones avanzadas** del panel lateral o en `DEFAULT_MODEL` dentro de `app.py`.
- Las apps gratuitas se "duermen" tras días sin uso; el primer acceso las despierta en ~30 s.
- Con el Secret configurado, todos los usuarios consumen la cuota de esa API Key. Si la app es pública, cualquiera con la URL la gasta: usa Sharing restringido o deja que cada usuario ingrese su propia key.
