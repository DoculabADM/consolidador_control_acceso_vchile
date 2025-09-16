# Consolidador de Asistencia (Streamlit)

App en Streamlit para consolidar la hoja **`asistencia`** desde múltiples archivos Excel (`.xls` / `.xlsx`), estandarizar columnas, validar/normalizar RUT y fechas, y descargar un único Excel consolidado.

## 🚀 Ejecutar localmente

```bash
# Crear entorno (opcional)
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate

# Instalar dependencias
pip install -r requirements.txt

# Ejecutar
streamlit run consolida_asistencia_streamlit.py
