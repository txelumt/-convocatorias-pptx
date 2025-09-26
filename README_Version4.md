# Datos Convocatorias (PPTX)

Este repositorio contiene:
- `generate_datos_convocatorias_pptx.py`: script que genera `datos_convocatorias.pptx` con el formulario editable.
- Workflow en `.github/workflows/build-pptx.yml` que construye el `.pptx` y lo comitea automáticamente.

## Uso rápido (GitHub Actions)
1. Ve a la pestaña **Actions** del repositorio.
2. Selecciona el workflow **Build PPTX** y pulsa **Run workflow**.
3. Al terminar, verás `datos_convocatorias.pptx` en la raíz del repo.

## Uso local
```bash
pip install python-pptx
python generate_datos_convocatorias_pptx.py
```
El archivo `datos_convocatorias.pptx` se generará en la raíz del proyecto.