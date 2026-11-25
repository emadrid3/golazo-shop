#!/bin/bash
# macOS: doble clic (la primera vez: clic derecho → Abrir).
# Genera data/catalogo.json y las fotos web en img/ a partir de datos/*.xlsx e imagenes/.
cd "$(dirname "$0")"
[ -d .venv ] || python3 -m venv .venv
source .venv/bin/activate
pip install --quiet openpyxl pillow
python herramientas/actualizar_catalogo.py --optimizar --ancho-mini 0 --ancho-galeria 720 --calidad 70
echo "Listo. Ahora en GitHub Desktop: Commit to main → Push origin"
read -p "Presione Enter para cerrar"
