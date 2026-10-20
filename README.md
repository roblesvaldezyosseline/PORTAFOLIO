# Joyería personalizada: del dato a la automatización
Caso de portafolio – Yosseline (Ing. de Sistemas, UCV)

## Contenido
- analisis_joyeria.py — calcula KPIs, embudo, RFM y control operativo
- metricas.json — salida del análisis (alimenta la página)
- crm_clientes.csv — 600 clientes con segmento RFM y última fecha especial
- n8n/*.json — 3 workflows importables (Workflows > Import from file)

## Ejecutar
pip install pandas openpyxl
python analisis_joyeria.py datos_joyeria.xlsx

## n8n
Reemplazar TU_GOOGLE_SHEET_ID y TU_PHONE_NUMBER_ID, y conectar credenciales
de Google Sheets y WhatsApp Cloud API (Header Auth: Bearer <token>).
