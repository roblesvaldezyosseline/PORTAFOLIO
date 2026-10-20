# Joyería personalizada: del dato a la automatización

Proyecto de portafolio de **Yosseline Robles Valdez**, estudiante de Ingeniería de Sistemas (6.º ciclo, UCV).

🔗 **Ver la página del proyecto:** https://roblesvaldezyosseline.github.io/PORTAFOLIO/

## Contenido

| Archivo | Qué es |
|---|---|
| `datos_joyeria.xlsx` | Base de práctica: 600 clientes, 20 productos, 800 pedidos y 1 500 consultas |
| `analisis_joyeria.py` | Script en Python que calcula los indicadores, el embudo y la segmentación RFM |
| `metricas.json` | Resultado del análisis (alimenta la página) |
| `crm_clientes.csv` | Clientes con su segmento RFM y su última fecha especial |
| `n8n/` | 3 automatizaciones importables en n8n |

## Cómo ejecutar el análisis

```
pip install pandas openpyxl
python analisis_joyeria.py datos_joyeria.xlsx
```

## Automatizaciones (n8n)

1. **Alerta de pedidos en riesgo:** avisa cada mañana qué pedidos llegarán tarde.
2. **Validar grabado:** revisa el texto antes de producir y pide confirmación al cliente.
3. **Recompra por fecha especial:** escribe al cliente 21 días antes de su fecha especial.

Para usarlas: en n8n, ir a *Workflows → Import from file*, reemplazar `TU_GOOGLE_SHEET_ID` y `TU_PHONE_NUMBER_ID` y conectar las credenciales de Google Sheets y WhatsApp.
