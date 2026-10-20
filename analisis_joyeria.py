"""
Análisis comercial y operativo – Joyería personalizada (caso de portafolio)
Autora: Yosseline

Entrada : datos_joyeria.xlsx (hojas Clientes, Productos, Pedidos, Consultas)
Salida  : metricas.json (alimenta el dashboard) + crm_clientes.csv (segmentación RFM)

Uso:  python analisis_joyeria.py datos_joyeria.xlsx
"""
import json
import sys
from pathlib import Path

import pandas as pd

SRC = Path(sys.argv[1] if len(sys.argv) > 1 else "datos_joyeria.xlsx")
OUT = Path(sys.argv[2] if len(sys.argv) > 2 else ".")
OUT.mkdir(parents=True, exist_ok=True)

xls = pd.ExcelFile(SRC)
cli = xls.parse("Clientes")
prod = xls.parse("Productos")
ped = xls.parse("Pedidos")
cons = xls.parse("Consultas")

# ------------------------------------------------------------------ 1. Calidad de datos
calidad = {
    "clientes": len(cli), "productos": len(prod), "pedidos": len(ped), "consultas": len(cons),
    "pedidos_sin_cliente_valido": int((~ped.id_cliente.isin(cli.id_cliente)).sum()),
    "pedidos_sin_producto_valido": int((~ped.id_producto.isin(prod.id_producto)).sum()),
    "totales_que_no_cuadran": int(
        ((ped.cantidad * ped.precio_unit_soles * (1 - ped.descuento_pct)).round(2) - ped.total_soles)
        .abs().gt(0.01).sum()),
    "nombres_repetidos_distinto_telefono": int(cli.nombre.duplicated().sum()),
    "telefonos_duplicados": int(cli.telefono.duplicated().sum()),
}

# ------------------------------------------------------------------ 2. Modelo de ventas
ped = ped.merge(prod[["id_producto", "producto", "material", "costo_soles"]], on="id_producto")
ped["mes"] = ped.fecha_pedido.dt.to_period("M").astype(str)
ped["costo_total"] = ped.cantidad * ped.costo_soles
ped["margen"] = ped.total_soles - ped.costo_total
validos = ped[ped.estado.isin(["Entregado", "En producción"])].copy()  # venta efectiva

entregados = ped[ped.estado == "Entregado"].copy()
entregados["dias_retraso"] = (entregados.fecha_entrega - entregados.fecha_prometida).dt.days
entregados["a_tiempo"] = entregados.dias_retraso <= 0

kpis = {
    "ventas_netas": round(validos.total_soles.sum(), 2),
    "margen_bruto": round(validos.margen.sum(), 2),
    "margen_pct": round(validos.margen.sum() / validos.total_soles.sum() * 100, 1),
    "pedidos_validos": int(len(validos)),
    "ticket_promedio": round(validos.total_soles.mean(), 2),
    "clientes_compradores": int(validos.id_cliente.nunique()),
    "tasa_cancelacion_pct": round((ped.estado == "Cancelado").mean() * 100, 1),
    "tasa_devolucion_pct": round((ped.estado == "Devuelto").mean() * 100, 1),
    "otif_pct": round(entregados.a_tiempo.mean() * 100, 1),
    "error_grabado_pct": round((ped.error_grabado == "Sí").mean() * 100, 1),
    "ventas_perdidas": round(ped[ped.estado.isin(["Cancelado", "Devuelto"])].total_soles.sum(), 2),
    "dias_produccion_prom": round(ped.dias_produccion.mean(), 1),
    "descuento_otorgado": round((validos.cantidad * validos.precio_unit_soles - validos.total_soles).sum(), 2),
}

mensual = (validos.groupby("mes")
           .agg(ventas=("total_soles", "sum"), pedidos=("id_pedido", "count"), margen=("margen", "sum"))
           .round(2).reset_index())
otif_mes = entregados.groupby("mes").a_tiempo.mean().mul(100).round(1)
mensual["otif"] = mensual.mes.map(otif_mes)

por_producto = (validos.groupby("producto")
                .agg(ventas=("total_soles", "sum"), unidades=("cantidad", "sum"), margen=("margen", "sum"))
                .sort_values("ventas", ascending=False).round(2).reset_index())
por_material = validos.groupby("material").agg(ventas=("total_soles", "sum"), margen=("margen", "sum")).round(2)
por_material["margen_pct"] = (por_material.margen / por_material.ventas * 100).round(1)

por_canal_venta = (validos.groupby("canal_venta").total_soles.sum().sort_values(ascending=False).round(2))
motivos = validos.groupby("motivo_compra").total_soles.sum().sort_values(ascending=False).round(2)
pago = validos.metodo_pago.value_counts()

# Causa raíz: ¿el error de grabado y el retraso explican cancelaciones/devoluciones?
ped["retraso"] = (ped.fecha_entrega - ped.fecha_prometida).dt.days > 0
causa = {
    "devolucion_con_error_pct": round(ped[ped.error_grabado == "Sí"].estado.eq("Devuelto").mean() * 100, 1),
    "devolucion_sin_error_pct": round(ped[ped.error_grabado == "No"].estado.eq("Devuelto").mean() * 100, 1),
    "cancelacion_con_error_pct": round(ped[ped.error_grabado == "Sí"].estado.eq("Cancelado").mean() * 100, 1),
    "cancelacion_sin_error_pct": round(ped[ped.error_grabado == "No"].estado.eq("Cancelado").mean() * 100, 1),
    "otif_por_dias_produccion": entregados.groupby("dias_produccion").a_tiempo.mean().mul(100).round(1).to_dict(),
    "pedidos_por_dias_produccion": entregados.dias_produccion.value_counts().sort_index().to_dict(),
    "envio_dias_prom": round(((entregados.fecha_entrega - entregados.fecha_pedido).dt.days - entregados.dias_produccion).mean(), 1),
    "entregas_tarde": int((entregados.dias_retraso > 0).sum()),
    "retraso_prom_dias": round(entregados[entregados.dias_retraso > 0].dias_retraso.mean(), 1),
    "con_fecha_especial": int(entregados.fecha_especial.notna().sum()),
    "llegaron_despues_fecha_especial": int((entregados.fecha_entrega > entregados.fecha_especial).sum()),
    "error_por_texto":ped.groupby("texto_grabado").error_grabado.apply(lambda s: round((s == "Sí").mean() * 100, 1)).sort_values(ascending=False).to_dict(),
}

# ------------------------------------------------------------------ 3. Embudo de consultas (CRM)
cons["convirtio"] = cons.convirtio_en_venta == "Sí"
bins = [-1, 15, 60, 240, 10_000]
labels = ["≤15 min", "16–60 min", "1–4 h", "> 4 h"]
cons["tramo_resp"] = pd.cut(cons.minutos_primera_respuesta, bins=bins, labels=labels)
embudo = {
    "consultas": int(len(cons)),
    "convertidas": int(cons.convirtio.sum()),
    "conversion_pct": round(cons.convirtio.mean() * 100, 1),
    "mediana_min_respuesta": float(cons.minutos_primera_respuesta.median()),
    "por_canal": cons.groupby("canal").agg(consultas=("id_consulta", "count"), conversion=("convirtio", "mean"),
                                           resp_mediana=("minutos_primera_respuesta", "median"))
                     .assign(conversion=lambda d: (d.conversion * 100).round(1)).sort_values("consultas", ascending=False)
                     .reset_index().to_dict("records"),
    "por_tipo": cons.groupby("tipo_consulta").agg(consultas=("id_consulta", "count"), conversion=("convirtio", "mean"))
                    .assign(conversion=lambda d: (d.conversion * 100).round(1)).sort_values("consultas", ascending=False)
                    .reset_index().to_dict("records"),
    "por_tiempo_respuesta": cons.groupby("tramo_resp", observed=False).agg(consultas=("id_consulta", "count"),
                                                                          conversion=("convirtio", "mean"))
                                .assign(conversion=lambda d: (d.conversion * 100).round(1)).reset_index()
                                .astype({"tramo_resp": str}).to_dict("records"),
    "consultas_sin_respuesta_1h_pct": round((cons.minutos_primera_respuesta > 60).mean() * 100, 1),
}

# ------------------------------------------------------------------ 4. Segmentación RFM (estructura CRM)
corte = validos.fecha_pedido.max() + pd.Timedelta(days=1)
rfm = (validos.groupby("id_cliente")
       .agg(ultima=("fecha_pedido", "max"), frecuencia=("id_pedido", "count"), monto=("total_soles", "sum")))
rfm["recencia_dias"] = (corte - rfm.ultima).dt.days
rfm["R"] = pd.qcut(rfm.recencia_dias, 4, labels=[4, 3, 2, 1]).astype(int)
rfm["F"] = rfm.frecuencia.clip(upper=4)
rfm["M"] = pd.qcut(rfm.monto.rank(method="first"), 4, labels=[1, 2, 3, 4]).astype(int)


def segmento(r):
    if r.R >= 3 and r.F >= 3:
        return "Campeones"
    if r.F >= 2 and r.R >= 2:
        return "Leales"
    if r.R >= 3 and r.F == 1:
        return "Nuevos"
    if r.R <= 2 and (r.F >= 2 or r.M >= 3):
        return "En riesgo"
    return "Dormidos"


rfm["segmento"] = rfm.apply(segmento, axis=1)
crm = cli.merge(rfm.reset_index(), on="id_cliente", how="left")
crm["segmento"] = crm.segmento.fillna("Prospecto (sin compra)")
crm["frecuencia"] = crm.frecuencia.fillna(0).astype(int)
crm["monto"] = crm.monto.fillna(0).round(2)

# Próxima fecha especial por cliente → disparador de campaña
fe = ped.dropna(subset=["fecha_especial"]).groupby("id_cliente").fecha_especial.max()
crm["ultima_fecha_especial"] = crm.id_cliente.map(fe)
crm.drop(columns=["ultima"]).to_csv(OUT / "crm_clientes.csv", index=False, encoding="utf-8-sig")

seg = (crm.groupby("segmento").agg(clientes=("id_cliente", "count"), monto=("monto", "sum"),
                                   frecuencia=("frecuencia", "mean"))
       .round(2).sort_values("monto", ascending=False).reset_index())
seg["pct_ventas"] = (seg.monto / seg.monto.sum() * 100).round(1)

recompra = {
    "clientes_con_2mas_pct": round((rfm.frecuencia >= 2).mean() * 100, 1),
    "ventas_de_recurrentes_pct": round(rfm[rfm.frecuencia >= 2].monto.sum() / rfm.monto.sum() * 100, 1),
    "registrados_sin_compra": int((crm.frecuencia == 0).sum()),
}
origen = crm.groupby("canal_origen").agg(clientes=("id_cliente", "count"), compradores=("frecuencia", lambda s: (s > 0).sum()),
                                         monto=("monto", "sum")).round(2)
origen["valor_por_cliente"] = (origen.monto / origen.clientes).round(2)

# ------------------------------------------------------------------ 5. Control operativo: pedidos abiertos
abiertos = ped[ped.estado == "En producción"][["id_pedido", "id_cliente", "producto", "fecha_pedido",
                                               "fecha_prometida", "dias_produccion", "fecha_especial"]]

out = {
    "calidad": calidad, "kpis": kpis, "mensual": mensual.to_dict("records"),
    "por_producto": por_producto.to_dict("records"),
    "por_material": por_material.reset_index().to_dict("records"),
    "por_canal_venta": por_canal_venta.to_dict(), "motivos": motivos.to_dict(), "pago": pago.to_dict(),
    "causa": causa, "embudo": embudo, "segmentos": seg.to_dict("records"), "recompra": recompra,
    "origen": origen.reset_index().to_dict("records"),
    "abiertos": abiertos.astype(str).to_dict("records"),
    "periodo": [str(ped.fecha_pedido.min().date()), str(ped.fecha_pedido.max().date())],
}
(OUT / "metricas.json").write_text(json.dumps(out, ensure_ascii=False, indent=1, default=str), encoding="utf-8")
print(json.dumps({k: out[k] for k in ["calidad", "kpis", "recompra"]}, ensure_ascii=False, indent=1))
