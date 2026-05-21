# 🔧 Extractor AUDATEX – Carrocería

Aplicación web para extraer partidas de **Mano de Obra Hojal/Mecánica** y **Pintura de Carrocería** de reportes PDF generados por AUDATEX / GNP Seguros.

## ¿Qué hace?

- Sube uno o varios PDFs AUDATEX
- Detecta automáticamente el **N° de Orden** (Referencia Interna del PDF)
- Extrae todas las partidas de M.O. y Pintura con su precio
- Exporta a **Excel** (una sola hoja con columnas: N° Orden | Sección | NR/Pos. | Descripción | Precio)
- Exporta a **CSV**

## Estructura del Excel generado

| N° ORDEN | SECCIÓN | NR / POS. | TRABAJO / DESCRIPCIÓN | PRECIO ($) |
|---|---|---|---|---|
| 3228 | Mano de Obra Hojal/Mecánica | SN | CAMARA DEL.:D+M TRAB.ADIC. | $52.00 |
| 3228 | Pintura de Carrocería | 1481 | PUERTA DEL.I. PINTURA SUSTITUCIÓN | $691.60 |

## Cómo correr localmente

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy en Streamlit Cloud

1. Sube este repositorio a GitHub
2. Ve a [share.streamlit.io](https://share.streamlit.io)
3. Conecta tu cuenta de GitHub
4. Selecciona este repositorio y `app.py` como archivo principal
5. Haz clic en **Deploy** — en 2 minutos tienes tu URL pública
