# 📦 Tablero Ejecutivo · Control de Salidas (Data Query)

link dashboard: https://dashboardpy-wx9jtjueeo4bqhhunaqzsc.streamlit.app/
Dashboard en **Streamlit** que lee **en vivo** el archivo `Data Query - estados.xlsx`
publicado en **SharePoint / Excel Web** y lo muestra como un tablero tipo "sala de
control" (tema oscuro).

link excel: https://diszasa365-my.sharepoint.com/:x:/g/personal/jaguilar_operadoralogistica_com/IQCoGmfBr55DS4E019WnIjZDAQx_bn-W09KKi1_nCjMx-hY?rtime=iv-BdxHv3kg

No hay que subir archivos manualmente: el dashboard descarga los datos desde el
link cada vez que se abre y se **refresca solo** cada pocos minutos. Soporta que se
agreguen nuevos pedidos (nuevas filas) sin cambiar el código.

---

## ¿Qué muestra?

- **KPIs**: total de documentos, **valor total**, despachados, sin cerrar y avance promedio.
- **Documentos por Estado** (barras).
- **Valor por Cuenta** (dona, con el total al centro).
- **Tendencia diaria**: documentos y valor por día (histórico real desde febrero 2026).
- **Alertas**: documentos sin cerrar, avance incompleto y estados sin catalogar.
- **Tabla de detalle** con búsqueda por **ID** o **Documento**, y descarga a CSV.
- **Filtros** en la barra lateral: Cuenta, Estado y rango de fechas.

---

## Estructura del repo

```
dashboard.py                    ← app principal
requirements.txt                ← dependencias
.streamlit/config.toml          ← tema oscuro
.streamlit/secrets.toml.example ← plantilla del link (NO subas el real)
.gitignore
README.md
```

---

## Cómo publicarlo (paso a paso)

### 1. Preparar el archivo en SharePoint
El archivo debe estar compartido como **"Cualquier persona con el enlace"**.
Se recomienda permiso de **solo lectura ("puede ver")**, no "puede editar",
para que nadie altere los datos por accidente.

### 2. Subir a GitHub
Sube estos archivos a tu repositorio (por ejemplo `Pedidos_prototipo_02`).
El `.gitignore` ya evita que se suba el archivo de secrets real.

### 3. Desplegar en Streamlit Community Cloud
1. Entra a [share.streamlit.io](https://share.streamlit.io) y conéctalo a tu repo.
2. En **Main file path** pon: `dashboard.py`
3. En **Settings → Secrets**, pega:
   ```toml
   DATA_QUERY_URL = "https://diszasa365-my.sharepoint.com/:x:/g/personal/jaguilar_operadoralogistica_com/IQCoGmfBr55DS4E019WnIjZDAQx_bn-W09KKi1_nCjMx-hY"
   ```
4. Deploy. Listo — tu papá y su equipo solo necesitan el link, sin instalar nada.

> Si no configuras el secret, el dashboard usa el link por defecto que ya viene
> escrito en `dashboard.py`. El secret solo sirve para poder cambiar el link
> después sin tocar el código.

---

## Correr en local (opcional)

```bash
pip install -r requirements.txt
streamlit run dashboard.py
```

---

## Notas técnicas

- La descarga se **cachea 5 minutos** (`TTL_SEGUNDOS`) para no saturar SharePoint;
  el botón **"🔄 Actualizar ahora"** limpia el caché al instante.
- Las fechas del Excel vienen como número serial y se convierten automáticamente.
- El campo **Cuenta** (1, 2, 3) se trata como categoría, no como número.
- Los estados `W/w (sin catálogo)` se unifican en una sola categoría "Sin catálogo".

## Siguiente fase (pendiente)

- Agregar las **Liquidaciones y Eventos** de las 3 empresas (Diszasa, Udisa, Disna)
  como una segunda sección del dashboard.
