# Dashboard Comercial

Panel de ventas interactivo hecho con **Streamlit + Plotly**. Está pensado para
que cualquier persona del equipo comercial abra un enlace, vea los indicadores y
filtre la información **sin escribir código**.

## Qué incluye

- **Indicadores**: ingresos, unidades vendidas, ticket promedio y mejor vendedor.
- **Filtros**: periodo, producto y vendedor, con botón para limpiar.
- **Gráficos interactivos**: evolución diaria, participación por producto,
  ingresos por producto, ingresos por vendedor y una vista 3D de las operaciones.
- **Tabla de operaciones** ordenable y con búsqueda.
- **Descargas**: reporte en Excel y CSV con los datos filtrados.
- **Carga de datos** desde la interfaz (CSV o Excel) sin tocar código.
- **Diseño minimalista y responsive** (escritorio y móvil).

## Ejecutar en local

```bash
pip install -r requirements.txt
streamlit run analisis_ventas.py
```

Se abre en `http://localhost:8501`. La app usa `ventas.csv` si tiene contenido;
si el archivo está vacío, muestra los datos de ejemplo.

## Abrir desde Google Colab

Instala dependencias, levanta Streamlit y expón el puerto con un túnel. Por
ejemplo, con Cloudflare Tunnel:

```python
!pip install -q -r requirements.txt

!wget -q https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-linux-amd64 -O cloudflared
!chmod +x cloudflared

!streamlit run analisis_ventas.py &> /content/streamlit.log &
!./cloudflared tunnel --url http://localhost:8501
```

Cloudflare imprime una URL `https://....trycloudflare.com`: ese es el enlace que
puedes compartir. Ten en cuenta que el enlace vive mientras la sesión de Colab
siga activa.

> Alternativa con localtunnel: `!npx localtunnel --port 8501`. Si aparece una
> pantalla de contraseña, usa la IP pública que muestre el propio localtunnel.

## Enlace permanente (recomendado para el gerente)

Para un enlace estable, publica la app gratis en **Streamlit Community Cloud**:

1. Sube el proyecto a un repositorio de GitHub (incluye `requirements.txt`).
2. Entra en <https://share.streamlit.io> y elige **New app**.
3. Selecciona el repositorio, la rama y `analisis_ventas.py` como archivo principal.
4. Pulsa **Deploy** y comparte la URL resultante.

Streamlit Cloud instala automáticamente las dependencias de `requirements.txt`.

## Formato de los datos

El archivo debe tener, al menos, una columna de ventas. El resto se detecta de
forma flexible:

| Columna     | Nombres aceptados (ejemplos)                     | Obligatoria |
|-------------|--------------------------------------------------|-------------|
| Fecha       | Fecha, Date, Día, Periodo                        | No          |
| Producto    | Producto, Product, Artículo, Descripción          | No          |
| Vendedor    | Vendedor, Seller, Agente, Comercial               | No          |
| Ventas      | Ventas, Ingresos, Total, Monto, Importe           | Sí*         |
| Cantidad    | Cantidad, Unidades, Qty, Units                    | No          |
| Precio      | Precio, Precio unitario, Price                    | No          |

\* Si no existe una columna de ventas pero sí `Precio` y `Cantidad`, la app
calcula los ingresos automáticamente.

Se aceptan CSV con coma, punto y coma o tabulador, y archivos Excel (`.xlsx`).
Los valores con formato `$ 3.500`, `1.234,56` o `3500` se interpretan como
números. Si falta la fecha, el panel sigue funcionando, pero no muestra la
tendencia diaria.

### Actualizar los datos

- **Desde la app**: botón **Cargar datos** (arriba a la derecha). El cambio
  aplica a la sesión actual del navegador.
- **De forma permanente**: reemplaza `ventas.csv` en el proyecto y vuelve a
  desplegar o recarga la app.

## Estructura del proyecto

```
analisis_ventas.py       Aplicación Streamlit
ventas.csv               Datos de ejemplo / datos reales
requirements.txt         Dependencias
.streamlit/config.toml   Tema visual claro
CHANGELOG.md             Historial de versiones
```

## Flujo de ramas

- `main`: versión estable que se comparte y despliega.
- `develop`: integración de mejoras y cambios en curso.

Para trabajar en una mejora:

```bash
git switch develop
# editar archivos
git add .
git commit -m "feat: descripcion del cambio"
git switch main
git merge --no-ff develop
git push origin main
git push origin develop
```

Los cambios se prueban en `develop` y, cuando están listos, se integran en
`main` con un commit de merge.
