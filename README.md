# Tablero de concentrados de cartera

Aplicación Streamlit para generar concentrados LATAM y Presico por ruta y
localidad. Admite reportes `.xlsb` y `.xlsx` y permite descargar el resultado
filtrado en Excel o CSV.

## Archivos del repositorio

```text
.
├── app.py
├── estructura.csv
├── README.md
└── requirements.txt
```

`estructura.csv` debe conservar ese nombre y permanecer junto a `app.py`.
La aplicación no utiliza claves, contraseñas ni variables secretas.

## Ejecución local

Se recomienda Python 3.12.

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
streamlit run app.py
```

La aplicación acepta:

- Reportes diarios crudos cuya hoja normalmente se llama `Hoja1`.
- Libros formulados que ya contienen la hoja `Base`.

## Publicación en GitHub

1. Crea un repositorio nuevo en GitHub.
2. Sube a la raíz los cuatro archivos indicados arriba.
3. Confirma que `app.py`, `estructura.csv` y `requirements.txt` aparezcan en
   la rama que vas a desplegar.

## Despliegue en Streamlit Community Cloud

1. Entra a [share.streamlit.io](https://share.streamlit.io/) con GitHub.
2. Selecciona **Create app** y elige el repositorio y la rama.
3. En **Main file path**, escribe `app.py`.
4. En **Advanced settings**, selecciona Python **3.12**. No agregues secretos.
5. Pulsa **Deploy**.

Community Cloud instalará automáticamente las versiones fijadas en
`requirements.txt`. Si el repositorio es privado, autoriza a Streamlit para
acceder a él.

## Lógica reproducida

- Agrega la jerarquía comercial desde `estructura.csv` cuando el reporte es
  crudo.
- Reproduce las fórmulas de clientes, cartera, calidad, colocaciones y
  coordinadoras.
- Agrupa por jerarquía comercial, ruta e `id_y_localidad`.
- En Presico, una coordinadora es productiva con 21 o más clientes totales y
  calidad de al menos 60%; está en desarrollo con menos de 21 y calidad de al
  menos 60%; el resto es improductiva.
- En LATAM, los clientes totales incluyen cartera regular, FP y PP.
- Usa suma para clientes y cartera, y máximo para banderas de coordinadora.
- Calcula `Calidad = Cartera sin atrasos / Cartera Total`.

## Validación

La versión entregada fue comprobada con Python 3.12, lectura del catálogo,
procesamiento de ejemplos LATAM y Presico, generación de Excel y arranque de
la interfaz Streamlit.
