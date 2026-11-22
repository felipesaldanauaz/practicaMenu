import streamlit as st
from datetime import datetime, time
from io import BytesIO
from pathlib import Path
from openpyxl import load_workbook

DIAS = ['Lunes', 'Martes', 'Miercoles', 'Jueves', 'Viernes', 'Sabado', 'Domingo']
CAMPOS_PLAN = ['dia', 'hora_inicio', 'hora_fin', 'tiempo_comida', 'menu', 'receta_id']
CAMPOS_RECETAS = ['receta_id', 'nombre', 'ingredientes_cantidades', 'preparacion',
                  'video_url', 'estado', 'fuente_preparacion']
BASE = Path(__file__).resolve().parent


def texto(valor):
    """Convierte celdas vacías en texto vacío."""
    return '' if valor is None else str(valor).strip()


def leerHoja(libro, nombre, campos):
    """Valida encabezados y transforma filas en diccionarios."""
    if nombre not in libro.sheetnames:
        raise ValueError(f'Falta la hoja {nombre}.')
    filas = list(libro[nombre].iter_rows(values_only=True))
    if not filas:
        raise ValueError(f'La hoja {nombre} está vacía.')
    encabezados = [texto(celda) for celda in filas[0]]
    if any(encabezados.count(campo) != 1 for campo in campos):
        raise ValueError(f'Revisa los encabezados de {nombre}: {", ".join(campos)}.')

    registros = []
    for fila in filas[1:]:
        if not any(texto(celda) for celda in fila):
            continue
        registro = {campo: texto(fila[encabezados.index(campo)]) for campo in campos}
        registros.append(registro)
    if not registros:
        raise ValueError(f'La hoja {nombre} no contiene registros.')
    return registros


def cargar_plan(contenido):
    """Abre el libro de Excel desde bytes y lee sus dos tablas."""
    libro = load_workbook(BytesIO(contenido), read_only=True, data_only=True)
    try:
        plan = leerHoja(libro, 'Plan', CAMPOS_PLAN)
        recetas = leerHoja(libro, 'Recetas', CAMPOS_RECETAS)
    finally:
        libro.close()
    return plan, recetas


def mostrar_menu(fila, recetas):
    """Muestra una comida y los datos de la receta asociada."""
    st.subheader(fila['menu'] or fila['tiempo_comida'] or 'Comida')
    st.caption(f"{fila['hora_inicio']} - {fila['hora_fin']} | {fila['tiempo_comida']}")
    receta = recetas.get(fila['receta_id'])
    if receta is None:
        st.warning(f"No se encontró la receta {fila['receta_id']}.")
        return
    st.write(f"**Receta:** {receta['nombre']}")
    st.write(f"**Ingredientes y cantidades:** {receta['ingredientes_cantidades']}")
    if receta['preparacion']:
        st.write(f"**Preparación:** {receta['preparacion']}")
    if receta['estado']:
        st.caption(f"Estado: {receta['estado']}")
    if receta['fuente_preparacion']:
        st.write(f"**Fuente:** {receta['fuente_preparacion']}")
    if receta['video_url'].startswith(('http://', 'https://')):
        st.link_button('Ver video', receta['video_url'])


def minutos(valor):
    """Acepta HH:MM y objetos time de Excel; devuelve minutos desde medianoche."""
    if isinstance(valor, (datetime, time)):
        return valor.hour * 60 + valor.minute
    cadena = texto(valor)
    partes = cadena.split(':')
    if len(partes) < 2:
        raise ValueError(f'Formato de hora no válido: {cadena!r}')
    horas, mins = int(partes[0]), int(partes[1])
    if not (0 <= horas <= 23 and 0 <= mins <= 59):
        raise ValueError(f'Hora fuera de rango: {cadena!r}')
    return horas * 60 + mins


def main():
    st.set_page_config(page_title='Mi Menú de hoy', page_icon='🍽️', layout='centered')
    st.title('Mi Menú de hoy - Felipe Saldaña')
    st.write('Consulta los horarios, alimentos y cantidades de tu documento.')
    dia = st.selectbox('Día', DIAS)
    st.write('Día elegido:', dia)

    ruta = BASE / 'Mi_menu_de_hoy.xlsx'
    archivo = st.file_uploader('Cargar otro archivo Excel (opcional)', type=['xlsx'])
    if archivo is not None:
        contenido = archivo.getvalue()
    elif ruta.exists():
        contenido = ruta.read_bytes()
        st.caption(f'Archivo utilizado: {ruta.name}')
    else:
        st.info('Coloca Mi_menu_de_hoy.xlsx junto a app.py o súbelo aquí.')
        return

    try:
        plan, recetas = cargar_plan(contenido)
        # Lista por comprensión: filtrar el día elegido.
        comidas = [fila for fila in plan if fila['dia'].casefold() == dia.casefold()]
        # Diccionario por comprensión: relacionar identificadores con recetas.
        recetas_por_id = {receta['receta_id']: receta for receta in recetas}
        # Orden cronológico por hora de inicio.
        comidas.sort(key=lambda fila: minutos(fila['hora_inicio']))
    except (ValueError, TypeError, KeyError, OSError) as error:
        st.error(f'No se pudieron procesar los datos: {error}')
        return

    if not comidas:
        st.warning(f'No hay comidas registradas para {dia}.')
        return
    st.header(f'Menú del {dia}')
    for fila in comidas:
        mostrar_menu(fila, recetas_por_id)
        st.divider()


if __name__ == '__main__':
    main()
