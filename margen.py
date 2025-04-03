import pandas as pd
from datetime import datetime
from fastapi import APIRouter, UploadFile, File, HTTPException, Response
import io
import logging
from typing import Annotated # Use Annotated for File uploads

# Configurar logging
logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO)

router = APIRouter()

def procesar_margenes(
    config_bytes: bytes,
    ventas_me_bytes: bytes,
    ventas_mn_bytes: bytes,
    kardex_bytes: bytes
):
    """
    Procesa los archivos de ventas y costos para calcular márgenes.

    Args:
        config_bytes: Bytes del archivo Articulos.xlsx.
        ventas_me_bytes: Bytes del archivo Ventas_ME.xlsx.
        ventas_mn_bytes: Bytes del archivo Ventas_MN.xlsx.
        kardex_bytes: Bytes del archivo Kardex.xlsx.

    Returns:
        Bytes del archivo Excel resultante.
    """
    try:
        logger.info("Iniciando carga de archivos en DataFrames.")
        # Cargar archivos desde bytes
        config_df = pd.read_excel(io.BytesIO(config_bytes), header=4)[['Código', 'Descripción de Familia']]
        config_df.columns = ['Artículo', 'Cuenta Contable']
        logger.info("Archivo Articulos.xlsx cargado.")

        me_df = pd.read_excel(io.BytesIO(ventas_me_bytes))
        logger.info("Archivo Ventas_ME.xlsx cargado.")
        mn_df = pd.read_excel(io.BytesIO(ventas_mn_bytes))
        logger.info("Archivo Ventas_MN.xlsx cargado.")

        # Seleccionar columnas relevantes y renombrar
        cols = ['Fecha', 'T. Doc.', 'Número', 'Cliente', 'Artículo', 'Descripción', 'Cantidad', 'Und.', 'Valor Unit', 'Valor Vta.']
        me_df = me_df[cols].rename(columns={'Valor Vta.': 'Valor Vta. ME'})
        mn_df = mn_df[cols].rename(columns={'Valor Vta.': 'Valor Vta. MN'})
        logger.info("Columnas seleccionadas y renombradas para ventas.")

        # Unir ME y MN
        # Asegurarse que las columnas clave sean del mismo tipo antes de unir
        key_cols = ['Fecha', 'T. Doc.', 'Número', 'Cliente', 'Artículo', 'Descripción', 'Cantidad', 'Und.']
        for col in key_cols:
            if col == 'Fecha':
                me_df[col] = pd.to_datetime(me_df[col])
                mn_df[col] = pd.to_datetime(mn_df[col])
            else:
                 # Intentar convertir a string como método genérico, ajustar si es necesario
                 me_df[col] = me_df[col].astype(str)
                 mn_df[col] = mn_df[col].astype(str)

        ventas_df = pd.merge(me_df, mn_df, on=key_cols, how='inner') # Usar inner join para asegurar coincidencia exacta
        logger.info("DataFrames de ventas ME y MN unidos.")

        # Normalizar códigos de artículo antes de unir con configuración y costos
        ventas_df['Artículo'] = ventas_df['Artículo'].astype(str).str.upper().str.strip().str.replace(r'\.0$', '', regex=True)
        config_df['Artículo'] = config_df['Artículo'].astype(str).str.upper().str.strip().str.replace(r'\.0$', '', regex=True)

        # Agregar cuenta contable
        ventas_df = pd.merge(ventas_df, config_df, on='Artículo', how='left')
        logger.info("Cuenta contable agregada.")

        # Cargar archivo de costos y procesar
        # Usar usecols para cargar solo las columnas necesarias eficientemente
        costos_df = pd.read_excel(io.BytesIO(kardex_bytes), header=None, usecols=[7, 22])
        costos_df.columns = ['Artículo', 'Costo unit Venta MN']
        costos_df['Costo unit Venta MN'] = pd.to_numeric(costos_df['Costo unit Venta MN'], errors='coerce')
        costos_df = costos_df.dropna(subset=['Costo unit Venta MN']) # Eliminar filas sin costo numérico
        costos_df = costos_df[costos_df['Costo unit Venta MN'] > 0]
        logger.info("Archivo Kardex.xlsx cargado y filtrado.")

        # Normalizar códigos de artículo en costos
        costos_df['Artículo'] = costos_df['Artículo'].astype(str).str.upper().str.strip().str.replace(r'\.0$', '', regex=True)

        # Tomar solo el primer costo válido por artículo
        costos_df = costos_df.groupby('Artículo', as_index=False).first()
        logger.info("Primer costo válido por artículo seleccionado.")

        # Unir costos a ventas
        ventas_df = pd.merge(ventas_df, costos_df, on='Artículo', how='left')
        logger.info("Costos unidos al DataFrame de ventas.")

        # Llenar costos faltantes con 0 para evitar errores en cálculos
        ventas_df['Costo unit Venta MN'] = ventas_df['Costo unit Venta MN'].fillna(0)

        # --- Agrupación Eliminada Temporalmente para Revisión ---
        # logger.info("Agrupando para evitar duplicados...")
        # group_cols = [
        #     'Cuenta Contable', 'Fecha', 'T. Doc.', 'Número', 'Cliente', 'Artículo',
        #     'Descripción', 'Und.', 'Valor Unit', # Añadido Descripción y Valor Unit si son relevantes
        #     'Valor Vta. ME', 'Valor Vta. MN', 'Costo unit Venta MN'
        # ]
        # # Asegurarse de que las columnas para agrupar no tengan NaNs donde no deberían
        # # (Ej: 'Cuenta Contable' puede ser NaN si el 'Artículo' no está en config_df)
        # ventas_df['Cuenta Contable'] = ventas_df['Cuenta Contable'].fillna('Desconocida')

        # # Revisar tipos de datos antes de agrupar
        # # logger.info(ventas_df[group_cols].dtypes) # Descomentar para depurar tipos

        # ventas_df = ventas_df.groupby(group_cols, as_index=False, dropna=False).agg({'Cantidad': 'sum'})
        # logger.info("Agrupación completada.")
        # --- Fin Agrupación Eliminada ---

        # Asegurar que QTY (Cantidad) es numérico
        ventas_df['Cantidad'] = pd.to_numeric(ventas_df['Cantidad'], errors='coerce').fillna(0)

        # Calcular métricas
        logger.info("Calculando métricas de costo y margen.")
        ventas_df['Costo Total Venta MN'] = ventas_df['Cantidad'] * ventas_df['Costo unit Venta MN']
        ventas_df['Margen Bruto MN'] = ventas_df['Valor Vta. MN'] - ventas_df['Costo Total Venta MN']

        # Evitar división por cero o NaN en el cálculo del porcentaje
        ventas_df['% Margen Bruto'] = 0.0 # Inicializar
        mask = ventas_df['Valor Vta. MN'].notna() & (ventas_df['Valor Vta. MN'] != 0)
        ventas_df.loc[mask, '% Margen Bruto'] = (ventas_df.loc[mask, 'Margen Bruto MN'] / ventas_df.loc[mask, 'Valor Vta. MN']) * 100

        # Formatear porcentaje como string
        ventas_df['%'] = ventas_df['% Margen Bruto'].map(lambda x: f"{x:.2f}%" if pd.notna(x) else "0.00%")

        # Redondear valores numéricos
        numeric_cols = ['Valor Vta. ME', 'Valor Vta. MN', 'Costo unit Venta MN',
                        'Costo Total Venta MN', 'Margen Bruto MN']
        for col in numeric_cols:
             ventas_df[col] = ventas_df[col].round(2)

        # Seleccionar y ordenar columnas finales
        final_cols = [
            'Cuenta Contable', 'Fecha', 'T. Doc.', 'Número', 'Cliente',
            'Artículo', 'Descripción', 'Und.', 'Cantidad', # Cambiar QTY por Cantidad si es necesario
            'Valor Vta. ME', 'Valor Vta. MN', 'Costo unit Venta MN',
            'Costo Total Venta MN', 'Margen Bruto MN', '%'
        ]
        # Añadir Descripción de vuelta si se quitó en la agrupación
        # if 'Descripción' not in ventas_df.columns:
             # Si se agrupó sin Descripción, se necesita decidir cómo reincorporarla
             # Por ahora, la omitimos si no está tras la agrupación
             # final_cols.remove('Descripción') # O manejarla de otra forma

        # Renombrar 'Cantidad' a 'QTY' al final si es necesario
        ventas_df.rename(columns={'Cantidad': 'QTY'}, inplace=True)
        final_cols[final_cols.index('Cantidad')] = 'QTY' # Actualizar nombre en la lista

        # Filtrar columnas inexistentes antes de seleccionar
        final_cols = [col for col in final_cols if col in ventas_df.columns]
        ventas_final_df = ventas_df[final_cols]
        logger.info("Columnas finales seleccionadas y ordenadas.")

        # Exportar a Excel en memoria
        output_buffer = io.BytesIO()
        with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
            ventas_final_df.to_excel(writer, index=False, sheet_name='Reporte Margen')
        output_buffer.seek(0)
        logger.info("DataFrame final exportado a buffer de Excel.")

        return output_buffer.getvalue()

    except FileNotFoundError as e:
        logger.error(f"Error: Archivo no encontrado - {e}", exc_info=True)
        raise HTTPException(status_code=404, detail=f"Archivo no encontrado: {e.filename}")
    except pd.errors.EmptyDataError as e:
        logger.error(f"Error: Uno de los archivos Excel está vacío o tiene un formato incorrecto - {e}", exc_info=True)
        raise HTTPException(status_code=400, detail=f"Archivo vacío o formato incorrecto: {e}")
    except KeyError as e:
        logger.error(f"Error: Columna esperada no encontrada en uno de los archivos - {e}", exc_info=True)
        raise HTTPException(status_code=400, detail=f"Columna faltante o incorrecta: {e}")
    except ValueError as e:
         logger.error(f"Error de valor durante el procesamiento: {e}", exc_info=True)
         raise HTTPException(status_code=400, detail=f"Error en los datos: {e}")
    except Exception as e:
        logger.error(f"Error inesperado durante el procesamiento: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=f"Error interno del servidor: {e}")

@router.post("/calcular-margen", response_class=Response)
async def calcular_margen_endpoint(
    articulos_file: Annotated[UploadFile, File(description="Archivo de configuración Articulos.xlsx")],
    ventas_me_file: Annotated[UploadFile, File(description="Archivo de Ventas_ME.xlsx")],
    ventas_mn_file: Annotated[UploadFile, File(description="Archivo de Ventas_MN.xlsx")],
    kardex_file: Annotated[UploadFile, File(description="Archivo de costos Kardex.xlsx")]
):
    """
    Endpoint para cargar archivos, calcular márgenes y devolver un Excel.
    """
    logger.info("Recibida solicitud para /calcular-margen")
    try:
        # Leer contenido de los archivos
        config_bytes = await articulos_file.read()
        ventas_me_bytes = await ventas_me_file.read()
        ventas_mn_bytes = await ventas_mn_file.read()
        kardex_bytes = await kardex_file.read()
        logger.info("Archivos leídos en memoria.")

        # Procesar los datos
        output_excel_bytes = procesar_margenes(
            config_bytes, ventas_me_bytes, ventas_mn_bytes, kardex_bytes
        )

        # Crear nombre de archivo para descarga
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        file_name = f"reporte_final_ventas_{timestamp}.xlsx"
        logger.info(f"Procesamiento completado. Enviando archivo: {file_name}")

        # Devolver el archivo Excel como respuesta
        return Response(
            content=output_excel_bytes,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={
                "Content-Disposition": f"attachment; filename=\"{file_name}\"",
                # Exponer Content-Disposition para que el frontend pueda leer el nombre del archivo
                "Access-Control-Expose-Headers": "Content-Disposition"
            }
        )

    except HTTPException as http_exc:
        # Re-lanzar excepciones HTTP generadas en procesar_margenes
        raise http_exc
    except Exception as e:
        logger.error(f"Error inesperado en el endpoint /calcular-margen: {e}", exc_info=True)
        # Devolver error genérico 500
        raise HTTPException(status_code=500, detail=f"Error interno del servidor al procesar la solicitud: {e}")
    finally:
        # Cerrar archivos si es necesario (FastAPI lo maneja usualmente)
        await articulos_file.close()
        await ventas_me_file.close()
        await ventas_mn_file.close()
        await kardex_file.close()
        logger.info("Archivos cerrados.")

# Puedes añadir más endpoints a este router si es necesario
# Ejemplo: router.get("/status") ... 
