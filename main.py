# -*- coding: utf-8 -*-
"""
Este script descarga el historial de tasas de cambio de USD a COP y de EUR a COP
para un rango de fechas especificado por el usuario y guarda los valores de
cierre en un archivo de Excel.
"""

import yfinance as yf
import pandas as pd
from datetime import datetime, date

def solicitar_fecha(mensaje_prompt):
    """
    Solicita al usuario una fecha en formato dd/mm/yyyy y la valida.
    
    Args:
        mensaje_prompt (str): El mensaje para mostrar al usuario.
        
    Returns:
        datetime.date: La fecha validada como un objeto date.
    """
    while True:
        fecha_str = input(mensaje_prompt)
        try:
            # Convierte el string a un objeto datetime y luego lo convierte a date
            return datetime.strptime(fecha_str, '%d/%m/%Y').date()
        except ValueError:
            print("Formato de fecha incorrecto. Por favor, use el formato dd/mm/yyyy.")

def obtener_y_guardar_historial_conversion():
    """
    Función principal para obtener el historial de tasas de cambio
    y guardarlo en un archivo de Excel.
    """
    print("Este script descarga datos de divisas para un rango de fechas específico.")
    
    # Solicita las fechas al usuario
    fecha_inicio = solicitar_fecha("Ingrese la fecha de inicio (dd/mm/yyyy): ")
    fecha_fin = solicitar_fecha("Ingrese la fecha de fin (dd/mm/yyyy): ")

    # Valida que la fecha de inicio no sea posterior a la de fin
    if fecha_inicio > fecha_fin:
        print("\nAdvertencia: La fecha de inicio era posterior a la fecha de fin. Se han intercambiado.")
        fecha_inicio, fecha_fin = fecha_fin, fecha_inicio

    print(f"\nIniciando la descarga del historial desde {fecha_inicio.strftime('%d/%m/%Y')} hasta {fecha_fin.strftime('%d/%m/%Y')}...")

    # Define los tickers para las conversiones
    tickers = ["USDCOP=X", "EURCOP=X"]

    try:
        # Descargamos los datos históricos para el rango de fechas especificado.
        # yfinance incluye la fecha de inicio pero excluye la de fin, por lo que no es necesario sumar un día.
        data = yf.download(tickers, start=fecha_inicio, end=fecha_fin, progress=False)

        if data.empty:
            print("\nNo se pudieron obtener datos para el período especificado.")
            print("El mercado podría haber estado cerrado o hay un problema de conexión.")
            return

        # Extraemos únicamente la columna de precios de cierre ('Close') y creamos una copia explícita
        historial_cierre = data['Close'].copy()
        
        # Eliminamos filas que no tengan datos (NaN en alguna de las columnas de tickers)
        historial_cierre.dropna(inplace=True)

        # Si después de eliminar NaNs el dataframe está vacío, no continuar.
        if historial_cierre.empty:
            print("\nNo hay datos de cierre disponibles para el período después de limpiar NaNs.")
            return
        
        # Renombramos las columnas
        # No es estrictamente necesario usar .loc para rename inplace, pero mantenemos consistencia con las asignaciones.
        historial_cierre.rename(columns={'USDCOP=X': 'Valor Cierre USD/COP',
                                         'EURCOP=X': 'Valor Cierre EUR/COP'}, inplace=True)

        # Formateamos los valores a dos decimales usando .loc para asegurar que se modifica el DataFrame original
        # y para evitar SettingWithCopyWarning.
        historial_cierre.loc[:, 'Valor Cierre USD/COP'] = historial_cierre['Valor Cierre USD/COP'].round(2)
        historial_cierre.loc[:, 'Valor Cierre EUR/COP'] = historial_cierre['Valor Cierre EUR/COP'].round(2)

        # Nombre del archivo de salida personalizado con las fechas
        nombre_archivo = f"historial_divisas_cop_{fecha_inicio.strftime('%Y%m%d')}_a_{fecha_fin.strftime('%Y%m%d')}.xlsx"

        # Guardamos el DataFrame en un archivo de Excel.
        try:
            historial_cierre.to_excel(nombre_archivo, sheet_name='HistorialTasasDeCambio')
            print(f"\n¡Éxito! El historial de datos se ha guardado en el archivo: '{nombre_archivo}'")
            print("\nResumen de los últimos 5 registros guardados:")
        except ImportError:
            print(f"\nError: Para guardar el archivo en formato Excel (.xlsx), necesitas la librería 'openpyxl'.")
            print(f"Por favor, instálala ejecutando: pip install openpyxl")
            print(f"O instala todas las dependencias con: pip install -r requirements.txt")
            print(f"\nSi prefieres, puedes modificar el script para guardar en formato CSV, que no requiere 'openpyxl'.")
            # Opcionalmente, podríamos guardar en CSV como fallback aquí.
            # nombre_archivo_csv = nombre_archivo.replace('.xlsx', '.csv')
            # historial_cierre.to_csv(nombre_archivo_csv)
            # print(f"\nComo alternativa, se ha guardado en formato CSV: '{nombre_archivo_csv}'")
            return # Salir si no se pudo guardar en Excel y no hay fallback implementado activamente

        print(historial_cierre.tail().to_string())

    # Este es el manejador general de excepciones para el bloque try principal
    except Exception as e:
        print(f"\nOcurrió un error al procesar la solicitud: {e}")
        print("Asegúrate de tener conexión a internet y las librerías necesarias instaladas.")

# --- Ejecución del Script ---
if __name__ == "__main__":
    obtener_y_guardar_historial_conversion()
