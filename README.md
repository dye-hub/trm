# Historial de Tasas de Cambio USD/EUR a COP

Este script de Python descarga el historial de las tasas de cambio de Dólar Estadounidense (USD) a Peso Colombiano (COP) y de Euro (EUR) a Peso Colombiano (COP) para un rango de fechas especificado por el usuario. Los datos obtenidos corresponden a los precios de cierre diarios y se guardan en un archivo de Excel.

## Funcionalidades

- Permite al usuario ingresar un rango de fechas (inicio y fin) para la descarga de datos.
- Valida el formato de las fechas ingresadas (dd/mm/yyyy).
- Intercambia automáticamente las fechas si la fecha de inicio es posterior a la fecha de fin.
- Descarga los datos históricos utilizando la librería `yfinance`.
- Extrae y formatea los precios de cierre para las conversiones USD/COP y EUR/COP.
- Guarda los datos en un archivo Excel (`.xlsx`) con un nombre descriptivo que incluye el rango de fechas consultado (ej: `historial_divisas_cop_YYYYMMDD_a_YYYYMMDD.xlsx`).
- Muestra un resumen de los últimos 5 registros guardados en la consola.
- Maneja errores comunes como problemas de conexión o falta de datos para el período especificado.

## Requisitos

- Python 3.x
- Las siguientes librerías de Python:
  - `yfinance`
  - `pandas`

Puedes instalar las dependencias necesarias ejecutando:
```bash
pip install -r requirements.txt
```

## Cómo ejecutar el script

1.  **Clona o descarga este repositorio.**
2.  **Instala las dependencias:**
    Abre una terminal o línea de comandos, navega hasta el directorio donde se encuentra el script y ejecuta:
    ```bash
    pip install -r requirements.txt
    ```
    (Asegúrate de crear el archivo `requirements.txt` primero si aún no existe).
3.  **Ejecuta el script:**
    En la misma terminal, ejecuta el script `main.py`:
    ```bash
    python main.py
    ```
4.  **Ingresa las fechas:**
    El script te solicitará que ingreses la fecha de inicio y la fecha de fin en formato `dd/mm/yyyy`.
    - Ejemplo de fecha de inicio: `01/01/2023`
    - Ejemplo de fecha de fin: `31/12/2023`

5.  **Revisa el archivo de salida:**
    Una vez que el script finalice, encontrarás un archivo Excel (ej: `historial_divisas_cop_20230101_a_20231231.xlsx`) en el mismo directorio, conteniendo el historial de las tasas de cambio.

## Ejemplo de Uso

Al ejecutar el script, verás algo como esto en tu consola:

```
Este script descarga datos de divisas para un rango de fechas específico.
Ingrese la fecha de inicio (dd/mm/yyyy): 01/11/2023
Ingrese la fecha de fin (dd/mm/yyyy): 05/11/2023

Iniciando la descarga del historial desde 01/11/2023 hasta 05/11/2023...

¡Éxito! El historial de datos se ha guardado en el archivo: 'historial_divisas_cop_20231101_a_20231105.xlsx'

Resumen de los últimos 5 registros guardados:
            Valor Cierre USD/COP  Valor Cierre EUR/COP
Date
2023-11-01             4043.00             4277.46
2023-11-02             4090.50             4341.18
2023-11-03             4010.49             4297.19
```
*(Los valores pueden variar dependiendo de los datos del mercado en el momento de la consulta).*

## Contribuciones

Las sugerencias y contribuciones son bienvenidas. Por favor, abre un *issue* o envía un *pull request*.
