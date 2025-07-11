# -*- coding: utf-8 -*-
"""
Este script descarga el historial de tasas de cambio de USD a COP y de EUR a COP
para un rango de fechas especificado por el usuario y guarda los valores de
cierre en un archivo de Excel, usando una interfaz gráfica profesional con Tkinter y ttkthemes.
"""

import yfinance as yf
import pandas as pd
from datetime import datetime, date, timedelta
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading # Para ejecutar la descarga en un hilo separado y no bloquear la GUI

try:
    from ttkthemes import ThemedTk
    THEMED_TK_AVAILABLE = True
except ImportError:
    THEMED_TK_AVAILABLE = False
    # Este print es más para depuración, en un .exe no se vería si es --windowed
    # print("ttkthemes no está instalado. Se usará Tkinter estándar.")
    # print("Puedes instalarlo con: pip install ttkthemes")

# --- Lógica de negocio (modificada para GUI y barra de progreso) ---

def obtener_y_guardar_historial_conversion_gui(fecha_inicio_str, fecha_fin_str, status_label, progress_bar, boton_descargar, root_window):
    """
    Función principal para obtener el historial de tasas de cambio
    y guardarlo en un archivo de Excel, adaptada para la GUI mejorada.
    """
    if boton_descargar:
        boton_descargar.config(state=tk.DISABLED)
    
    progress_bar.start(10) # Iniciar animación de progreso indeterminado
    status_label.config(text="Procesando fechas...")
    root_window.update_idletasks() # Forzar actualización de la GUI

    try:
        fecha_inicio = datetime.strptime(fecha_inicio_str, '%d/%m/%Y').date()
        fecha_fin = datetime.strptime(fecha_fin_str, '%d/%m/%Y').date()

        if fecha_inicio > fecha_fin:
            messagebox.showwarning("Advertencia", "La fecha de inicio es posterior a la fecha de fin. Se han intercambiado.", parent=root_window)
            fecha_inicio, fecha_fin = fecha_fin, fecha_inicio
            # Actualizar campos de entrada en la GUI (requiere pasar los widgets de entrada o sus StringVars)
            # Por simplicidad, no se implementa la actualización de campos aquí, pero sería una mejora.

        status_label.config(text=f"Descargando datos para: {fecha_inicio.strftime('%d/%m/%Y')} - {fecha_fin.strftime('%d/%m/%Y')}...")
        root_window.update_idletasks()

        tickers = ["USDCOP=X", "EURCOP=X"]
        # yfinance descarga hasta el día anterior a 'end'. Para incluir 'fecha_fin', sumamos un día.
        fecha_fin_yf = fecha_fin + timedelta(days=1)
        data = yf.download(tickers, start=fecha_inicio, end=fecha_fin_yf, progress=False)

        status_label.config(text="Procesando datos descargados...")
        root_window.update_idletasks()

        if data.empty:
            messagebox.showerror("Error de Descarga", "No se pudieron obtener datos para el período especificado.\nEl mercado podría haber estado cerrado o hay un problema de conexión.", parent=root_window)
            status_label.config(text="Error en la descarga.")
            progress_bar.stop()
            if boton_descargar:
                boton_descargar.config(state=tk.NORMAL)
            return

        historial_cierre = data['Close'].copy()
        # Reindexar para asegurar que todas las fechas dentro del rango original estén presentes si es necesario,
        # aunque yfinance debería devolver solo los días con datos.
        # Considerar el caso de días festivos donde una moneda tiene datos y la otra no.
        # dropna() eliminará filas si CUALQUIER ticker no tiene dato ese día.
        historial_cierre.dropna(inplace=True)

        if historial_cierre.empty:
            messagebox.showerror("Sin Datos", "No hay datos de cierre disponibles para el período después de limpiar NaNs.", parent=root_window)
            status_label.config(text="Sin datos válidos.")
            progress_bar.stop()
            if boton_descargar:
                boton_descargar.config(state=tk.NORMAL)
            return
        
        historial_cierre.rename(columns={'USDCOP=X': 'Valor Cierre USD/COP', 'EURCOP=X': 'Valor Cierre EUR/COP'}, inplace=True)

        # Asegurarse que las columnas existen antes de redondear (si un ticker falla completamente)
        if 'Valor Cierre USD/COP' in historial_cierre.columns:
            historial_cierre.loc[:, 'Valor Cierre USD/COP'] = historial_cierre['Valor Cierre USD/COP'].round(2)
        if 'Valor Cierre EUR/COP' in historial_cierre.columns:
            historial_cierre.loc[:, 'Valor Cierre EUR/COP'] = historial_cierre['Valor Cierre EUR/COP'].round(2)

        status_label.config(text="Preparando para guardar archivo...")
        root_window.update_idletasks()

        nombre_sugerido = f"historial_divisas_cop_{fecha_inicio.strftime('%Y%m%d')}_a_{fecha_fin.strftime('%Y%m%d')}.xlsx"
        nombre_archivo = filedialog.asksaveasfilename(
            parent=root_window,
            defaultextension=".xlsx",
            initialfile=nombre_sugerido,
            title="Guardar archivo Excel",
            filetypes=[("Archivos Excel", "*.xlsx"), ("Todos los archivos", "*.*")]
        )

        if not nombre_archivo:
            status_label.config(text="Guardado cancelado.")
            progress_bar.stop()
            if boton_descargar:
                boton_descargar.config(state=tk.NORMAL)
            return

        try:
            status_label.config(text=f"Guardando archivo: {nombre_archivo.split('/')[-1]}...")
            root_window.update_idletasks()
            historial_cierre.to_excel(nombre_archivo, sheet_name='HistorialTasasDeCambio')
            progress_bar.stop()
            messagebox.showinfo("Éxito", f"¡Éxito! El historial de datos se ha guardado en:\n{nombre_archivo}", parent=root_window)
            status_label.config(text=f"¡Éxito! Guardado en {nombre_archivo.split('/')[-1]}")
        except ImportError:
            progress_bar.stop()
            messagebox.showerror("Error de Dependencia", "Para guardar en Excel, necesitas 'openpyxl'.\n\nInstálala con: pip install openpyxl", parent=root_window)
            status_label.config(text="Error: Falta 'openpyxl'.")
        except Exception as e_save:
            progress_bar.stop()
            messagebox.showerror("Error al Guardar", f"Ocurrió un error al guardar el archivo:\n{e_save}", parent=root_window)
            status_label.config(text="Error al guardar.")

    except ValueError:
        progress_bar.stop()
        messagebox.showerror("Error de Formato de Fecha", "Formato de fecha incorrecto. Usa dd/mm/yyyy.", parent=root_window)
        status_label.config(text="Error en formato de fecha.")
    except Exception as e_general:
        progress_bar.stop()
        messagebox.showerror("Error General", f"Ocurrió un error inesperado:\n{e_general}", parent=root_window)
        status_label.config(text="Error general.")
    finally:
        progress_bar.stop() # Asegurarse que la barra se detenga
        if boton_descargar:
            boton_descargar.config(state=tk.NORMAL)

# --- Creación de la GUI ---
def crear_gui():
    if THEMED_TK_AVAILABLE:
        root = ThemedTk(theme="arc") # Ejemplo de tema, otros: "plastik", "clearlooks", "radiance", "equilux", "itft1"
    else:
        root = tk.Tk()
        # Si ttkthemes no está, al menos intentamos poner un estilo ttk a los widgets
        style = ttk.Style()
        # style.theme_use('clam') # 'clam', 'alt', 'default', 'classic'

    root.title("Descargador Profesional de Tasas de Cambio")
    root.geometry("550x300") # Ajustar tamaño para más espacio
    root.resizable(False, False) # Evitar redimensionamiento para mantener el diseño

    # Estilo para los widgets (opcional, pero ayuda a la apariencia)
    style = ttk.Style(root)
    style.configure("TLabel", padding=5, font=('Helvetica', 10))
    style.configure("TButton", padding=8, font=('Helvetica', 10, 'bold'))
    style.configure("TEntry", padding=5, font=('Helvetica', 10))
    style.configure("Status.TLabel", padding=5, font=('Helvetica', 9), relief=tk.SUNKEN) # Estilo para el status_label
    style.configure("Error.TLabel", foreground="red") # No usado directamente, pero como ejemplo

    # Frame principal con padding
    main_frame = ttk.Frame(root, padding="15 15 15 15")
    main_frame.pack(expand=True, fill=tk.BOTH)

    # Sección de Fechas
    dates_frame = ttk.LabelFrame(main_frame, text="Selección de Fechas", padding="10")
    dates_frame.pack(fill=tk.X, pady=(0,10))

    ttk.Label(dates_frame, text="Fecha de Inicio (dd/mm/yyyy):").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
    entry_inicio_str = tk.StringVar(value=(date.today() - timedelta(days=7)).strftime('%d/%m/%Y')) # Por defecto una semana atrás
    entry_inicio = ttk.Entry(dates_frame, width=15, textvariable=entry_inicio_str, font=('Helvetica', 10))
    entry_inicio.grid(row=0, column=1, sticky=tk.EW, padx=5, pady=5)

    ttk.Label(dates_frame, text="Fecha de Fin (dd/mm/yyyy):").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
    entry_fin_str = tk.StringVar(value=date.today().strftime('%d/%m/%Y'))
    entry_fin = ttk.Entry(dates_frame, width=15, textvariable=entry_fin_str, font=('Helvetica', 10))
    entry_fin.grid(row=1, column=1, sticky=tk.EW, padx=5, pady=5)

    dates_frame.columnconfigure(1, weight=1) # Para que los Entry se expandan

    # Botón de descarga
    boton_descargar = ttk.Button(main_frame, text="Descargar y Guardar Historial")

    # Barra de progreso
    progress_bar = ttk.Progressbar(main_frame, mode='indeterminate', length=100)
    progress_bar.pack(fill=tk.X, pady=(5, 10))

    # Etiqueta de estado
    status_label = ttk.Label(main_frame, text="Listo para iniciar.", style="Status.TLabel")
    status_label.pack(fill=tk.X, side=tk.BOTTOM, pady=(5,0))

    # Acción del botón (usa threading)
    def on_descargar_click_thread():
        # Validar formato de fecha antes de iniciar el hilo para errores rápidos
        try:
            datetime.strptime(entry_inicio.get(), '%d/%m/%Y')
            datetime.strptime(entry_fin.get(), '%d/%m/%Y')
        except ValueError:
            messagebox.showerror("Error de Formato", "Formato de fecha incorrecto. Usa dd/mm/yyyy.", parent=root)
            status_label.config(text="Error en formato de fecha.")
            return

        thread = threading.Thread(target=obtener_y_guardar_historial_conversion_gui,
                                  args=(entry_inicio.get(), entry_fin.get(), status_label, progress_bar, boton_descargar, root))
        thread.daemon = True
        thread.start()

    boton_descargar.config(command=on_descargar_click_thread)
    boton_descargar.pack(fill=tk.X, ipady=5, pady=(0,5)) # ipady para altura interna del botón

    root.mainloop()

# --- Ejecución del Script ---
if __name__ == "__main__":
    crear_gui()
