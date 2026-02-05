"""
AUTOMATIZACIÓN COMPLETA - CONTROL DE PAGOS - VERSIÓN 2.0
"""

import pandas as pd
import win32com.client
import pythoncom
from openpyxl.styles import Font, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
from pathlib import Path
from datetime import datetime, timedelta
import locale
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry
import configparser
import sys
import time
import traceback

# Configuración de español
try:
    locale.setlocale(locale.LC_TIME, 'Spanish_Spain.1252')  # Windows
except locale.Error:
    try:
        locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')  # Linux
    except locale.Error:
        pass

class ConfiguradorRutas:
    """Manejador de configuración de rutas"""
    
    def __init__(self):
        self.config_file = Path("config_pagos.ini")
        self.config = configparser.ConfigParser()
        
    def cargar_o_crear_config(self):
        """Carga configuración existente o crea una nueva"""
        if self.config_file.exists():
            self.config.read(self.config_file, encoding='utf-8')
            return True
        else:
            return self.crear_configuracion_inicial()
    
    def crear_configuracion_inicial(self):
        """Crea configuración inicial solicitando rutas al usuario"""
        root = tk.Tk()
        root.withdraw()
        
        messagebox.showinfo(
            "Primera Configuración",
            "Por favor, seleccione las rutas necesarias para el programa."
        )
        
        # Solicitar archivo origen
        messagebox.showinfo("Paso 1", "Seleccione el archivo CONTROL DE PAGOS.xlsm")
        archivo_origen = filedialog.askopenfilename(
            title="Seleccionar CONTROL DE PAGOS.xlsm",
            filetypes=[("Excel Macro", "*.xlsm"), ("Todos", "*.*")]
        )
        
        if not archivo_origen:
            messagebox.showerror("Error", "Debe seleccionar el archivo origen.")
            return False
        
        # Solicitar carpeta de proyecciones
        messagebox.showinfo("Paso 2", "Seleccione la carpeta donde se guardarán las PROYECCIONES")
        carpeta_proyecciones = filedialog.askdirectory(
            title="Seleccionar carpeta de PROYECCIONES"
        )
        
        if not carpeta_proyecciones:
            messagebox.showerror("Error", "Debe seleccionar la carpeta de proyecciones.")
            return False
        
        # Solicitar archivo final
        messagebox.showinfo("Paso 3", "Seleccione el archivo CONTROL PAGOS.xlsx (archivo final)")
        archivo_final = filedialog.askopenfilename(
            title="Seleccionar CONTROL PAGOS.xlsx",
            filetypes=[("Excel", "*.xlsx"), ("Todos", "*.*")],
            initialdir=carpeta_proyecciones
        )
        
        if not archivo_final:
            messagebox.showerror("Error", "Debe seleccionar el archivo final.")
            return False
        
        # Guardar configuración
        self.config['RUTAS'] = {
            'archivo_origen': archivo_origen,
            'carpeta_proyecciones': carpeta_proyecciones,
            'archivo_final': archivo_final
        }
        
        with open(self.config_file, 'w', encoding='utf-8') as f:
            self.config.write(f)
        
        messagebox.showinfo("Configuración Guardada", 
                          f"La configuración se ha guardado en:\n{self.config_file.absolute()}")
        
        root.destroy()
        return True
    
    def obtener_rutas(self):
        """Obtiene las rutas configuradas"""
        return {
            'origen': Path(self.config['RUTAS']['archivo_origen']),
            'proyecciones': Path(self.config['RUTAS']['carpeta_proyecciones']),
            'final': Path(self.config['RUTAS']['archivo_final'])
        }

class InterfazModerna:
    """
    Interfaz gráfica moderna para seleccionar fecha y tipo de proceso
    """
    def __init__(self):
        self.fecha_seleccionada = None
        self.ejecutar_proceso = False
        self.opcion_proceso = None  # 1: Solo proyección, 2: Solo anexar, 3: Ambos
        
        # Colores del tema (mismo que V1)
        self.COLOR_PRIMARIO = "#2C3E50"
        self.COLOR_SECUNDARIO = "#3498DB"
        self.COLOR_ACENTO = "#27AE60"
        self.COLOR_FONDO = "#ECF0F1"
        self.COLOR_TEXTO = "#2C3E50"
        self.COLOR_ERROR = "#E74C3C"
        self.COLOR_NARANJA = "#E67E22"
        
    def crear_ventana(self):
        """Crea la ventana de interfaz moderna"""
        self.root = tk.Tk()
        self.root.title("Control de Pagos GCO - V2")
        self.root.geometry("700x750")
        self.root.resizable(False, False)
        self.root.configure(bg=self.COLOR_FONDO)
        
        # Centrar ventana
        self.centrar_ventana()
        
        # Configurar estilo
        self.configurar_estilos()
        
        # Frame principal
        main_frame = tk.Frame(self.root, bg=self.COLOR_FONDO)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=0, pady=0)
        
        # Header
        self.crear_header(main_frame)
        
        # Contenido principal
        self.crear_contenido(main_frame)
        
        # Footer con botones (mismo diseño que V1)
        self.crear_footer(main_frame)
        
        self.root.mainloop()
    
    def centrar_ventana(self):
        """Centra la ventana en la pantalla"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def configurar_estilos(self):
        """Configura los estilos personalizados"""
        style = ttk.Style()
        style.theme_use('clam')
        
        style.configure(
            "Modern.TLabelframe",
            background=self.COLOR_FONDO,
            bordercolor=self.COLOR_SECUNDARIO,
            borderwidth=2
        )
        style.configure(
            "Modern.TLabelframe.Label",
            background=self.COLOR_FONDO,
            foreground=self.COLOR_PRIMARIO,
            font=("Segoe UI", 11, "bold")
        )
    
    def crear_header(self, parent):
        """Crea el header con título"""
        header_frame = tk.Frame(parent, bg=self.COLOR_PRIMARIO, height=140)
        header_frame.pack(fill=tk.X, pady=0)
        header_frame.pack_propagate(False)
        
        content = tk.Frame(header_frame, bg=self.COLOR_PRIMARIO)
        content.place(relx=0.5, rely=0.5, anchor="center")
        
        icon_label = tk.Label(
            content,
            text="📊",
            font=("Segoe UI", 30),
            bg=self.COLOR_PRIMARIO,
            fg="white"
        )
        icon_label.pack(side=tk.LEFT, padx=(0, 15))
        
        text_frame = tk.Frame(content, bg=self.COLOR_PRIMARIO)
        text_frame.pack(side=tk.LEFT)
        
        titulo = tk.Label(
            text_frame,
            text="Control de Pagos",
            font=("Segoe UI", 20, "bold"),
            bg=self.COLOR_PRIMARIO,
            fg="white"
        )
        titulo.pack(anchor="w")
        
        subtitulo = tk.Label(
            text_frame,
            text="Sistema de Gestión de Importaciones - V2",
            font=("Segoe UI", 11),
            bg=self.COLOR_PRIMARIO,
            fg="#BDC3C7"
        )
        subtitulo.pack(anchor="w")
    
    def crear_contenido(self, parent):
        """Crea el contenido principal"""
        content_frame = tk.Frame(parent, bg=self.COLOR_FONDO)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=30)
        
        # Tarjeta principal
        card_frame = tk.Frame(
            content_frame,
            bg="white",
            relief=tk.FLAT,
            borderwidth=0
        )
        card_frame.pack(fill=tk.BOTH, expand=True)
        
        # Agregar sombra simulada
        self.agregar_sombra(card_frame)
        
        # Padding interno
        inner_frame = tk.Frame(card_frame, bg="white")
        inner_frame.pack(fill=tk.BOTH, expand=True, padx=25, pady=25)
        
        # SECCIÓN 1: Selección de Proceso
        self.crear_seccion_proceso(inner_frame)
        
        # Separador
        ttk.Separator(inner_frame, orient='horizontal').pack(fill=tk.X, pady=20)
        
        # SECCIÓN 2: Selección de fecha
        self.crear_seccion_fecha(inner_frame)
    
    def agregar_sombra(self, widget):
        """Agrega efecto de sombra al widget"""
        widget.configure(
            highlightbackground="#BDC3C7",
            highlightcolor="#BDC3C7",
            highlightthickness=1
        )
    
    def crear_seccion_proceso(self, parent):
        """Crea la sección de selección de proceso"""
        # Título de sección
        titulo_frame = tk.Frame(parent, bg="white")
        titulo_frame.pack(fill=tk.X, pady=(0, 15))
        
        icono = tk.Label(
            titulo_frame,
            text="⚙️",
            font=("Segoe UI", 16),
            bg="white"
        )
        icono.pack(side=tk.LEFT, padx=(0, 10))
        
        titulo = tk.Label(
            titulo_frame,
            text="Seleccione el Proceso a Ejecutar",
            font=("Segoe UI", 13, "bold"),
            bg="white",
            fg=self.COLOR_PRIMARIO
        )
        titulo.pack(side=tk.LEFT)
        
        # Variable para opciones
        self.var_opcion = tk.IntVar(value=3)
        
        # Frame para opciones
        opciones_frame = tk.Frame(parent, bg="white")
        opciones_frame.pack(fill=tk.X, pady=10)
        
        # Opción 1: Solo Proyección
        opcion1_frame = self.crear_opcion_radio(
            opciones_frame,
            texto="📋 Crear Solo Proyección",
            descripcion="Genera únicamente el archivo de proyección semanal",
            valor=1,
            variable=self.var_opcion,
            color_acento=self.COLOR_SECUNDARIO
        )
        opcion1_frame.pack(fill=tk.X, pady=5)
        
        # Opción 2: Solo Anexar
        opcion2_frame = self.crear_opcion_radio(
            opciones_frame,
            texto="➕ Anexar a Archivo Final",
            descripcion="Agrega registros de una proyección existente al archivo final",
            valor=2,
            variable=self.var_opcion,
            color_acento=self.COLOR_NARANJA
        )
        opcion2_frame.pack(fill=tk.X, pady=5)
        
        # Opción 3: Proceso Completo
        opcion3_frame = self.crear_opcion_radio(
            opciones_frame,
            texto="🔄 Proceso Completo",
            descripcion="Crea proyección y anexa automáticamente al archivo final",
            valor=3,
            variable=self.var_opcion,
            color_acento=self.COLOR_ACENTO
        )
        opcion3_frame.pack(fill=tk.X, pady=5)
    
    def crear_opcion_radio(self, parent, texto, descripcion, valor, variable, color_acento):
        """Crea una opción de radio button personalizada"""
        frame_contenedor = tk.Frame(parent, bg="white")
        
        # Frame interno con borde
        frame_opcion = tk.Frame(
            frame_contenedor,
            bg="#F8F9FA",
            relief=tk.SOLID,
            borderwidth=1,
            highlightbackground="#E0E0E0",
            highlightthickness=1
        )
        frame_opcion.pack(fill=tk.X, padx=5, pady=2)
        
        # Radio button
        radio = tk.Radiobutton(
            frame_opcion,
            variable=variable,
            value=valor,
            bg="#F8F9FA",
            activebackground="#F8F9FA",
            selectcolor=color_acento,
            cursor="hand2"
        )
        radio.pack(side=tk.LEFT, padx=10, pady=10)
        
        # Textos
        text_frame = tk.Frame(frame_opcion, bg="#F8F9FA")
        text_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, pady=10)
        
        label_texto = tk.Label(
            text_frame,
            text=texto,
            font=("Segoe UI", 11, "bold"),
            bg="#F8F9FA",
            fg=self.COLOR_PRIMARIO,
            cursor="hand2"
        )
        label_texto.pack(anchor="w")
        
        label_desc = tk.Label(
            text_frame,
            text=descripcion,
            font=("Segoe UI", 9),
            bg="#F8F9FA",
            fg="#7F8C8D",
            cursor="hand2"
        )
        label_desc.pack(anchor="w")
        
        # Hacer clickeable todo el frame
        def seleccionar(e=None):
            variable.set(valor)
        
        for widget in [frame_opcion, label_texto, label_desc]:
            widget.bind("<Button-1>", seleccionar)
        
        return frame_contenedor
    
    def crear_seccion_fecha(self, parent):
        """Crea la sección de selección de fecha"""
        # Título de sección
        titulo_frame = tk.Frame(parent, bg="white")
        titulo_frame.pack(fill=tk.X, pady=(0, 15))
        
        icono = tk.Label(
            titulo_frame,
            text="📅",
            font=("Segoe UI", 16),
            bg="white"
        )
        icono.pack(side=tk.LEFT, padx=(0, 10))
        
        titulo = tk.Label(
            titulo_frame,
            text="Seleccione la Fecha de Proyección",
            font=("Segoe UI", 13, "bold"),
            bg="white",
            fg=self.COLOR_PRIMARIO
        )
        titulo.pack(side=tk.LEFT)
        
        # Frame para el calendario
        cal_frame = tk.Frame(parent, bg="white")
        cal_frame.pack(pady=10)
        
        # Calendario
        self.calendario = DateEntry(
            cal_frame,
            width=25,
            background=self.COLOR_SECUNDARIO,
            foreground='white',
            borderwidth=2,
            font=("Segoe UI", 11),
            date_pattern='dd/mm/yyyy',
            locale='es_ES'
        )
        self.calendario.pack(pady=5)
        
        # Texto de ayuda
        ayuda = tk.Label(
            parent,
            text="💡 Seleccione la fecha hasta la cual desea proyectar los pagos",
            font=("Segoe UI", 9, "italic"),
            bg="white",
            fg="#7F8C8D"
        )
        ayuda.pack(pady=(10, 0))
    
    def crear_footer(self, parent):
        """Crea el footer con botones (mismo diseño que V1)"""
        footer_frame = tk.Frame(parent, bg=self.COLOR_FONDO, height=80)
        footer_frame.pack(fill=tk.X, side=tk.BOTTOM)
        footer_frame.pack_propagate(False)
        
        # Contenedor de botones
        buttons_frame = tk.Frame(footer_frame, bg=self.COLOR_FONDO)
        buttons_frame.place(relx=0.5, rely=0.5, anchor="center")
        
        # Botón Cancelar (mismo estilo que V1)
        btn_cancelar = tk.Button(
            buttons_frame,
            text="✖  Cancelar",
            command=self.cancelar,
            font=("Segoe UI", 11, "bold"),
            bg="white",
            fg=self.COLOR_ERROR,
            activebackground="#fadbd8",
            activeforeground=self.COLOR_ERROR,
            relief=tk.FLAT,
            borderwidth=0,
            padx=30,
            pady=12,
            cursor="hand2",
            width=15
        )
        btn_cancelar.pack(side=tk.LEFT, padx=10)
        
        # Efectos hover para cancelar
        def on_enter_cancelar(e):
            btn_cancelar.configure(bg="#fadbd8")
        
        def on_leave_cancelar(e):
            btn_cancelar.configure(bg="white")
        
        btn_cancelar.bind("<Enter>", on_enter_cancelar)
        btn_cancelar.bind("<Leave>", on_leave_cancelar)
        
        # Botón Continuar (mismo estilo que V1)
        btn_continuar = tk.Button(
            buttons_frame,
            text="✓  Continuar",
            command=self.continuar,
            font=("Segoe UI", 11, "bold"),
            bg=self.COLOR_ACENTO,
            fg="white",
            activebackground="#229954",
            activeforeground="white",
            relief=tk.FLAT,
            borderwidth=0,
            padx=30,
            pady=12,
            cursor="hand2",
            width=15
        )
        btn_continuar.pack(side=tk.LEFT, padx=10)
        
        # Efectos hover para continuar
        def on_enter_continuar(e):
            btn_continuar.configure(bg="#229954")
        
        def on_leave_continuar(e):
            btn_continuar.configure(bg=self.COLOR_ACENTO)
        
        btn_continuar.bind("<Enter>", on_enter_continuar)
        btn_continuar.bind("<Leave>", on_leave_continuar)
    
    def cancelar(self):
        """Cancela la operación"""
        self.ejecutar_proceso = False
        self.root.destroy()
    
    def continuar(self):
        """Confirma y continúa con el proceso"""
        self.fecha_seleccionada = self.calendario.get_date()
        self.opcion_proceso = self.var_opcion.get()
        self.ejecutar_proceso = True
        self.root.destroy()

class VentanaProgreso:
    """Ventana de progreso"""
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Procesando...")
        self.root.geometry("500x200")
        self.root.resizable(False, False)
        
        # Centrar
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - 250
        y = (self.root.winfo_screenheight() // 2) - 100
        self.root.geometry(f'500x200+{x}+{y}')
        
        # Contenido
        self.label = tk.Label(
            self.root,
            text="Procesando...",
            font=("Segoe UI", 12),
            pady=20
        )
        self.label.pack()
        
        self.progress = ttk.Progressbar(
            self.root,
            mode='indeterminate',
            length=400
        )
        self.progress.pack(pady=20)
        self.progress.start(10)
        
        self.detalle = tk.Label(
            self.root,
            text="",
            font=("Segoe UI", 9),
            fg="#7F8C8D"
        )
        self.detalle.pack()
        
        self.root.update()
    
    def actualizar(self, mensaje, detalle=""):
        """Actualiza el mensaje de progreso"""
        self.label.config(text=mensaje)
        self.detalle.config(text=detalle)
        self.root.update()
    
    def cerrar(self):
        """Cierra la ventana"""
        try:
            self.root.destroy()
        except:
            pass

class CopiarArchivo:
    """Clase principal para copiar y procesar archivo"""
    
    def __init__(self, fecha_filtrado, ventana_progreso, opcion_proceso, rutas_config):
        self.fecha_filtrado = fecha_filtrado
        self.ventana_progreso = ventana_progreso
        self.opcion_proceso = opcion_proceso
        
        # Configurar rutas desde config
        self.ruta_origen = rutas_config['origen']
        self.carpeta_proyecciones = rutas_config['proyecciones']
        self.ruta_destino_final = rutas_config['final']
        
        # Archivo de proyección seleccionado (para opción 2)
        self.archivo_proyeccion_seleccionado = None
    
    def log(self, mensaje, tipo="INFO"):
        """Sistema de logging mejorado"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        prefijos = {
            "INFO": "ℹ️",
            "OK": "✅",
            "WARN": "⚠️",
            "ERROR": "❌"
        }
        prefijo = prefijos.get(tipo, "•")
        print(f"[{timestamp}] {prefijo} {mensaje}")
        
        if self.ventana_progreso:
            self.ventana_progreso.actualizar(mensaje, f"Tipo: {tipo}")
    
    def crear_estructura_carpetas(self, fecha):
        """Crea estructura de carpetas por año/mes"""
        año = fecha.strftime("%Y")
        mes = fecha.strftime("%B").capitalize()
        
        carpeta_año = self.carpeta_proyecciones / año
        carpeta_mes = carpeta_año / mes
        
        carpeta_mes.mkdir(parents=True, exist_ok=True)
        self.log(f"Carpeta creada: {carpeta_mes}", "OK")
        
        return carpeta_mes
    
    def crear_nombre_archivo(self, fecha):
        """Crea el nombre del archivo de proyección"""
        fecha_str = fecha.strftime("%d.%m.%Y")
        return f"PROYECCIÓN {fecha_str}.xlsm"
    
    def copiar_archivo_base(self, destino):
        """Copia el archivo base"""
        self.log("Copiando archivo base...", "INFO")
        
        import shutil
        shutil.copy2(self.ruta_origen, destino)
        
        self.log(f"Archivo copiado: {destino.name}", "OK")
        time.sleep(1)
    
    def leer_datos_control_pagos(self, ruta_archivo):
        """Lee datos del archivo de control de pagos"""
        self.log("Leyendo datos del archivo...", "INFO")
        
        pythoncom.CoInitialize()
        excel = None
        wb = None
        
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            
            wb = excel.Workbooks.Open(str(ruta_archivo.absolute()))
            ws = wb.Sheets("PAGOS IMPORTACIÓN")
            
            ultima_fila = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
            ultima_columna = ws.Cells(1, ws.Columns.Count).End(-4159).Column
            
            datos = ws.Range(ws.Cells(1, 1), ws.Cells(ultima_fila, ultima_columna)).Value
            
            wb.Close(SaveChanges=False)
            excel.Quit()
            
            df = pd.DataFrame(datos[1:], columns=datos[0])
            
            self.log(f"Datos leídos: {len(df)} registros", "OK")
            return df
            
        except Exception as e:
            self.log(f"Error al leer datos: {str(e)}", "ERROR")
            return None
            
        finally:
            if wb:
                wb.Close(SaveChanges=False)
            if excel:
                excel.Quit()
            pythoncom.CoUninitialize()
    
    def filtrar_por_fecha(self, df, fecha_limite):
        """Filtra registros por fecha"""
        self.log(f"Filtrando datos hasta {fecha_limite.strftime('%d/%m/%Y')}...", "INFO")
        
        df['FECHA PAGO'] = pd.to_datetime(df['FECHA PAGO'], errors='coerce')
        df_filtrado = df[df['FECHA PAGO'] <= pd.Timestamp(fecha_limite)].copy()
        
        self.log(f"Registros filtrados: {len(df_filtrado)}", "OK")
        return df_filtrado
    
    def preparar_datos_segunda_hoja(self, df):
        """Prepara datos para segunda hoja"""
        self.log("Preparando datos para segunda hoja...", "INFO")
        
        columnas_necesarias = [
            'FACTURA COMERCIAL', 'PROVEEDOR', 'BANCO', 'BENEFICIARIO',
            'FECHA PAGO', 'MONTO DIVISA', 'DIVISA', 'TASA CAMBIO',
            'MONTO BOLIVIANOS', 'REFERENCIA', 'N° SOLICITUD', 'OBSERVACIONES'
        ]
        
        df_segunda = df[columnas_necesarias].copy()
        
        self.log(f"Datos preparados: {len(df_segunda)} registros", "OK")
        return df_segunda
    
    def agrupar_y_calcular(self, df):
        """Agrupa y calcula totales"""
        self.log("Agrupando y calculando totales...", "INFO")
        
        df_agrupado = df.groupby('BANCO').agg({
            'MONTO DIVISA': 'sum',
            'MONTO BOLIVIANOS': 'sum'
        }).reset_index()
        
        df_agrupado.columns = ['BANCO', 'MONTO DIVISA', 'MONTO BOLIVIANOS']
        
        # Agregar fila de totales
        total_divisa = df_agrupado['MONTO DIVISA'].sum()
        total_bs = df_agrupado['MONTO BOLIVIANOS'].sum()
        
        fila_total = pd.DataFrame({
            'BANCO': ['TOTAL'],
            'MONTO DIVISA': [total_divisa],
            'MONTO BOLIVIANOS': [total_bs]
        })
        
        df_agrupado = pd.concat([df_agrupado, fila_total], ignore_index=True)
        
        self.log(f"Totales calculados: {len(df_agrupado)} bancos", "OK")
        return df_agrupado
    
    def crear_nombre_segunda_hoja(self, fecha):
        """Crea nombre para segunda hoja"""
        dia = fecha.strftime("%d")
        mes = fecha.strftime("%b").upper()
        return f"{dia} {mes}"
    
    def guardar_proyeccion_com(self, ruta_archivo, df_agrupado, nombre_hoja):
        """Guarda proyección usando COM"""
        self.log("Guardando proyección en segunda hoja...", "INFO")
        
        pythoncom.CoInitialize()
        excel = None
        wb = None
        
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            
            wb = excel.Workbooks.Open(str(ruta_archivo.absolute()))
            
            # Verificar si existe la hoja
            hoja_existe = False
            for sheet in wb.Sheets:
                if sheet.Name == nombre_hoja:
                    ws = sheet
                    hoja_existe = True
                    break
            
            if not hoja_existe:
                ws = wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))
                ws.Name = nombre_hoja
            
            # Limpiar hoja
            ws.Cells.Clear()
            
            # Escribir encabezados
            encabezados = ['BANCO', 'MONTO DIVISA', 'MONTO BOLIVIANOS']
            for col_idx, encabezado in enumerate(encabezados, start=1):
                celda = ws.Cells(1, col_idx)
                celda.Value = encabezado
                celda.Font.Bold = True
                celda.Font.Size = 11
            
            # Escribir datos
            datos = df_agrupado.values.tolist()
            for fila_idx, fila in enumerate(datos, start=2):
                for col_idx, valor in enumerate(fila, start=1):
                    ws.Cells(fila_idx, col_idx).Value = valor
            
            # Formatear última fila (totales)
            ultima_fila = len(datos) + 1
            for col in range(1, 4):
                celda = ws.Cells(ultima_fila, col)
                celda.Font.Bold = True
            
            # Autoajustar columnas
            ws.Columns.AutoFit()
            
            wb.Save()
            self.log(f"Proyección guardada en hoja '{nombre_hoja}'", "OK")
            
        except Exception as e:
            self.log(f"Error al guardar proyección: {str(e)}", "ERROR")
            raise
            
        finally:
            if wb:
                wb.Close(SaveChanges=True)
            if excel:
                excel.Quit()
            pythoncom.CoUninitialize()
    
    def preparar_df_final(self, df):
        """Prepara dataframe para archivo final"""
        self.log("Preparando datos para archivo final...", "INFO")
        
        columnas_final = [
            'FACTURA COMERCIAL', 'PROVEEDOR', 'BANCO', 'BENEFICIARIO',
            'FECHA PAGO', 'MONTO DIVISA', 'DIVISA', 'TASA CAMBIO',
            'MONTO BOLIVIANOS', 'REFERENCIA', 'N° SOLICITUD', 'OBSERVACIONES'
        ]
        
        df_final = df[columnas_final].copy()
        
        # Formatear fecha
        df_final['FECHA PAGO'] = pd.to_datetime(df_final['FECHA PAGO']).dt.strftime('%d/%m/%Y')
        
        self.log(f"Datos preparados para archivo final: {len(df_final)} registros", "OK")
        return df_final
    
    def anexar_archivo_final_com(self, df_detalle):
        """Anexa datos al archivo final usando COM"""
        self.log("Anexando datos al archivo final", "INFO")
        
        pythoncom.CoInitialize()
        excel = None
        wb = None
        
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            
            archivo_path = str(self.ruta_destino_final.absolute())
            wb = excel.Workbooks.Open(archivo_path)
            
            # Búsqueda flexible de la hoja
            ws = None
            for sheet in wb.Sheets:
                nombre_u = sheet.Name.upper()
                if 'PAGOS' in nombre_u and 'IMPOR' in nombre_u:
                    ws = sheet
                    break
            
            if not ws:
                ws = wb.Sheets(1)
                self.log(f"Usando hoja: '{ws.Name}'", "WARN")
            else:
                self.log(f"Escribiendo en hoja: '{ws.Name}'", "OK")
            
            # Preparar datos
            datos = df_detalle.fillna("").values.tolist()
            if not datos:
                self.log("No hay datos para anexar", "WARN")
                return
            
            # Encontrar última fila
            last_row = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
            if last_row == 1 and ws.Cells(1, 1).Value is None:
                start_row = 1
            else:
                start_row = last_row + 1
            
            self.log(f"Insertando {len(datos)} registros en fila {start_row}", "INFO")
            
            # Escribir por rango
            num_filas = len(datos)
            num_cols = len(datos[0])
            
            destino = ws.Range(
                ws.Cells(start_row, 1),
                ws.Cells(start_row + num_filas - 1, num_cols)
            )
            destino.Value = datos
            
            wb.Save()
            self.log(f"¡ÉXITO! Datos guardados en {self.ruta_destino_final.name}", "OK")
            
        except Exception as e:
            self.log(f"Error crítico al anexar: {str(e)}", "ERROR")
            raise
            
        finally:
            if wb:
                wb.Close(SaveChanges=True)
            if excel:
                excel.Quit()
            pythoncom.CoUninitialize()
    
    def agregar_a_archivo_final(self, df_detalle):
        """Proceso completo para agregar a archivo final"""
        try:
            df_final = self.preparar_df_final(df_detalle)
            self.anexar_archivo_final_com(df_final)
        except Exception as e:
            self.log(f"Error en archivo final: {str(e)}", "ERROR")
    
    def verificar_archivo_proyeccion(self, fecha):
        """Verifica si existe archivo de proyección para la fecha"""
        carpeta_destino = self.crear_estructura_carpetas(fecha)
        nombre_archivo = self.crear_nombre_archivo(fecha)
        ruta_archivo = carpeta_destino / nombre_archivo
        
        return ruta_archivo.exists(), ruta_archivo
    
    def seleccionar_archivo_proyeccion(self):
        """Permite al usuario seleccionar un archivo de proyección"""
        self.log("Solicitando selección de archivo de proyección...", "INFO")
        
        archivo_seleccionado = filedialog.askopenfilename(
            title="Seleccione el archivo de PROYECCIÓN",
            filetypes=[
                ("Excel Macro", "*.xlsm"),
                ("Excel", "*.xlsx"),
                ("Todos", "*.*")
            ],
            initialdir=self.carpeta_proyecciones
        )
        
        if archivo_seleccionado:
            self.log(f"Archivo seleccionado: {Path(archivo_seleccionado).name}", "OK")
            return Path(archivo_seleccionado)
        else:
            self.log("No se seleccionó ningún archivo", "WARN")
            return None
    
    def leer_datos_proyeccion(self, ruta_proyeccion):
        """Lee datos del archivo de proyección"""
        self.log(f"Leyendo datos de proyección: {ruta_proyeccion.name}", "INFO")
        
        pythoncom.CoInitialize()
        excel = None
        wb = None
        
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            
            wb = excel.Workbooks.Open(str(ruta_proyeccion.absolute()))
            ws = wb.Sheets("PAGOS IMPORTACIÓN")
            
            ultima_fila = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
            ultima_columna = ws.Cells(1, ws.Columns.Count).End(-4159).Column
            
            datos = ws.Range(ws.Cells(1, 1), ws.Cells(ultima_fila, ultima_columna)).Value
            
            wb.Close(SaveChanges=False)
            excel.Quit()
            
            df = pd.DataFrame(datos[1:], columns=datos[0])
            
            self.log(f"Datos leídos de proyección: {len(df)} registros", "OK")
            return df
            
        except Exception as e:
            self.log(f"Error al leer proyección: {str(e)}", "ERROR")
            return None
            
        finally:
            if wb:
                wb.Close(SaveChanges=False)
            if excel:
                excel.Quit()
            pythoncom.CoUninitialize()
    
    def ejecutar_proceso_completo(self):
        """Ejecuta el proceso completo (Opción 3)"""
        try:
            if not self.ruta_origen.exists():
                self.log(f"No se encuentra el archivo original", "ERROR")
                messagebox.showerror("Error", f"No se encuentra el archivo:\n{self.ruta_origen}")
                return None
            
            fecha_proyeccion = self.fecha_filtrado
            self.log(f"Fecha de proyección: {fecha_proyeccion.strftime('%d/%m/%Y')}", "INFO")
            
            carpeta_destino = self.crear_estructura_carpetas(fecha_proyeccion)
            nombre_archivo = self.crear_nombre_archivo(fecha_proyeccion)
            ruta_archivo_nuevo = carpeta_destino / nombre_archivo
            
            self.copiar_archivo_base(ruta_archivo_nuevo)
            
            time.sleep(2)
            
            df_original = self.leer_datos_control_pagos(ruta_archivo_nuevo)
            if df_original is None:
                return None
            
            df_filtrado = self.filtrar_por_fecha(df_original, fecha_proyeccion)
            
            if len(df_filtrado) == 0:
                self.log("No se encontraron registros", "WARN")
                messagebox.showwarning("Sin registros", "No se encontraron registros para la fecha seleccionada.")
                return None
            
            df_segunda = self.preparar_datos_segunda_hoja(df_filtrado)
            df_agrupado = self.agrupar_y_calcular(df_segunda)
            
            nombre_segunda_hoja = self.crear_nombre_segunda_hoja(fecha_proyeccion)
            
            self.guardar_proyeccion_com(ruta_archivo_nuevo, df_agrupado, nombre_segunda_hoja)
            
            self.agregar_a_archivo_final(df_segunda)
            
            messagebox.showinfo(
                "¡Proceso Completado!",
                f"El proceso ha finalizado exitosamente.\n\n"
                f"📁 Proyección guardada en:\n{ruta_archivo_nuevo}\n\n"
                f"📁 Archivo final actualizado:\n{self.ruta_destino_final.name}"
            )
            return str(ruta_archivo_nuevo)
            
        except Exception as e:
            self.log(f"ERROR CRÍTICO: {str(e)}", "ERROR")
            traceback.print_exc()
            messagebox.showerror("Error", f"Ocurrió un error:\n\n{str(e)}")
            return None
    
    def ejecutar_solo_proyeccion(self):
        """Ejecuta solo la creación de proyección (Opción 1)"""
        try:
            if not self.ruta_origen.exists():
                self.log(f"No se encuentra el archivo original", "ERROR")
                messagebox.showerror("Error", f"No se encuentra el archivo:\n{self.ruta_origen}")
                return None
            
            fecha_proyeccion = self.fecha_filtrado
            self.log(f"Fecha de proyección: {fecha_proyeccion.strftime('%d/%m/%Y')}", "INFO")
            
            carpeta_destino = self.crear_estructura_carpetas(fecha_proyeccion)
            nombre_archivo = self.crear_nombre_archivo(fecha_proyeccion)
            ruta_archivo_nuevo = carpeta_destino / nombre_archivo
            
            self.copiar_archivo_base(ruta_archivo_nuevo)
            
            df_original = self.leer_datos_control_pagos(ruta_archivo_nuevo)
            if df_original is None:
                return None
            
            df_filtrado = self.filtrar_por_fecha(df_original, fecha_proyeccion)
            
            if len(df_filtrado) == 0:
                self.log("No se encontraron registros", "WARN")
                messagebox.showwarning("Sin registros", "No se encontraron registros para la fecha seleccionada.")
                return None
            
            df_segunda = self.preparar_datos_segunda_hoja(df_filtrado)
            df_agrupado = self.agrupar_y_calcular(df_segunda)
            
            nombre_segunda_hoja = self.crear_nombre_segunda_hoja(fecha_proyeccion)
            
            self.guardar_proyeccion_com(ruta_archivo_nuevo, df_agrupado, nombre_segunda_hoja)
            
            messagebox.showinfo(
                "¡Proyección Creada!",
                f"La proyección se ha creado exitosamente.\n\n"
                f"📁 Archivo guardado en:\n{ruta_archivo_nuevo}"
            )
            return str(ruta_archivo_nuevo)
            
        except Exception as e:
            self.log(f"ERROR CRÍTICO: {str(e)}", "ERROR")
            traceback.print_exc()
            messagebox.showerror("Error", f"Ocurrió un error:\n\n{str(e)}")
            return None
    
    def ejecutar_solo_anexar(self):
        """Ejecuta solo el anexado al archivo final (Opción 2)"""
        try:
            # MEJORA: Permitir seleccionar archivo de proyección
            self.log("Solicitando selección de archivo de proyección...", "INFO")
            
            messagebox.showinfo(
                "Seleccionar Archivo",
                "Por favor, seleccione el archivo de PROYECCIÓN\n"
                "que desea anexar al archivo final."
            )
            
            ruta_proyeccion = self.seleccionar_archivo_proyeccion()
            
            if not ruta_proyeccion:
                self.log("Usuario canceló la selección de archivo", "WARN")
                messagebox.showwarning(
                    "Operación Cancelada",
                    "No se seleccionó ningún archivo de proyección."
                )
                return None
            
            if not ruta_proyeccion.exists():
                self.log(f"Archivo no encontrado: {ruta_proyeccion}", "ERROR")
                messagebox.showerror(
                    "Archivo No Encontrado",
                    f"El archivo seleccionado no existe:\n{ruta_proyeccion}"
                )
                return None
            
            self.log(f"Archivo de proyección seleccionado: {ruta_proyeccion.name}", "OK")
            
            # Leer datos del archivo de proyección
            df_proyeccion = self.leer_datos_proyeccion(ruta_proyeccion)
            
            if df_proyeccion is None or len(df_proyeccion) == 0:
                self.log("No se pudieron leer datos del archivo de proyección", "ERROR")
                messagebox.showerror("Error", "No se pudieron leer datos del archivo de proyección.")
                return None
            
            self.log(f"Se procesarán {len(df_proyeccion)} registros", "INFO")
            
            # Preparar y agregar al archivo final
            df_segunda = self.preparar_datos_segunda_hoja(df_proyeccion)
            self.agregar_a_archivo_final(df_segunda)
            
            messagebox.showinfo(
                "¡Registros Anexados!",
                f"Los registros se han anexado exitosamente al archivo final.\n\n"
                f"📁 Archivo de proyección utilizado:\n{ruta_proyeccion.name}\n\n"
                f"📁 Archivo final actualizado:\n{self.ruta_destino_final.name}\n\n"
                f"📊 Total de registros anexados: {len(df_segunda)}"
            )
            return str(ruta_proyeccion)
            
        except Exception as e:
            self.log(f"ERROR CRÍTICO: {str(e)}", "ERROR")
            traceback.print_exc()
            messagebox.showerror("Error", f"Ocurrió un error:\n\n{str(e)}")
            return None
    
    def ejecutar_proceso(self):
        """Ejecuta el proceso según la opción seleccionada"""
        print("\n" + "="*80)
        print("    AUTOMATIZACIÓN DE CONTROL DE PAGOS - VERSIÓN 2.0")
        print("="*80 + "\n")
        
        opciones_texto = {
            1: "CREAR SOLO PROYECCIÓN",
            2: "ANEXAR SOLO A ARCHIVO FINAL",
            3: "PROCESO COMPLETO (PROYECCIÓN + ANEXAR)"
        }
        
        self.log(f"Opción seleccionada: {opciones_texto.get(self.opcion_proceso, 'DESCONOCIDA')}", "INFO")
        
        if self.opcion_proceso == 1:
            resultado = self.ejecutar_solo_proyeccion()
        elif self.opcion_proceso == 2:
            resultado = self.ejecutar_solo_anexar()
        elif self.opcion_proceso == 3:
            resultado = self.ejecutar_proceso_completo()
        else:
            self.log("Opción de proceso no válida", "ERROR")
            messagebox.showerror("Error", "Opción de proceso no válida.")
            return None
        
        print("\n" + "="*80)
        print("PROCESO COMPLETADO")
        print("="*80)
        
        return resultado

def main():
    """Función principal"""
    try:
        # Configurar rutas
        configurador = ConfiguradorRutas()
        if not configurador.cargar_o_crear_config():
            return
        
        rutas = configurador.obtener_rutas()
        
        # Mostrar ventana de selección
        interfaz = InterfazModerna()
        interfaz.crear_ventana()
        
        if not interfaz.ejecutar_proceso:
            return
        
        # Mensajes de confirmación según la opción
        mensajes_confirmacion = {
            1: "Antes de continuar, asegúrese de:\n\n"
               "✓ Haber actualizado el archivo 'CONTROL DE PAGOS.xlsm'\n"
               "✓ Haber guardado todos los cambios\n"
               "✓ Cerrar el archivo si está abierto\n\n"
               "Se creará la proyección semanal.\n\n"
               "¿Desea continuar?",
            2: "Antes de continuar, asegúrese de:\n\n"
               "✓ Tener el archivo de proyección disponible\n"
               "✓ El archivo 'CONTROL PAGOS.xlsx' está cerrado\n\n"
               "Se le solicitará seleccionar el archivo de proyección.\n"
               "Los registros se anexarán al archivo final.\n\n"
               "¿Desea continuar?",
            3: "Antes de continuar, asegúrese de:\n\n"
               "✓ Haber actualizado el archivo 'CONTROL DE PAGOS.xlsm'\n"
               "✓ Haber guardado todos los cambios\n"
               "✓ Cerrar todos los archivos Excel relacionados\n\n"
               "Se ejecutará el proceso completo.\n\n"
               "¿Desea continuar?"
        }
        
        mensaje = mensajes_confirmacion.get(interfaz.opcion_proceso, "¿Desea continuar?")
        
        if not messagebox.askyesno("Confirmar Ejecución", mensaje):
            return
        
        # Crear ventana de progreso
        ventana_prog = VentanaProgreso()
        
        try:
            # Ejecutar proceso
            copiador = CopiarArchivo(
                fecha_filtrado=interfaz.fecha_seleccionada,
                ventana_progreso=ventana_prog,
                opcion_proceso=interfaz.opcion_proceso,
                rutas_config=rutas
            )
            resultado = copiador.ejecutar_proceso()
            
            # Cerrar ventana de progreso
            ventana_prog.cerrar()
            
        except Exception as e:
            ventana_prog.cerrar()
            messagebox.showerror("Error Fatal", f"Error inesperado:\n\n{str(e)}")
    
    except Exception as e:
        messagebox.showerror("Error de Configuración", f"Error al iniciar:\n\n{str(e)}")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        import traceback
        error_msg = traceback.format_exc()
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        try:
            with open("CRASH_LOG.txt", "a", encoding='utf-8') as f:
                f.write(f"\n{'='*50}\n")
                f.write(f"FECHA: {timestamp}\n")
                f.write(f"ERROR:\n{error_msg}\n")
                f.write(f"{'='*50}\n")
        except:
            pass
        
        try:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("Error Fatal", f"Ocurrió un error crítico:\n\n{str(e)}\n\nConsulte CRASH_LOG.txt")
        except:
            print(f"Error fatal: {e}")
            input("Presione Enter para salir...")