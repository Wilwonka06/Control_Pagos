from pathlib import Path
from datetime import datetime, timedelta
import locale
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry
import configparser
import sys

# --- PATCH: Silenciar errores de compatibilidad de Pandas/Dateutil en consola ---
class StderrFilter:
    def __init__(self, original_stderr):
        self.original_stderr = original_stderr
        self.buffer = ""

    def write(self, s):
        self.buffer += s
        if "\n" in self.buffer or len(self.buffer) > 500:
            self.process_buffer()

    def flush(self):
        if self.buffer:
            self.process_buffer(force=True)
        try:
            self.original_stderr.flush()
        except:
            pass
            
    def process_buffer(self, force=False):
        error_signatures = [
            "pandas._libs.tslibs", 
            "total_seconds", 
            "_localize_tso",
            "AttributeError: 'NoneType' object"
        ]
        
        if any(sig in self.buffer for sig in error_signatures):
            self.buffer = ""
            return

        suspicious_starts = [
            "Exception ignored in:", 
            "Traceback (most recent call last):", 
            "AttributeError:"
        ]
        
        is_suspicious = any(start in self.buffer for start in suspicious_starts)
        
        if is_suspicious and len(self.buffer) < 300 and not force:
            return

        try:
            self.original_stderr.write(self.buffer)
        except:
            pass
        finally:
            self.buffer = ""

sys.stderr = StderrFilter(sys.stderr)

# Configuración de español
try:
    locale.setlocale(locale.LC_TIME, 'Spanish_Spain.1252')  # Windows
except locale.Error:
    try:
        locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')  # Linux
    except locale.Error:
        pass

def obtener_ruta_recurso(nombre_archivo: str) -> Path:
    base_dir = getattr(sys, "_MEIPASS", None)
    if base_dir:
        return Path(base_dir) / nombre_archivo
    return Path(__file__).resolve().parent / nombre_archivo

def aplicar_icono_ventana(ventana: tk.Tk) -> None:
    try:
        icon_path = obtener_ruta_recurso("icon.ico")
        if icon_path.exists():
            ventana.iconbitmap(str(icon_path))
    except Exception:
        pass


class ConfiguradorRutas:
    """Manejador de configuración de rutas"""

    def __init__(self):
        self.config_file = Path("config_pagos.ini")
        self.config = configparser.ConfigParser()
        
    def cargar_o_crear_config(self):
        if self.config_file.exists():
            self.config.read(self.config_file, encoding='utf-8')
            return True
        else:
            return self.crear_configuracion_inicial()
    
    def crear_configuracion_inicial(self):
        root = tk.Tk()
        aplicar_icono_ventana(root)
        root.withdraw()
        
        messagebox.showinfo(
            "Primera Configuración",
            "Por favor, seleccione las rutas necesarias para el programa."
        )
        
        messagebox.showinfo("Paso 1", "Seleccione el archivo CONTROL DE PAGOS de comercio")
        archivo_origen = filedialog.askopenfilename(
            title="Seleccionar CONTROL DE PAGOS.xlsm",
            filetypes=[("Excel Macro", "*.xlsm"), ("Todos", "*.*")]
        )
        
        if not archivo_origen:
            messagebox.showerror("Error", "Debe seleccionar el archivo origen.")
            return False
        
        messagebox.showinfo("Paso 2", "Seleccione la carpeta donde se guardarán las PROYECCIONES")
        carpeta_proyecciones = filedialog.askdirectory(
            title="Seleccionar carpeta de PROYECCIONES"
        )
        
        if not carpeta_proyecciones:
            messagebox.showerror("Error", "Debe seleccionar la carpeta de proyecciones.")
            return False
        
        messagebox.showinfo("Paso 3", "Seleccione el archivo CONTROL PAGOS.xlsx - archivo final")
        archivo_final = filedialog.askopenfilename(
            title="Seleccionar CONTROL PAGOS.xlsx",
            filetypes=[("Excel", "*.xlsx"), ("Todos", "*.*")],
            initialdir=carpeta_proyecciones
        )
        
        if not archivo_final:
            messagebox.showerror("Error", "Debe seleccionar el archivo final.")
            return False
        
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
        return {
            'origen': Path(self.config['RUTAS']['archivo_origen']),
            'proyecciones': Path(self.config['RUTAS']['carpeta_proyecciones']),
            'final': Path(self.config['RUTAS']['archivo_final'])
        }


class VentanaSeleccionTipo:
    """Ventana inicial para seleccionar el tipo de proyección"""
    
    def __init__(self):
        self.tipo_seleccionado = None
        
        # Modern color palette
        self.COLOR_PRIMARIO = "#2C3E50"
        self.COLOR_SECUNDARIO = "#3498DB"
        self.COLOR_SEMANAL = "#3498DB"
        self.COLOR_MENSUAL = "#9B59B6"
        self.COLOR_ACENTO = "#27AE60"
        self.COLOR_FONDO = "#ECF0F1"
        self.COLOR_CARTA = "#FFFFFF"
        self.COLOR_TEXTO = "#2C3E50"
        self.COLOR_TEXTO_CLARO = "#7F8C8D"
        self.COLOR_BORDE = "#BDC3C7"
        
    def crear_ventana(self):
        self.root = tk.Tk()
        aplicar_icono_ventana(self.root)
        self.root.title("Control de Pagos GCO - Selector")
        self.root.geometry("650x500")
        self.root.resizable(False, False)
        self.root.configure(bg=self.COLOR_FONDO)
        
        self.centrar_ventana()
        
        # Header with gradient effect
        header_frame = tk.Frame(self.root, bg=self.COLOR_PRIMARIO, height=120)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        # Header content
        header_content = tk.Frame(header_frame, bg=self.COLOR_PRIMARIO)
        header_content.place(relx=0.5, rely=0.5, anchor="center")
        
        # Logo/Icon
        icon_label = tk.Label(
            header_content,
            text="💰",
            font=("Segoe UI", 32),
            bg=self.COLOR_PRIMARIO,
            fg="white"
        )
        icon_label.pack(side=tk.LEFT, padx=(0, 15))
        
        # Title text
        title_frame = tk.Frame(header_content, bg=self.COLOR_PRIMARIO)
        title_frame.pack(side=tk.LEFT)
        
        title_label = tk.Label(
            title_frame,
            text="CONTROL DE PAGOS GCO",
            font=("Segoe UI", 22, "bold"),
            bg=self.COLOR_PRIMARIO,
            fg="white"
        )
        title_label.pack(anchor="w")
        
        subtitle_label = tk.Label(
            title_frame,
            text="Sistema de Gestión de Importaciones",
            font=("Segoe UI", 10),
            bg=self.COLOR_PRIMARIO,
            fg="#BDC3C7"
        )
        subtitle_label.pack(anchor="w")
        
        # Main content area
        content_frame = tk.Frame(self.root, bg=self.COLOR_FONDO)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=40, pady=30)
        
        # Selection prompt
        label_titulo = tk.Label(
            content_frame,
            text="Seleccione el tipo de proyección:",
            font=("Segoe UI", 14, "bold"),
            bg=self.COLOR_FONDO,
            fg=self.COLOR_TEXTO
        )
        label_titulo.pack(pady=(0, 25))
        
        # Card-based selection buttons
        cards_frame = tk.Frame(content_frame, bg=self.COLOR_FONDO)
        cards_frame.pack(fill=tk.BOTH, expand=True)
        
        # Weekly Card
        weekly_card = self.crear_card_seleccion(
            cards_frame,
            titulo="📅 PROYECCIÓN SEMANAL",
            color=self.COLOR_SEMANAL,
            command=lambda: self.seleccionar_tipo("SEMANAL")
        )
        weekly_card.pack(fill=tk.X, pady=(0, 15))
        
        # Monthly Card
        monthly_card = self.crear_card_seleccion(
            cards_frame,
            titulo="📊 PROYECCIÓN MENSUAL",
            color=self.COLOR_MENSUAL,
            command=lambda: self.seleccionar_tipo("MENSUAL")
        )
        monthly_card.pack(fill=tk.X, pady=(0, 15))
        
        # Footer with cancel button
        footer_frame = tk.Frame(self.root, bg=self.COLOR_FONDO)
        footer_frame.pack(fill=tk.X, padx=40, pady=(0, 20))
        
        btn_cancelar = tk.Button(
            footer_frame,
            text="✕ Cancelar",
            font=("Segoe UI", 10),
            bg="white",
            fg=self.COLOR_TEXTO_CLARO,
            activebackground="#E74C3C",
            activeforeground="white",
            relief=tk.FLAT,
            borderwidth=2,
            cursor="hand2",
            padx=25,
            pady=8,
            command=self.cancelar
        )
        btn_cancelar.pack(side=tk.RIGHT)
        
        # Hover effect for cancel button
        def on_enter_cancel(e):
            btn_cancelar.configure(bg="#E74C3C", fg="white")
        def on_leave_cancel(e):
            btn_cancelar.configure(bg="white", fg=self.COLOR_TEXTO_CLARO)
        btn_cancelar.bind("<Enter>", on_enter_cancel)
        btn_cancelar.bind("<Leave>", on_leave_cancel)
        
        self.root.protocol("WM_DELETE_WINDOW", self.cancelar)
        self.root.mainloop()
    
    def crear_card_seleccion(self, parent, titulo, color, command):
        """Creates a card-style selection button with hover effect"""
        card_frame = tk.Frame(parent, bg=self.COLOR_CARTA, relief=tk.FLAT, borderwidth=0)
        
        # Add shadow effect (sombra debajo de la tarjeta)
        for i in range(3):
            shade = tk.Frame(
                parent,
                bg=f"#{230-i*10:02x}{230-i*10:02x}{230-i*10:02x}"
            )
            shade.place(in_=card_frame, x=i+2, y=i+2, relwidth=1, relheight=1)
            shade.lower(card_frame)
        
        # Main card content
        card_inner = tk.Frame(card_frame, bg=self.COLOR_CARTA, cursor="hand2", relief=tk.FLAT)
        card_inner.pack(fill=tk.BOTH, expand=True, padx=3, pady=3)
        card_inner.pack_propagate(False)
        card_inner.configure(height=80)
        
        # Left color accent bar
        accent_bar = tk.Frame(card_inner, bg=color, width=6)
        accent_bar.pack(side=tk.LEFT, fill=tk.Y)
        
        # Content
        content = tk.Frame(card_inner, bg=self.COLOR_CARTA)
        content.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=20, pady=15)
        
        title_label = tk.Label(
            content,
            text=titulo,
            font=("Segoe UI", 13, "bold"),
            bg=self.COLOR_CARTA,
            fg=color,
            anchor="w"
        )
        title_label.pack(anchor="w")
        
        desc_label = tk.Label(
            content,
            font=("Segoe UI", 9),
            bg=self.COLOR_CARTA,
            fg=self.COLOR_TEXTO_CLARO,
            anchor="w"
        )
        desc_label.pack(anchor="w", pady=(3, 0))
        
        # Arrow indicator
        arrow_label = tk.Label(
            card_inner,
            text="➔",
            font=("Segoe UI", 16),
            bg=self.COLOR_CARTA,
            fg=color
        )
        arrow_label.pack(side=tk.RIGHT, padx=20)
        
        # Bind click and hover events
        card_inner.bind("<Button-1>", lambda e: command())
        card_inner.bind("<Enter>", lambda e: self.on_card_hover(card_frame, card_inner, color, True))
        card_inner.bind("<Leave>", lambda e: self.on_card_hover(card_frame, card_inner, color, False))
        
        # Bind to all children
        for child in card_inner.winfo_children():
            child.bind("<Button-1>", lambda e: command())
            child.bind("<Enter>", lambda e, c=card_frame, ci=card_inner, col=color: self.on_card_hover(c, ci, col, True))
            child.bind("<Leave>", lambda e, c=card_frame, ci=card_inner, col=color: self.on_card_hover(c, ci, col, False))
        
        return card_frame
    
    def on_card_hover(self, card_frame, card_inner, color, entering):
        """Handle hover effect on cards"""
        if entering:
            card_inner.configure(bg="#F8F9FA")
            for child in card_inner.winfo_children():
                try:
                    child.configure(bg="#F8F9FA")
                except:
                    pass
        else:
            card_inner.configure(bg=self.COLOR_CARTA)
            for child in card_inner.winfo_children():
                try:
                    child.configure(bg=self.COLOR_CARTA)
                except:
                    pass
    
    def centrar_ventana(self):
        self.root.update_idletasks()
        ancho = self.root.winfo_width()
        alto = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (ancho // 2)
        y = (self.root.winfo_screenheight() // 2) - (alto // 2)
        self.root.geometry(f'{ancho}x{alto}+{x}+{y}')
    
    def seleccionar_tipo(self, tipo):
        self.tipo_seleccionado = tipo
        self.root.destroy()
    
    def cancelar(self):
        self.tipo_seleccionado = None
        self.root.destroy()


class VentanaProgreso:
    """Ventana de progreso durante la ejecución"""
    
    def __init__(self):
        self.root = tk.Tk()
        aplicar_icono_ventana(self.root)
        self.root.title("Procesando...")
        self.root.geometry("650x400")
        self.root.resizable(False, False)
        self.root.configure(bg="#ECF0F1")
        
        # Modern color palette
        self.COLOR_PRIMARIO = "#2C3E50"
        self.COLOR_SECUNDARIO = "#3498DB"
        self.COLOR_ACENTO = "#27AE60"
        self.COLOR_FONDO = "#ECF0F1"
        self.COLOR_CARTA = "#FFFFFF"
        self.COLOR_TEXTO = "#2C3E50"
        self.COLOR_TEXTO_CLARO = "#7F8C8D"
        
        self.centrar_ventana()
        
        # Header
        header_frame = tk.Frame(self.root, bg=self.COLOR_PRIMARIO, height=70)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        header_content = tk.Frame(header_frame, bg=self.COLOR_PRIMARIO)
        header_content.place(relx=0.5, rely=0.5, anchor="center")
        
        icon_label = tk.Label(
            header_content,
            text="⚙️",
            font=("Segoe UI", 20),
            bg=self.COLOR_PRIMARIO,
            fg="white"
        )
        icon_label.pack(side=tk.LEFT, padx=(0, 10))
        
        title_label = tk.Label(
            header_content,
            text="Procesando Control de Pagos",
            font=("Segoe UI", 14, "bold"),
            bg=self.COLOR_PRIMARIO,
            fg="white"
        )
        title_label.pack(side=tk.LEFT)
        
        # Main content frame
        main_frame = tk.Frame(self.root, bg=self.COLOR_FONDO)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=25, pady=20)
        
        # Progress card
        progress_card = tk.Frame(main_frame, bg=self.COLOR_CARTA, relief=tk.FLAT)
        progress_card.pack(fill=tk.X)
        
        # Add shadow
        for i in range(3):
            shade = tk.Frame(main_frame, bg=f"#{230-i*10:02x}{230-i*10:02x}{230-i*10:02x}")
            shade.place(x=i+2, y=i+2, relwidth=1, relheight=0.25)
            shade.lower(progress_card)
        
        progress_inner = tk.Frame(progress_card, bg=self.COLOR_CARTA)
        progress_inner.pack(fill=tk.X, padx=3, pady=15)
        
        # Status label
        self.status_label = tk.Label(
            progress_inner,
            text="Iniciando proceso...",
            font=("Segoe UI", 11, "bold"),
            bg=self.COLOR_CARTA,
            fg=self.COLOR_PRIMARIO,
            anchor="w"
        )
        self.status_label.pack(fill=tk.X, padx=15, pady=(0, 10))
        
        # Progress bar frame with custom styling
        progress_bar_frame = tk.Frame(progress_inner, bg="#E8E8E8", relief=tk.FLAT)
        progress_bar_frame.pack(fill=tk.X, padx=15, pady=5)
        progress_bar_frame.configure(height=20)
        
        # Custom progress canvas
        self.progress_canvas = tk.Canvas(
            progress_bar_frame,
            bg="#E8E8E8",
            highlightthickness=0,
            relief=tk.FLAT,
            height=20
        )
        self.progress_canvas.pack(fill=tk.X, padx=0, pady=0)
        
        # Progress fill rectangle
        self.progress_fill = self.progress_canvas.create_rectangle(0, 0, 0, 20, fill=self.COLOR_SECUNDARIO, width=0)
        self.progress_bg = self.progress_canvas.create_rectangle(0, 0, 1000, 20, fill="#E0E0E0", width=0)
        
        # Percentage label
        self.percent_label = tk.Label(
            progress_inner,
            text="0%",
            font=("Segoe UI", 10, "bold"),
            bg=self.COLOR_CARTA,
            fg=self.COLOR_SECUNDARIO
        )
        self.percent_label.pack(pady=(5, 0))
        
        # Log area card
        log_card = tk.Frame(main_frame, bg=self.COLOR_CARTA, relief=tk.FLAT)
        log_card.pack(fill=tk.BOTH, expand=True, pady=(15, 0))
        
        # Add shadow to log card
        for i in range(3):
            shade = tk.Frame(main_frame, bg=f"#{230-i*10:02x}{230-i*10:02x}{230-i*10:02x}")
            shade.place(x=i+2, y=i+2+130, relwidth=1, relheight=0.5)
            shade.lower(log_card)
        
        log_inner = tk.Frame(log_card, bg=self.COLOR_CARTA)
        log_inner.pack(fill=tk.BOTH, expand=True, padx=3, pady=3)
        
        # Log header
        log_header = tk.Frame(log_inner, bg=self.COLOR_CARTA)
        log_header.pack(fill=tk.X, padx=15, pady=(10, 5))
        
        log_icon = tk.Label(
            log_header,
            text="📋",
            font=("Segoe UI", 12),
            bg=self.COLOR_CARTA
        )
        log_icon.pack(side=tk.LEFT)
        
        log_title = tk.Label(
            log_header,
            text="Registro de Actividad",
            font=("Segoe UI", 10, "bold"),
            bg=self.COLOR_CARTA,
            fg=self.COLOR_TEXTO
        )
        log_title.pack(side=tk.LEFT, padx=(5, 0))
        
        # Log text area
        log_text_frame = tk.Frame(log_inner, bg="#F5F5F5", relief=tk.FLAT)
        log_text_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 10))
        
        self.log_text = tk.Text(
            log_text_frame,
            font=("Consolas", 9),
            bg="#F5F5F5",
            fg=self.COLOR_TEXTO,
            relief=tk.FLAT,
            state=tk.DISABLED,
            wrap=tk.WORD
        )
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Scrollbar for log
        scrollbar = tk.Scrollbar(log_text_frame, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)
        
        # Animation state
        self.animating = True
        self.animation_pos = 0
        self.animate_progress()
        
        self.root.update()
    
    def animate_progress(self):
        """Animate the progress bar"""
        if not self.animating:
            return
        
        self.animation_pos = (self.animation_pos + 2) % 100
        
        # Update progress bar visually
        canvas_width = self.progress_canvas.winfo_width()
        if canvas_width < 10:
            canvas_width = 560  # Default width
        
        fill_width = int(canvas_width * (self.animation_pos / 100))
        self.progress_canvas.coords(self.progress_fill, 0, 0, fill_width, 20)
        
        # Update percentage
        self.percent_label.config(text=f"{self.animation_pos}%")
        
        # Schedule next animation frame
        self.root.after(50, self.animate_progress)
    
    def set_progress(self, percent, status_text=""):
        """Set progress percentage and status text"""
        if status_text:
            self.status_label.config(text=status_text)
        
        # Update progress bar
        canvas_width = self.progress_canvas.winfo_width()
        if canvas_width < 10:
            canvas_width = 560
        
        fill_width = int(canvas_width * (percent / 100))
        self.progress_canvas.coords(self.progress_fill, 0, 0, fill_width, 20)
        self.percent_label.config(text=f"{percent}%")
        
        if percent >= 100:
            self.percent_label.config(fg=self.COLOR_ACENTO)
        
        self.root.update()
    
    def centrar_ventana(self):
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def log(self, mensaje, tipo="INFO"):
        """Agrega mensaje al log"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        iconos = {
            "INFO": "ℹ️",
            "OK": "✅",
            "WARN": "⚠️",
            "ERROR": "❌",
            "PROCESS": "🔄"
        }
        
        colores = {
            "INFO": "#3498DB",
            "OK": "#27AE60",
            "WARN": "#F39C12",
            "ERROR": "#E74C3C",
            "PROCESS": "#9B59B6"
        }
        
        icono = iconos.get(tipo, "ℹ️")
        color = colores.get(tipo, "#2C3E50")
        
        # Enable text widget
        self.log_text.config(state=tk.NORMAL)
        
        # Insert timestamp and icon
        self.log_text.insert(tk.END, f"[{timestamp}] ", "timestamp")
        self.log_text.insert(tk.END, f"{icono} ", f"icon_{tipo}")
        self.log_text.insert(tk.END, f"{mensaje}\n", "message")
        
        # Configure tags
        self.log_text.tag_config("timestamp", foreground="#95A5A6", font=("Consolas", 9))
        self.log_text.tag_config(f"icon_{tipo}", foreground=color, font=("Consolas", 9))
        self.log_text.tag_config("message", foreground=self.COLOR_TEXTO, font=("Consolas", 9))
        
        # Scroll to bottom
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        
        # Update status label with last message
        if tipo == "ERROR":
            self.status_label.config(text=f"❌ Error: {mensaje}", fg="#E74C3C")
        elif tipo == "OK":
            self.status_label.config(text=f"✅ {mensaje}", fg="#27AE60")
        else:
            self.status_label.config(text=mensaje, fg=self.COLOR_PRIMARIO)
        
        self.root.update()
    
    def cerrar(self):
        """Cierra la ventana"""
        self.animating = False
        self.root.destroy()


def main():
    """Función principal unificada"""
    # Cargar configuración
    configurador = ConfiguradorRutas()
    if not configurador.cargar_o_crear_config():
        return
    
    rutas = configurador.obtener_rutas()
    
    # Bucle principal
    while True:
        try:
            # 1. Seleccionar tipo de proyección
            selector = VentanaSeleccionTipo()
            selector.crear_ventana()
            
            if selector.tipo_seleccionado is None:
                break
            
            tipo_proyeccion = selector.tipo_seleccionado
            
            # 2. Importar módulo correspondiente
            if tipo_proyeccion == "SEMANAL":
                import proceso_semanal as proceso_semanal
                
                # Ejecutar interfaz semanal
                interfaz = proceso_semanal.InterfazSemanal()
                interfaz.crear_ventana()
                
                if not interfaz.ejecutar_proceso:
                    continue
                
                # Confirmar ejecución
                if not messagebox.askyesno(
                    "Confirmar Ejecución",
                    "Antes de continuar, asegúrese de:\n\n"
                    "   ✓ Cerrar el archivo si está abierto\n\n"
                    "¿Desea continuar?"
                ):
                    continue
                
                # Ejecutar proceso
                ventana_prog = VentanaProgreso()
                try:
                    procesador = proceso_semanal.ProcesadorSemanal(
                        fecha_filtrado=interfaz.fecha_seleccionada,
                        ventana_progreso=ventana_prog,
                        rutas_config=rutas
                    )
                    resultado = procesador.ejecutar_proceso()
                    ventana_prog.cerrar()
                except Exception as e:
                    ventana_prog.cerrar()
                    messagebox.showerror("Error Fatal", f"Error inesperado:\n\n{str(e)}")
            
            elif tipo_proyeccion == "MENSUAL":
                import proceso_mensual as proceso_mensual
                
                # Ejecutar interfaz mensual
                interfaz = proceso_mensual.InterfazMensual()
                interfaz.crear_ventana()
                
                if not interfaz.ejecutar_proceso:
                    continue
                
                # Ejecutar proceso
                ventana_prog = VentanaProgreso()
                try:
                    procesador = proceso_mensual.ProcesadorMensual(
                        fecha_filtrado=interfaz.fecha_seleccionada,
                        ventana_progreso=ventana_prog,
                        rutas_config=rutas
                    )
                    resultado = procesador.ejecutar_proceso()
                    ventana_prog.cerrar()
                except Exception as e:
                    ventana_prog.cerrar()
                    messagebox.showerror("Error Fatal", f"Error inesperado:\n\n{str(e)}")
            
            # Preguntar si desea procesar otra proyección
            if not messagebox.askyesno(
                "Proceso Completado",
                "¿Desea realizar otra proyección?"
            ):
                break
        
        except Exception as e:
            messagebox.showerror("Error de Configuración", f"Error al iniciar:\n\n{str(e)}")
            if not messagebox.askyesno("Error", "¿Desea intentar nuevamente?"):
                break


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
            aplicar_icono_ventana(root)
            root.withdraw()
            messagebox.showerror("Error Fatal", f"Ocurrió un error crítico:\n\n{str(e)}\n\nConsulte CRASH_LOG.txt")
        except:
            print(f"Error fatal: {e}")
            input("Presione Enter para salir...")
