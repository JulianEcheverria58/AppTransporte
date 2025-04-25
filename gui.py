import tkinter as tk
from tkinter import messagebox
from datetime import datetime
from tkcalendar import Calendar
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from PIL import Image, ImageTk
import os
import subprocess
import sys
from sharepoint_client import SharePointClient
from report_generator.base_report import ReportGenerator
from email_sender import EmailSender

class TransporteGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Sistema de Reportes MHC - Transporte")
        self.root.geometry("1000x700")
        
        # Configuración de estilo
        self.style = ttk.Style(theme='morph')
        self.style.configure('TButton', font=('Helvetica', 10))
        self.style.configure('TLabel', font=('Helvetica', 10))
        
        # Cargar recursos
        self._load_assets()
        
        # Variables de control
        self.fecha_var = tk.StringVar(value=datetime.now().strftime('%d/%m/%Y'))
        self.report_type = tk.StringVar(value='detallado')
        self.approval_var = tk.BooleanVar(value=True)
        
        # Clientes y servicios
        self.sharepoint_client = SharePointClient()
        self.report_generator = ReportGenerator()
        self.email_sender = EmailSender()
        
        # Construir interfaz
        self._setup_ui()
        self._center_window()
        self._log("Sistema listo. Reportes disponibles: " + 
                 ", ".join(self.report_generator.get_available_reports()), "info")

    def _load_assets(self):
        """Carga imágenes y recursos visuales"""
        try:
            logo_path = os.path.join(os.path.dirname(__file__), "logo.png")
            if os.path.exists(logo_path):
                img = Image.open(logo_path)
                img = img.resize((60, 60), Image.LANCZOS)
                self.logo_img = ImageTk.PhotoImage(img)
            else:
                self.logo_img = None
                
            try:
                approve_img = Image.open(os.path.join(os.path.dirname(__file__), "approve.png"))
                approve_img = approve_img.resize((20, 20), Image.LANCZOS)
                self.approve_icon = ImageTk.PhotoImage(approve_img)
            except:
                self.approve_icon = None
                
        except Exception as e:
            print(f"Error cargando recursos: {str(e)}")
            self.logo_img = None
            self.approve_icon = None

    def _center_window(self):
        """Centra la ventana en la pantalla"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')

    def _setup_ui(self):
        """Construye la interfaz de usuario"""
        # Frame principal
        main_frame = ttk.Frame(self.root, padding=(20, 15))
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Header
        self._build_header(main_frame)
        
        # Panel de configuración
        self._build_control_panel(main_frame)
        
        # Consola de registro
        self._build_console(main_frame)
        
        # Barra de estado
        self._build_status_bar(main_frame)

    def _build_header(self, parent):
        """Construye el encabezado con logo y título"""
        header_frame = ttk.Frame(parent)
        header_frame.pack(fill=tk.X, pady=(0, 20))
        
        if self.logo_img:
            ttk.Label(header_frame, image=self.logo_img).pack(side=tk.LEFT, padx=(0, 15))
        
        title_frame = ttk.Frame(header_frame)
        title_frame.pack(side=tk.LEFT, fill=tk.Y)
        
        ttk.Label(
            title_frame,
            text="Sistema de Reportes de Transporte",
            font=('Helvetica', 16, 'bold'),
            bootstyle=PRIMARY
        ).pack(anchor=tk.W)
        
        ttk.Label(
            title_frame,
            text="MHC - Gestión de Vales Diarios",
            font=('Helvetica', 11),
            bootstyle=SECONDARY
        ).pack(anchor=tk.W)

    def _build_control_panel(self, parent):
        """Construye el panel de controles"""
        control_frame = ttk.Labelframe(
            parent,
            text="Configuración del Reporte",
            padding=(20, 15),
            bootstyle=INFO
        )
        control_frame.pack(fill=tk.X, pady=(0, 20))
        
        # Sección de fecha
        self._build_date_selector(control_frame)
        
        # Sección de tipo de reporte
        self._build_report_selector(control_frame)
        
        # Opción de aprobación
        self._build_approval_option(control_frame)
        
        # Botones de acción
        self._build_action_buttons(control_frame)

    def _build_date_selector(self, parent):
        """Construye el selector de fecha"""
        date_frame = ttk.Frame(parent)
        date_frame.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Label(
            date_frame,
            text="Fecha del Reporte:",
            font=('Helvetica', 10)
        ).pack(side=tk.LEFT, padx=(0, 10))
        
        self.date_entry = ttk.Entry(
            date_frame,
            textvariable=self.fecha_var,
            width=15,
            font=('Helvetica', 10),
            bootstyle=PRIMARY
        )
        self.date_entry.pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(
            date_frame,
            text="Seleccionar Fecha",
            command=self._show_calendar,
            bootstyle=(OUTLINE, INFO),
            width=15
        ).pack(side=tk.LEFT)

    def _show_calendar(self):
        """Muestra el calendario para seleccionar fecha"""
        try:
            top = ttk.Toplevel(self.root)
            top.title("Seleccionar Fecha")
            top.geometry("350x300")
            top.resizable(False, False)
            top.transient(self.root)
            top.grab_set()
            
            cal = Calendar(
                top,
                selectmode='day',
                date_pattern='dd/mm/yyyy',
                locale='es_CO',
                font=('Helvetica', 10),
                background='white',
                foreground='black',
                selectbackground='#007bff',
                normalbackground='white',
                headersbackground='#f8f9fa',
                headersforeground='#495057'
            )
            cal.pack(padx=15, pady=15, fill=tk.BOTH, expand=True)
            
            btn_frame = ttk.Frame(top)
            btn_frame.pack(pady=(0, 15))
            
            ttk.Button(
                btn_frame,
                text="Seleccionar",
                command=lambda: self._set_date(cal.get_date(), top),
                bootstyle=SUCCESS
            ).pack(side=tk.LEFT, padx=10)
            
            ttk.Button(
                btn_frame,
                text="Cancelar",
                command=top.destroy,
                bootstyle=DANGER
            ).pack(side=tk.LEFT, padx=10)
            
        except Exception as e:
            self._log(f"Error al abrir calendario: {str(e)}", "error")
            messagebox.showerror("Error", f"No se pudo abrir el calendario: {str(e)}", parent=self.root)

    def _set_date(self, date_str, window):
        """Establece la fecha seleccionada"""
        self.fecha_var.set(date_str)
        window.destroy()
        self._log(f"Fecha seleccionada: {date_str}", "info")

    def _build_report_selector(self, parent):
        """Selector de tipo de reporte dinámico"""
        type_frame = ttk.Frame(parent)
        type_frame.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Label(
            type_frame,
            text="Tipo de Reporte:",
            font=('Helvetica', 10)
        ).pack(side=tk.LEFT, padx=(0, 10))
        
        reports = self.report_generator.get_available_reports()
        
        # Mapeo de nombres amigables
        friendly_names = {
            'daily_report': 'Detallado',
            'general_report': 'General',
            'f049_report': 'F-049'
        }
        
        for i, report in enumerate(reports):
            ttk.Radiobutton(
                type_frame,
                text=friendly_names.get(report, report.replace('_', ' ').title()),
                variable=self.report_type,
                value=report,
                bootstyle="info-toolbutton"
            ).pack(side=tk.LEFT, padx=5)
        
        # Establecer primer reporte como predeterminado
        if reports:
            self.report_type.set(reports[0])

    def _build_approval_option(self, parent):
        """Construye la opción para incluir aprobación"""
        approval_frame = ttk.Frame(parent)
        approval_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Checkbutton para incluir aprobación
        check = ttk.Checkbutton(
            approval_frame,
            text="Incluir botón de aprobación",
            variable=self.approval_var,
            bootstyle="info-round-toggle"
        )
        check.pack(side=tk.LEFT, padx=(0, 10))
        
        # Tooltip/info sobre la aprobación
        if self.approve_icon:
            ttk.Label(
                approval_frame,
                image=self.approve_icon,
                bootstyle=(INFO, INVERSE)
            ).pack(side=tk.LEFT)

    def _build_action_buttons(self, parent):
        """Botones para generar y enviar reportes"""
        btn_frame = ttk.Frame(parent)
        btn_frame.pack(fill=tk.X)
        
        self.generate_btn = ttk.Button(
            btn_frame,
            text="Generar Reporte",
            command=self._generate_report,
            bootstyle=SUCCESS,
            width=20
        )
        self.generate_btn.pack(side=tk.LEFT, padx=5, pady=10)
        
        self.send_btn = ttk.Button(
            btn_frame,
            text="Enviar por Email",
            command=self._send_report,
            bootstyle=INFO,
            width=20
        )
        self.send_btn.pack(side=tk.LEFT, padx=5, pady=10)
        
        self.open_btn = ttk.Button(
            btn_frame,
            text="Abrir Reporte",
            command=self._open_report,
            bootstyle=WARNING,
            width=20
        )
        self.open_btn.pack(side=tk.LEFT, padx=5, pady=10)

    def _build_console(self, parent):
        """Construye la consola de registro"""
        console_frame = ttk.Labelframe(
            parent,
            text="Registro de Actividades",
            padding=(15, 10),
            bootstyle=INFO
        )
        console_frame.pack(fill=tk.BOTH, expand=True)
        
        # Configurar texto con scroll
        self.console = tk.Text(
            console_frame,
            height=15,
            wrap=tk.WORD,
            font=('Consolas', 9),
            bg='#f8f9fa',
            fg='#212529',
            padx=10,
            pady=10,
            state=tk.DISABLED
        )
        self.console.pack(fill=tk.BOTH, expand=True)
        
        # Configurar scrollbar
        scrollbar = ttk.Scrollbar(
            console_frame,
            orient=tk.VERTICAL,
            command=self.console.yview,
            bootstyle=ROUND
        )
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.console.config(yscrollcommand=scrollbar.set)
        
        # Configurar tags para diferentes niveles de log
        self.console.tag_config("info", foreground="#007bff")
        self.console.tag_config("success", foreground="#28a745")
        self.console.tag_config("warning", foreground="#fd7e14")
        self.console.tag_config("error", foreground="#dc3545")

    def _build_status_bar(self, parent):
        """Construye la barra de estado"""
        self.status_var = tk.StringVar(value="Listo")
        status_bar = ttk.Label(
            parent,
            textvariable=self.status_var,
            relief=tk.SUNKEN,
            anchor=tk.W,
            font=('Helvetica', 9),
            bootstyle=(SECONDARY, INVERSE)
        )
        status_bar.pack(fill=tk.X, pady=(10, 0))

    def _log(self, message, level="info"):
        """Escribe mensajes en la consola con formato"""
        self.console.config(state=tk.NORMAL)
        self.console.insert(tk.END, message + "\n", level)
        self.console.see(tk.END)
        self.console.config(state=tk.DISABLED)
        self.status_var.set(message)
        self.root.update()

    def _disable_ui(self):
        """Deshabilita los controles durante operaciones"""
        self.generate_btn.config(state=tk.DISABLED)
        self.send_btn.config(state=tk.DISABLED)
        self.open_btn.config(state=tk.DISABLED)
        self.date_entry.config(state=tk.DISABLED)
        self.root.config(cursor="watch")
        self.root.update()

    def _enable_ui(self):
        """Habilita los controles nuevamente"""
        self.generate_btn.config(state=tk.NORMAL)
        self.send_btn.config(state=tk.NORMAL)
        self.open_btn.config(state=tk.NORMAL)
        self.date_entry.config(state=tk.NORMAL)
        self.root.config(cursor="")
        self.root.update()

    def _generate_report(self):
        """Genera el reporte seleccionado"""
        try:
            self._disable_ui()
            fecha_str = self.fecha_var.get()
            
            # Validar formato de fecha (corrección para el error reportado)
            try:
                fecha = datetime.strptime(fecha_str, '%d/%m/%Y').date()
            except ValueError:
                try:
                    # Intentar otro formato común si falla el primero
                    fecha = datetime.strptime(fecha_str, '%d-%m-%Y').date()
                    self.fecha_var.set(fecha.strftime('%d/%m/%Y'))  # Actualizar al formato estándar
                except ValueError as e:
                    raise ValueError("Formato de fecha no válido. Use DD/MM/AAAA o DD-MM-AAAA")
            
            report_type = self.report_type.get()
            
            self._log(f"Generando reporte {report_type} para {fecha_str}...", "info")
            self._log(f"Parámetros: Fecha={fecha_str}, Tipo={report_type}, Aprobación={'Sí' if self.approval_var.get() else 'No'}", "info")
            
            # Obtener datos optimizados para el reporte
            self._log("Conectando a SharePoint...", "info")
            df = self.sharepoint_client.get_items_for_report(report_type, fecha)
            
            if df.empty:
                self._log("No hay datos para la fecha seleccionada", "warning")
                messagebox.showwarning("Sin datos", "No se encontraron registros", parent=self.root)
                return
            
            self._log(f"✓ {len(df)} registros encontrados", "success")
            
            # Generar reporte
            self._log("Generando reporte...", "info")
            output = self.report_generator.generate(
                report_type, 
                df, 
                report_date=fecha,
                incluir_aprobacion=self.approval_var.get()
            )
            
            if report_type == 'f049_report':
                self.current_report = output  # Guarda la ruta del PDF
                self._log(f"PDF generado: {output}", "success")
                messagebox.showinfo("Éxito", f"Informe F-049 generado:\n{output}", parent=self.root)
            else:
                self.current_report = output  # Guarda el HTML
                self._log("Reporte generado", "success")
                
        except ValueError as e:
            self._log(f"✗ Error en formato de fecha: {str(e)}", "error")
            messagebox.showerror("Error", str(e), parent=self.root)
        except Exception as e:
            self._log(f"✗ Error inesperado: {str(e)}", "error")
            messagebox.showerror("Error", f"Ocurrió un error inesperado: {str(e)}", parent=self.root)
        finally:
            self._enable_ui()
            self._log("Proceso completado", "info")
            self.status_var.set("Listo")

    def _send_report(self):
        """Envía el reporte por email"""
        if not hasattr(self, 'current_report'):
            messagebox.showwarning("Advertencia", "Primero genere un reporte", parent=self.root)
            return
            
        try:
            fecha_str = self.fecha_var.get()
            report_type = self.report_type.get()
            
            if report_type == 'f049_report':
                # Para F-049, adjuntar el PDF
                success, msg = self.email_sender.send_with_attachment(
                    f"Informe F-049 - {fecha_str}",
                    f"Se adjunta el informe F-049 para {fecha_str}",
                    self.current_report
                )
            else:
                # Para otros reportes, enviar HTML en el cuerpo
                fecha = datetime.strptime(fecha_str, '%d/%m/%Y').date()
                success, msg = self.email_sender.send_report(
                    self.current_report, 
                    fecha,
                    report_type
                )
            
            if success:
                self._log(f"✓ {msg}", "success")
                messagebox.showinfo("Éxito", "Reporte enviado correctamente", parent=self.root)
            else:
                self._log(f"✗ {msg}", "error")
                messagebox.showerror("Error", msg, parent=self.root)
                
        except Exception as e:
            self._log(f"Error al enviar: {str(e)}", "error")
            messagebox.showerror("Error", f"Error al enviar email: {str(e)}", parent=self.root)

    def _open_report(self):
        """Abre el reporte generado"""
        if not hasattr(self, 'current_report'):
            messagebox.showwarning("Advertencia", "Primero genere un reporte", parent=self.root)
            return
            
        try:
            if isinstance(self.current_report, str) and self.current_report.endswith('.pdf'):
                # Abrir PDF
                if os.name == 'nt':  # Windows
                    os.startfile(self.current_report)
                elif os.name == 'posix':  # Mac/Linux
                    subprocess.run(['open' if sys.platform == 'darwin' else 'xdg-open', self.current_report])
            else:
                # Abrir HTML en navegador
                import webbrowser
                webbrowser.open_new_tab(self.current_report if isinstance(self.current_report, str) else "data:text/html," + self.current_report)
                
        except Exception as e:
            self._log(f"Error al abrir reporte: {str(e)}", "error")
            messagebox.showerror("Error", f"No se pudo abrir el reporte: {str(e)}", parent=self.root)