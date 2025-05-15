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
        self._configurar_ventana()
        self._inicializar_variables()
        self._cargar_recursos()
        self._construir_interfaz()
        self._log("Sistema listo. Reportes disponibles: " + 
                 ", ".join(self.report_generator.get_available_reports()), "info")

    def _configurar_ventana(self):
        """Configura los parámetros básicos de la ventana principal"""
        self.root.title("Sistema de Reportes MHC - Transporte")
        self.root.geometry("1000x700")
        self.style = ttk.Style(theme='morph')
        self.style.configure('TButton', font=('Helvetica', 10))
        self.style.configure('TLabel', font=('Helvetica', 10))

    def _inicializar_variables(self):
        """Inicializa las variables de control"""
        self.fecha_var = tk.StringVar(value=datetime.now().strftime('%d/%m/%Y'))
        self.report_type = tk.StringVar(value='daily_report')
        self.sharepoint_client = SharePointClient()
        self.report_generator = ReportGenerator()
        self.email_sender = EmailSender()

    def _cargar_recursos(self):
        """Carga imágenes y recursos visuales"""
        try:
            logo_path = os.path.join(os.path.dirname(__file__), "logo.png")
            if os.path.exists(logo_path):
                img = Image.open(logo_path)
                self.logo_img = ImageTk.PhotoImage(img.resize((60, 60), Image.LANCZOS))
            else:
                self.logo_img = None
        except Exception as e:
            print(f"Error cargando logo: {str(e)}")
            self.logo_img = None

    def _construir_interfaz(self):
        """Construye todos los componentes de la interfaz"""
        main_frame = ttk.Frame(self.root, padding=(20, 15))
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        self._construir_encabezado(main_frame)
        self._construir_panel_control(main_frame)
        self._construir_consola(main_frame)
        self._construir_barra_estado(main_frame)
        self._center_window()

    def _construir_encabezado(self, parent):
        """Construye la sección de encabezado con logo y título"""
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

    def _construir_panel_control(self, parent):
        """Construye el panel de control principal"""
        control_frame = ttk.Labelframe(
            parent,
            text="Configuración del Reporte",
            padding=(20, 15),
            bootstyle=INFO
        )
        control_frame.pack(fill=tk.X, pady=(0, 20))
        
        self._construir_selector_fecha(control_frame)
        self._construir_selector_reporte(control_frame)
        self._construir_botones_accion(control_frame)

    def _construir_selector_fecha(self, parent):
        """Construye el selector de fecha"""
        frame = ttk.Frame(parent)
        frame.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Label(
            frame,
            text="Fecha del Reporte:",
            font=('Helvetica', 10)
        ).pack(side=tk.LEFT, padx=(0, 10))
        
        self.date_entry = ttk.Entry(
            frame,
            textvariable=self.fecha_var,
            width=15,
            font=('Helvetica', 10),
            bootstyle=PRIMARY
        )
        self.date_entry.pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(
            frame,
            text="Seleccionar Fecha",
            command=self._mostrar_calendario,
            bootstyle=(OUTLINE, INFO),
            width=15
        ).pack(side=tk.LEFT)

    def _mostrar_calendario(self):
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
                font=('Helvetica', 10)
            )
            cal.pack(padx=15, pady=15, fill=tk.BOTH, expand=True)
            
            btn_frame = ttk.Frame(top)
            btn_frame.pack(pady=(0, 15))
            
            ttk.Button(
                btn_frame,
                text="Seleccionar",
                command=lambda: self._establecer_fecha(cal.get_date(), top),
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

    def _establecer_fecha(self, date_str, window):
        """Establece la fecha seleccionada"""
        self.fecha_var.set(date_str)
        window.destroy()
        self._log(f"Fecha seleccionada: {date_str}", "info")

    def _construir_selector_reporte(self, parent):
        """Construye el selector de tipo de reporte"""
        frame = ttk.Frame(parent)
        frame.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Label(
            frame,
            text="Tipo de Reporte:",
            font=('Helvetica', 10)
        ).pack(side=tk.LEFT, padx=(0, 10))
        
        reports = self.report_generator.get_available_reports()
        friendly_names = {
            'daily_report': 'Detallado',
            'general_report': 'General',
            'f049_report': 'F-049'
        }
        
        for report in reports:
            ttk.Radiobutton(
                frame,
                text=friendly_names.get(report, report.replace('_', ' ').title()),
                variable=self.report_type,
                value=report,
                bootstyle="info-toolbutton"
            ).pack(side=tk.LEFT, padx=5)
        
        if reports:
            self.report_type.set(reports[0])

    def _construir_botones_accion(self, parent):
        """Construye los botones de acción principales"""
        frame = ttk.Frame(parent)
        frame.pack(fill=tk.X)
        
        self.generate_btn = ttk.Button(
            frame,
            text="Generar Reporte",
            command=self._generar_reporte,
            bootstyle=SUCCESS,
            width=20
        )
        self.generate_btn.pack(side=tk.LEFT, padx=5, pady=10)
        
        self.send_btn = ttk.Button(
            frame,
            text="Enviar por Email",
            command=self._enviar_reporte,
            bootstyle=INFO,
            width=20
        )
        self.send_btn.pack(side=tk.LEFT, padx=5, pady=10)
        
        self.open_btn = ttk.Button(
            frame,
            text="Abrir Reporte",
            command=self._abrir_reporte,
            bootstyle=WARNING,
            width=20
        )
        self.open_btn.pack(side=tk.LEFT, padx=5, pady=10)

    def _construir_consola(self, parent):
        """Construye la consola de registro"""
        frame = ttk.Labelframe(
            parent,
            text="Registro de Actividades",
            padding=(15, 10),
            bootstyle=INFO
        )
        frame.pack(fill=tk.BOTH, expand=True)
        
        self.console = tk.Text(
            frame,
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
        
        scrollbar = ttk.Scrollbar(
            frame,
            orient=tk.VERTICAL,
            command=self.console.yview,
            bootstyle=ROUND
        )
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.console.config(yscrollcommand=scrollbar.set)
        
        self.console.tag_config("info", foreground="#007bff")
        self.console.tag_config("success", foreground="#28a745")
        self.console.tag_config("warning", foreground="#fd7e14")
        self.console.tag_config("error", foreground="#dc3545")

    def _construir_barra_estado(self, parent):
        """Construye la barra de estado"""
        self.status_var = tk.StringVar(value="Listo")
        ttk.Label(
            parent,
            textvariable=self.status_var,
            relief=tk.SUNKEN,
            anchor=tk.W,
            font=('Helvetica', 9),
            bootstyle=(SECONDARY, INVERSE)
        ).pack(fill=tk.X, pady=(10, 0))

    def _center_window(self):
        """Centra la ventana en la pantalla"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')

    def _log(self, message, level="info"):
        """Escribe mensajes en la consola con formato"""
        self.console.config(state=tk.NORMAL)
        self.console.insert(tk.END, message + "\n", level)
        self.console.see(tk.END)
        self.console.config(state=tk.DISABLED)
        self.status_var.set(message)
        self.root.update()

    def _deshabilitar_interfaz(self):
        """Deshabilita los controles durante operaciones"""
        for widget in [self.generate_btn, self.send_btn, self.open_btn, self.date_entry]:
            widget.config(state=tk.DISABLED)
        self.root.config(cursor="watch")
        self.root.update()

    def _habilitar_interfaz(self):
        """Habilita los controles nuevamente"""
        for widget in [self.generate_btn, self.send_btn, self.open_btn, self.date_entry]:
            widget.config(state=tk.NORMAL)
        self.root.config(cursor="")
        self.root.update()

    def _generar_reporte(self):
        """Genera el reporte seleccionado"""
        try:
            self._deshabilitar_interfaz()
            fecha_str = self.fecha_var.get()
            
            try:
                fecha = datetime.strptime(fecha_str, '%d/%m/%Y').date()
            except ValueError:
                try:
                    fecha = datetime.strptime(fecha_str, '%d-%m-%Y').date()
                    self.fecha_var.set(fecha.strftime('%d/%m/%Y'))
                except ValueError as e:
                    raise ValueError("Formato de fecha no válido. Use DD/MM/AAAA o DD-MM-AAAA")
            
            report_type = self.report_type.get()
            
            self._log(f"Buscando registros para: {fecha_str} ({fecha.isoformat()})", "info")
            
            df = self.sharepoint_client.get_items_for_report(report_type, fecha)
            
            if df.empty:
                self._log("No hay datos para la fecha seleccionada", "warning")
                messagebox.showwarning("Sin datos", "No se encontraron registros", parent=self.root)
                return
            
            self._log(f"✓ {len(df)} registros encontrados", "success")
            
            output = self.report_generator.generate(
                report_type, 
                df, 
                report_date=fecha
            )
            
            if report_type == 'f049_report':
                self.current_report = output
                self._log(f"PDF generado: {output}", "success")
                messagebox.showinfo("Éxito", f"Informe F-049 generado:\n{output}", parent=self.root)
            else:
                self.current_report = output
                self._log("Reporte generado", "success")
                
        except ValueError as e:
            self._log(f"✗ Error en formato de fecha: {str(e)}", "error")
            messagebox.showerror("Error", str(e), parent=self.root)
        except Exception as e:
            self._log(f"✗ Error inesperado: {str(e)}", "error")
            messagebox.showerror("Error", f"Ocurrió un error inesperado: {str(e)}", parent=self.root)
        finally:
            self._habilitar_interfaz()
            self._log("Proceso completado", "info")
            self.status_var.set("Listo")

    def _enviar_reporte(self):
        """Envía el reporte por email"""
        if not hasattr(self, 'current_report'):
            messagebox.showwarning("Advertencia", "Primero genere un reporte", parent=self.root)
            return
            
        try:
            self._deshabilitar_interfaz()
            fecha_str = self.fecha_var.get()
            report_type = self.report_type.get()
            
            if report_type == 'f049_report':
                success, msg = self.email_sender.send_with_attachment(
                    f"Informe F-049 - {fecha_str}",
                    f"Se adjunta el informe F-049 para {fecha_str}",
                    self.current_report
                )
            else:
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
        finally:
            self._habilitar_interfaz()

    def _abrir_reporte(self):
        """Abre el reporte generado"""
        if not hasattr(self, 'current_report'):
            messagebox.showwarning("Advertencia", "Primero genere un reporte", parent=self.root)
            return
            
        try:
            if isinstance(self.current_report, str) and self.current_report.endswith('.pdf'):
                if os.name == 'nt':  # Windows
                    os.startfile(self.current_report)
                elif os.name == 'posix':  # Mac/Linux
                    subprocess.run(['open' if sys.platform == 'darwin' else 'xdg-open', self.current_report])
            else:
                import webbrowser
                webbrowser.open_new_tab(self.current_report if isinstance(self.current_report, str) else "data:text/html," + self.current_report)
                
        except Exception as e:
            self._log(f"Error al abrir reporte: {str(e)}", "error")
            messagebox.showerror("Error", f"No se pudo abrir el reporte: {str(e)}", parent=self.root)