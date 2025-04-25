from fpdf import FPDF
from datetime import datetime
from PIL import Image
import os
import base64
from io import BytesIO

class F049Report:
    def __init__(self):
        # Definir la ruta exacta del logo
        self.default_logo_path = r"C:\Users\julian.echeverria\OneDrive - MHC\Escritorio\Python\AppTransporte\logo.png"

    def generate(self, df, output_path=None, logo_path=None, **kwargs):
        """Genera el informe F-049 con la ruta específica del logo"""
        if output_path is None:
            output_path = f"F049_Reporte_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
        
        pdf = FPDF(orientation='L', unit='mm', format='A4')
        pdf.add_page()
        
        # Configurar márgenes
        pdf.set_margins(15, 10, 15)
        pdf.set_auto_page_break(True, margin=15)
        
        # --- Cabecera con logo garantizado ---
        self._draw_header_with_logo(pdf, logo_path)
        
        # --- Contenido principal ---
        self._add_report_content(pdf, df, **kwargs)
        
        # Generar PDF
        pdf.output(output_path)
        print(f"PDF generado correctamente en: {output_path}")
        return output_path

    def _draw_header_with_logo(self, pdf, logo_path=None):
        """Dibuja la cabecera con el logo en la ruta especificada"""
        header_height = 25
        total_width = 277  # Ancho A4 horizontal menos márgenes
        
        # Usar la ruta proporcionada o la ruta por defecto
        final_logo_path = logo_path if logo_path else self.default_logo_path
        
        # 1. Sección del logo (25%)
        logo_section = total_width * 0.25
        pdf.set_xy(15, 10)
        
        # Intentar cargar el logo
        logo_loaded = False
        if os.path.exists(final_logo_path):
            try:
                img = Image.open(final_logo_path)
                img_w, img_h = img.size
                aspect_ratio = img_w / img_h
                
                max_logo_h = header_height - 4
                max_logo_w = logo_section - 10
                
                logo_w = min(max_logo_w, max_logo_h * aspect_ratio)
                logo_h = logo_w / aspect_ratio
                
                logo_x = 15 + (logo_section - logo_w) / 2
                logo_y = 10 + (header_height - logo_h) / 2
                
                pdf.image(final_logo_path, x=logo_x, y=logo_y, w=logo_w, h=logo_h)
                logo_loaded = True
                print(f"Logo cargado desde: {final_logo_path}")
            except Exception as e:
                print(f"Error al cargar el logo: {str(e)}")
        
        if not logo_loaded:
            # Mostrar área de logo con mensaje
            pdf.set_draw_color(0, 0, 0)
            pdf.set_fill_color(240, 240, 240)
            pdf.rect(15, 10, logo_section, header_height, 'DF')
            pdf.set_font("Arial", 'I', 8)
            pdf.set_xy(15, 10 + header_height/2 - 3)
            pdf.cell(logo_section, 6, "LOGO NO ENCONTRADO", 0, 0, 'C')
            print(f"No se pudo cargar el logo desde: {final_logo_path}")

        # 2. Sección central (50%)
        center_section = total_width * 0.5
        pdf.set_font("Arial", 'B', 14)
        pdf.set_xy(15 + logo_section + 5, 10 + 5)
        pdf.cell(center_section - 10, 8, "RECEPCION Y CONTROL DE MATERIALES", 0, 0, 'C')
        pdf.set_xy(15 + logo_section + 5, 10 + 15)
        pdf.cell(center_section - 10, 8, "EN LA VIA Y/O ALMACEN", 0, 0, 'C')

        # 3. Sección derecha (25%)
        right_section = total_width * 0.25
        pdf.set_font("Arial", 'B', 16)
        pdf.set_xy(15 + logo_section + center_section, 10 + header_height/2 - 8)
        pdf.cell(right_section, 16, "F-049", 0, 0, 'C')

    def _add_report_content(self, pdf, df, **kwargs):
        """Añade el contenido principal del reporte"""
        start_y = 40
        
        # Información de obra y fecha
        pdf.set_font("Arial", 'B', 10)
        pdf.set_xy(15, start_y)
        pdf.cell(35, 8, "CODIGO DE OBRA:", 0, 0)
        pdf.set_font("Arial", '', 10)
        pdf.cell(40, 8, kwargs.get('codigo_obra', 'MAVA'), 0, 0)
        
        pdf.set_font("Arial", 'B', 10)
        pdf.cell(15, 8, "FECHA:", 0, 0)
        pdf.set_font("Arial", '', 10)
        pdf.cell(40, 8, kwargs.get('fecha', datetime.now().strftime('%d/%m/%Y %H:%M')), 0, 1)
        pdf.ln(8)
        
        # Tabla principal
        headers = [
            "PROVEEDOR", "REMISIÓN", "IDENTIFICACION LOTE", 
            "DESCRIPCION", "UND", "CERTIFICADO DE CALIDAD",
            "INSPECCION VISUAL", "ACTIVIDAD A E."
        ]
        col_widths = [35, 25, 35, 50, 15, 40, 30, 30]
        
        # Encabezados
        pdf.set_font("Arial", 'B', 8)
        for i, header in enumerate(headers):
            pdf.cell(col_widths[i], 8, header, 1, 0, 'C')
        pdf.ln()
        
        # Datos
        pdf.set_font("Arial", '', 8)
        for _, row in df.iterrows():
            for i, header in enumerate(headers):
                value = str(row.get(header.split()[0].upper(), ''))
                if header == "UND":
                    value = "M3"
                pdf.cell(col_widths[i], 8, value, 1, 0, 'C' if header in ['UND', 'REMISIÓN'] else 'L')
            pdf.ln()
        
        # Pie de página
        pdf.ln(10)
        pdf.set_font("Arial", 'B', 10)
        pdf.cell(60, 8, "RESPONSABLE:", 0, 0)
        pdf.cell(60, 8, "CARGO:", 0, 0)
        pdf.cell(60, 8, "FIRMA:", 0, 1)
        
        pdf.set_font("Arial", 'I', 8)
        pdf.cell(0, 8, f"Generado el: {datetime.now().strftime('%d/%m/%Y %H:%M')}", 0, 0, 'R')


# Ejemplo de uso
if __name__ == "__main__":
    import pandas as pd
    
    # Datos de ejemplo
    data = {
        'PROVEEDOR': ['MHC', 'Proveedor 2'],
        'REMISION': ['12345', '67890'],
        'LOTE': ['LOTE-001', 'LOTE-002'],
        'DESCRIPCION': ['Material de construcción', 'Tuberías PVC'],
        'CERTIFICADO_CALIDAD': ['SI', 'NO'],
        'INSPECCION_VISUAL': ['Aprobado', 'Rechazado'],
        'ACTIVIDAD_EJECUTAR': ['Almacenar', 'Devolver']
    }
    df = pd.DataFrame(data)
    
    # Generar reporte (usará la ruta por defecto del logo)
    report = F049Report()
    pdf_path = report.generate(
        df,
        codigo_obra="MAVA-2023",
        output_path="Reporte_F049_Final.pdf"
    )