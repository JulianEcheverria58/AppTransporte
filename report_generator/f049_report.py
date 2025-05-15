from fpdf import FPDF
from datetime import datetime
from PIL import Image
import os

class F049Report:
    def __init__(self):
        self.default_logo_path = r"C:\Users\julian.echeverria\OneDrive - MHC\Escritorio\Python\AppTransporte\logo.png"
        self.font_size = 8
        self.line_height = 4
        self.min_row_height = 8
        self.header_bg_color = (240, 240, 240)
        
    def generate(self, df, output_path=None, logo_path=None, **kwargs):
        """Generate F-049 report PDF with perfect cell formatting"""
        if not output_path:
            output_path = f"F049_Reporte_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
        
        pdf = FPDF(orientation='L', unit='mm', format='A4')
        pdf.add_page()
        pdf.set_margins(15, 10, 15)
        pdf.set_auto_page_break(True, margin=15)
        
        self._draw_header_with_logo(pdf, logo_path)
        self._add_report_content(pdf, df, **kwargs)
        
        pdf.output(output_path)
        print(f"PDF generated successfully: {output_path}")
        return output_path

    def _draw_header_with_logo(self, pdf, logo_path=None):
        """Draw report header with centered logo"""
        header_height = 25
        total_width = 277
        
        # Try to load logo
        final_logo_path = logo_path or self.default_logo_path
        logo_loaded = False
        
        if os.path.exists(final_logo_path):
            try:
                img = Image.open(final_logo_path)
                aspect_ratio = img.width / img.height
                max_logo_h = header_height - 4
                max_logo_w = (total_width * 0.25) - 10
                logo_w = min(max_logo_w, max_logo_h * aspect_ratio)
                logo_h = logo_w / aspect_ratio
                
                pdf.image(
                    final_logo_path,
                    x=15 + ((total_width * 0.25) - logo_w)/2,
                    y=10 + (header_height - logo_h)/2,
                    w=logo_w,
                    h=logo_h
                )
                logo_loaded = True
            except Exception as e:
                print(f"Error loading logo: {str(e)}")
        
        if not logo_loaded:
            pdf.set_draw_color(0, 0, 0)
            pdf.set_fill_color(*self.header_bg_color)
            pdf.rect(15, 10, total_width * 0.25, header_height, 'DF')
            pdf.set_font("Arial", 'I', 8)
            pdf.set_xy(15, 10 + header_height/2 - 3)
            pdf.cell(total_width * 0.25, 6, "LOGO NO ENCONTRADO", 0, 0, 'C')

        # Center section
        pdf.set_font("Arial", 'B', 14)
        pdf.set_xy(15 + (total_width * 0.25) + 5, 10 + 5)
        pdf.cell(total_width * 0.5 - 10, 8, "RECEPCION Y CONTROL DE MATERIALES", 0, 0, 'C')
        pdf.set_xy(15 + (total_width * 0.25) + 5, 10 + 15)
        pdf.cell(total_width * 0.5 - 10, 8, "EN LA VIA Y/O ALMACEN", 0, 0, 'C')

        # Right section
        pdf.set_font("Arial", 'B', 16)
        pdf.set_xy(15 + (total_width * 0.75), 10 + header_height/2 - 8)
        pdf.cell(total_width * 0.25, 16, "F-049", 0, 0, 'C')

    def _add_report_content(self, pdf, df, **kwargs):
        """Generate main report content with perfect cell alignment"""
        start_y = 40
        
        # Project info
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
        
        # Table configuration
        headers = [
            "PROVEEDOR", "REMISIÓN", "IDENTIFICACION LOTE", 
            "DESCRIPCION", "UND", "CERTIFICADO DE CALIDAD",
            "INSPECCION VISUAL", "ACTIVIDAD A E."
        ]
        col_widths = [40, 25, 35, 50, 15, 40, 30, 30]
        
        # Draw headers
        pdf.set_font("Arial", 'B', self.font_size)
        initial_y = pdf.get_y()
        
        # Header background
        pdf.set_fill_color(*self.header_bg_color)
        for i in range(len(headers)):
            pdf.rect(15 + sum(col_widths[:i]), initial_y, col_widths[i], self.min_row_height, 'F')
        
        # Header text
        for i, header in enumerate(headers):
            pdf.set_xy(15 + sum(col_widths[:i]), initial_y)
            pdf.cell(col_widths[i], self.min_row_height, header, 0, 0, 'C')
        
        pdf.set_y(initial_y + self.min_row_height)
        
        # Data rows
        pdf.set_font("Arial", '', self.font_size)
        
        for _, row in df.iterrows():
            current_y = pdf.get_y()
            max_cell_height = self.min_row_height
            
            # Calculate required height for each cell
            cell_data = []
            for i, header in enumerate(headers):
                value = str(row.get(header.split()[0].upper(), ''))
                if header == "UND": value = "M3"
                
                # Calculate text height
                lines = pdf.multi_cell(
                    col_widths[i], 
                    self.line_height,
                    value, 
                    border=0,
                    align='C' if header in ['UND', 'REMISIÓN'] else 'L',
                    split_only=True
                )
                cell_height = max(self.min_row_height, len(lines) * self.line_height)
                max_cell_height = max(max_cell_height, cell_height)
                cell_data.append((value, cell_height))
            
            # Draw cells with uniform height
            for i, (value, _) in enumerate(cell_data):
                pdf.set_xy(15 + sum(col_widths[:i]), current_y)
                pdf.multi_cell(
                    col_widths[i], 
                    max_cell_height,
                    value, 
                    border=1,
                    align='C' if headers[i] in ['UND', 'REMISIÓN'] else 'L',
                    fill=False
                )
            
            # Move to next row
            pdf.set_xy(15, current_y + max_cell_height)
        
        # Footer
        pdf.ln(10)
        pdf.set_font("Arial", 'B', 10)
        pdf.cell(60, 8, "RESPONSABLE:", 0, 0)
        pdf.cell(60, 8, "CARGO:", 0, 0)
        pdf.cell(60, 8, "FIRMA:", 0, 1)
        
        pdf.set_font("Arial", 'I', 8)
        pdf.cell(0, 8, f"Generado el: {datetime.now().strftime('%d/%m/%Y %H:%M')}", 0, 0, 'R')


# Example usage
if __name__ == "__main__":
    import pandas as pd
    
    # Sample data matching your image
    data = {
        'PROVEEDOR': ['TRANSPORTES Y MOVIMIENTOS CIVILES ACC.S.A.S'] * 6,
        'REMISION': ['121762', '121803', '121763', '121764', '121811', '121811'],
        'LOTE': ['2812', '2816', '2816', '2817', '2841', '2902'],
        'DESCRIPCION': ['Relleno Adecuado (Terraglen)'] * 6,
        'CERTIFICADO_CALIDAD': [''] * 6,
        'INSPECCION_VISUAL': ['17', '16,38', '16,82', '16,69', '18,03', '18,03'],
        'ACTIVIDAD_EJECUTAR': [''] * 6
    }
    df = pd.DataFrame(data)
    
    report = F049Report()
    pdf_path = report.generate(
        df,
        codigo_obra="MAVA-2023",
        output_path="F049_Report_Final.pdf"
    )