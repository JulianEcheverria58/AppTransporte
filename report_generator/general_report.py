from datetime import datetime
import pandas as pd

class GeneralReport:
    def generate(self, df, **kwargs):
        """Genera reporte general consolidado"""
        report_date = kwargs.get('report_date', datetime.now().date())
        incluir_aprobacion = kwargs.get('incluir_aprobacion', False)
        
        # Determinar si el dataframe ya está consolidado
        is_consolidated = 'Ruta' in df.columns and 'Volumen Total (m³)' in df.columns
        
        # Encabezado con logo MHC
        header = """
        <div style="display: flex; align-items: center; margin-bottom: 20px;">
            <img src="https://www.mhc.com.co/wp-content/uploads/2021/07/logo-mhc.png" 
                 alt="Logo MHC" style="height: 60px; margin-right: 20px;">
            <div>
                <h1 style="color: #E31937; margin: 0;">Reporte General de Transporte</h1>
                <p style="margin: 5px 0 0; color: #555;">{fecha}</p>
            </div>
        </div>
        """.format(fecha=report_date.strftime('%d/%m/%Y'))

        # Tabla consolidada
        table_html = self._generate_consolidated_table(df, is_consolidated)
        
        # Footer
        footer = """
        <div style="margin-top: 30px; font-size: 0.9em; color: #666; border-top: 1px solid #eee; padding-top: 10px;">
            <p>Reporte generado automáticamente el {fecha_hora}</p>
            <p>© {año} MHC - Todos los derechos reservados</p>
        </div>
        """.format(
            fecha_hora=datetime.now().strftime('%d/%m/%Y %H:%M'),
            año=datetime.now().year
        )
        
        # Sección de aprobación
        approval_section = ""
        if incluir_aprobacion:
            approval_section = """
            <div style="margin-top: 40px; padding-top: 20px; border-top: 1px dashed #ccc;">
                <h3 style="color: #E31937;">Aprobación</h3>
                <p>Por favor confirme la recepción y aprobación de este reporte:</p>
                
                <div style="margin-top: 30px;">
                    <p>_________________________________________</p>
                    <p>Nombre y Firma</p>
                </div>
                
                <div style="margin-top: 20px;">
                    <p>Fecha: ____/____/______</p>
                    <p>Hora: ______</p>
                </div>
            </div>
            """
        
        # HTML completo
        html = f"""
        <html>
            <head>
                <style>
                    body {{ 
                        font-family: Arial, sans-serif; 
                        margin: 40px; 
                        color: #333;
                        line-height: 1.6;
                    }}
                    h2 {{ 
                        color: #E31937; 
                        border-bottom: 2px solid #E31937;
                        padding-bottom: 5px;
                        margin-top: 30px;
                    }}
                    table {{
                        border-collapse: collapse; 
                        width: 100%; 
                        margin: 20px 0;
                        box-shadow: 0 2px 3px rgba(0,0,0,0.1);
                    }}
                    th {{ 
                        background-color: #E31937; 
                        color: white; 
                        padding: 12px; 
                        text-align: left;
                    }}
                    td {{
                        padding: 10px; 
                        border-bottom: 1px solid #ddd;
                    }}
                    tr:nth-child(even) {{
                        background-color: #f9f9f9;
                    }}
                    tr:hover {{
                        background-color: #f1f1f1;
                    }}
                    .total-row {{
                        font-weight: bold; 
                        background-color: #f5f5f5;
                    }}
                    .note {{
                        font-size: 0.9em;
                        color: #666;
                        margin-top: 5px;
                    }}
                </style>
            </head>
            <body>
                {header}
                
                <h2>Resumen Consolidado</h2>
                {table_html}
                
                <p class="note">* Volúmenes expresados en metros cúbicos (m³)</p>
                
                {approval_section}
                {footer}
            </body>
        </html>
        """
        return html
    
    def _generate_consolidated_table(self, df, is_consolidated):
        """Genera la tabla HTML consolidada con formato profesional"""
        if is_consolidated:
            # Crear copia para no modificar el original
            df_display = df.copy()
            
            # Formatear números
            if 'Volumen Total (m³)' in df_display.columns:
                df_display['Volumen Total (m³)'] = df_display['Volumen Total (m³)'].apply(lambda x: f"{x:,.2f}")
            
            # Generar HTML
            html = df_display.to_html(index=False, escape=False, classes="consolidated-table")
            
            # Agregar fila de totales si corresponde
            if 'Volumen Total (m³)' in df.columns and 'Cantidad de Viajes' in df.columns:
                total_volumen = df['Volumen Total (m³)'].sum()
                total_viajes = df['Cantidad de Viajes'].sum()
                
                total_row = f"""
                <tr class="total-row">
                    <td><strong>TOTAL</strong></td>
                    <td><strong>{total_viajes}</strong></td>
                    <td><strong>{total_volumen:,.2f}</strong></td>
                </tr>
                """
                
                html = html.replace('</tbody>', total_row + '</tbody>')
            
            return html
        else:
            # Si no está consolidado, mostrar advertencia y tabla original
            warning = """
            <div style="background-color: #fff3cd; color: #856404; padding: 10px; 
                        border-radius: 4px; border: 1px solid #ffeeba; margin-bottom: 20px;">
                <strong>Nota:</strong> Los datos no están consolidados. Mostrando detalles completos.
            </div>
            """
            return warning + df.to_html(index=False)