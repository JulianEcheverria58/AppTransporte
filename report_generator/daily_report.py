from datetime import datetime

class DailyReport:
    def generate(self, df, **kwargs):
        """Genera reporte diario detallado"""
        report_date = kwargs.get('report_date', datetime.now().date())
        
        html = f"""
        <html>
            <head>
                <style>
                    body {{ font-family: Arial, sans-serif; margin: 20px; }}
                    h2 {{ color: #E31937; border-bottom: 2px solid #E31937; }}
                    table {{ border-collapse: collapse; width: 100%; }}
                    th {{ background-color: #E31937; color: white; padding: 10px; }}
                    td {{ padding: 8px; border-bottom: 1px solid #ddd; }}
                </style>
            </head>
            <body>
                <h2>Reporte Diario Detallado - {report_date.strftime('%d/%m/%Y')}</h2>
                {df.to_html(index=False)}
                <p>Reporte generado el {datetime.now().strftime('%d/%m/%Y %H:%M')}</p>
            </body>
        </html>
        """
        return html