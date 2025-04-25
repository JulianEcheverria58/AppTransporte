from importlib import import_module
from pathlib import Path

class ReportGenerator:
    def __init__(self):
        self.report_types = self._discover_reports()
    
    def _discover_reports(self):
        report_types = {}
        reports_dir = Path(__file__).parent
        
        for report_file in reports_dir.glob('*.py'):
            if report_file.stem not in ['__init__', 'base_report']:
                module_name = f"report_generator.{report_file.stem}"
                try:
                    module = import_module(module_name)
                    
                    # Obtener la clase por su nombre (ej: DailyReport, F049Report)
                    class_name = report_file.stem.replace('_report', '').title() + 'Report'
                    if hasattr(module, class_name):
                        report_class = getattr(module, class_name)
                        report_types[report_file.stem] = report_class
                except Exception as e:
                    print(f"Error cargando {module_name}: {str(e)}")
                    continue
        
        return report_types
    
    def generate(self, report_name, df, **kwargs):
        if report_name not in self.report_types:
            raise ValueError(f"Tipo de reporte no válido: {report_name}")
        return self.report_types[report_name]().generate(df, **kwargs)
    
    def get_available_reports(self):
        return list(self.report_types.keys())