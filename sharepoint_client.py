from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
from datetime import datetime
import pandas as pd
from config import SHAREPOINT_SITE, LIST_NAME, USERNAME, PASSWORD

class SharePointClient:
    def __init__(self):
        self.ctx = ClientContext(SHAREPOINT_SITE).with_credentials(
            UserCredential(USERNAME, PASSWORD))
        self.sp_list = self.ctx.web.lists.get_by_title(LIST_NAME)
        self._cache = {}  # Cache para optimizar consultas

    def get_items_by_date(self, target_date, fields=None):
        """Obtiene items filtrados por fecha con campos específicos"""
        cache_key = f"items_{target_date}_{'_'.join(fields) if fields else 'all'}"
        
        if cache_key in self._cache:
            return self._cache[cache_key]
            
        try:
            # Selección eficiente de campos
            if fields:
                select_fields = ["ID", "Created"] + fields
                items = self.sp_list.get_items().select(select_fields).execute_query()
            else:
                items = self.sp_list.get_items().select(["*"]).execute_query()
            
            filtered_items = [
                item for item in items 
                if self._is_item_from_date(item, target_date)
            ]
            
            result = self._process_items(filtered_items, fields)
            self._cache[cache_key] = result
            return result
            
        except Exception as e:
            raise Exception(f"Error al obtener datos: {str(e)}")

    def get_items_for_report(self, report_name, target_date):
        """Obtiene datos optimizados para un reporte específico"""
        # Mapeo de campos necesarios por reporte
        report_fields = {
            'daily': ["PROYECTO", "RUTA", "DESCRIPCION", "PLACA", "MATERIAL", "VOLUMEN", "OBSERVACION"],
            'general': ["PROYECTO", "RUTA", "VOLUMEN"],
            'f049': ["PROVEEDOR", "REMISION", "LOTE", "DESCRIPCION", "UNIDAD", "VOLUMEN", "INSPECCION_VISUAL", "OBSERVACIONES"]
        }
        
        return self.get_items_by_date(target_date, report_fields.get(report_name))

    def _is_item_from_date(self, item, target_date):
        """Filtra items por fecha de creación"""
        created_date = item.properties.get('Created')
        if not created_date:
            return False
        try:
            item_date = datetime.strptime(created_date.split('T')[0], '%Y-%m-%d').date()
            return item_date == target_date
        except ValueError:
            return False

    def _process_items(self, items, specific_fields=None):
        """Convierte items a DataFrame optimizado"""
        processed_data = []
        for item in items:
            registro = {"ID": item.properties.get("ID")}
            
            all_fields = {
                "PROYECTO": "PROYECTO",
                "RUTA": "RUTA",
                "DESCRIPCION": "DESCRIPCION",
                "PLACA": "PLACA",
                "MATERIAL": "MATERIAL",
                "VOLUMEN": "VOLUMEN",
                "OBSERVACION": "OBSERVACION",
                "PROVEEDOR": "PROVEEDOR",
                "REMISION": "REMISION",
                "LOTE": "LOTE",
                "UNIDAD": "UNIDAD",
                "INSPECCION_VISUAL": "INSPECCION_VISUAL",
                "OBSERVACIONES": "OBSERVACIONES"
            }
            
            # Filtrar campos si se especificó
            fields_to_get = specific_fields if specific_fields else all_fields.keys()
            
            for field in fields_to_get:
                registro[field] = self._get_field_value(item, all_fields[field])
            
            processed_data.append(registro)
        
        return pd.DataFrame(processed_data)

    def _get_field_value(self, item, field_name):
        """Obtiene valores de campos con manejo seguro"""
        value = item.properties.get(field_name)
        if isinstance(value, dict):
            return value.get('Title', value.get('Value', ""))
        elif isinstance(value, list):
            return ", ".join([str(v) for v in value])
        return str(value) if value else ""

    def clear_cache(self):
        """Limpia la caché de consultas"""
        self._cache = {}