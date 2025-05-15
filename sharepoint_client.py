from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
from datetime import datetime, timedelta
import pandas as pd
from dateutil.parser import parse
from pytz import timezone
import pytz
from config import SHAREPOINT_SITE, LIST_NAME, USERNAME, PASSWORD

class SharePointClient:
    def __init__(self):
        self.ctx = ClientContext(SHAREPOINT_SITE).with_credentials(
            UserCredential(USERNAME, PASSWORD))
        self.sp_list = self.ctx.web.lists.get_by_title(LIST_NAME)
        self._cache = {}  # Cache para optimizar consultas

    def get_items_by_date(self, target_date, fields=None):
        """Obtiene items filtrados por fecha con logging detallado"""
        cache_key = f"items_{target_date}_{'_'.join(fields) if fields else 'all'}"
        
        if cache_key in self._cache:
            return self._cache[cache_key]
            
        try:
            print(f"\n[DEBUG] Buscando registros para fecha: {target_date}")
            
            # Campos base que siempre necesitamos para el filtrado por fecha
            base_fields = ["ID", "Created", "Modified", "Fecha", "Date", "FechaRegistro"]
            select_fields = base_fields + (fields if fields else [])
            
            items = self.sp_list.get_items().select(select_fields).execute_query()
            print(f"[DEBUG] Total items recuperados: {len(items)}")
            
            filtered_items = []
            for item in items:
                if self._is_item_from_date(item, target_date):
                    print(f"[DEBUG] Item {item.properties.get('ID')} con fecha válida")
                    filtered_items.append(item)
            
            print(f"[DEBUG] Items filtrados: {len(filtered_items)}")
            
            result = self._process_items(filtered_items, fields)
            self._cache[cache_key] = result
            return result
            
        except Exception as e:
            print(f"[ERROR] Al obtener datos: {str(e)}")
            raise Exception(f"Error al obtener datos: {str(e)}")

    def get_items_for_report(self, report_name, target_date):
        """Obtiene datos optimizados para un reporte específico"""
        report_fields = {
            'daily_report': ["PROYECTO", "RUTA", "DESCRIPCION", "PLACA", "MATERIAL", "VOLUMEN", "OBSERVACION"],
            'general_report': ["PROYECTO", "RUTA", "VOLUMEN"],
            'f049_report': ["PROVEEDOR", "REMISION", "LOTE", "DESCRIPCION", "UNIDAD", "VOLUMEN", "INSPECCION_VISUAL", "OBSERVACIONES"]
        }
        
        return self.get_items_by_date(target_date, report_fields.get(report_name))

    def get_all_items(self):
        """Obtiene todos los items sin filtrar (para diagnóstico)"""
        try:
            items = self.sp_list.get_items().select(["*"]).execute_query()
            return items
        except Exception as e:
            print(f"[ERROR] Al obtener todos los items: {str(e)}")
            raise

    def _is_item_from_date(self, item, target_date):
        """Filtra items por fecha con manejo robusto de zonas horarias"""
        # Lista de posibles campos de fecha a verificar
        date_fields = [
            'Created',        # Fecha creación estándar
            'Modified',       # Fecha modificación
            'Fecha',          # Posible campo personalizado
            'Date',           # Posible campo en inglés
            'FechaRegistro',  # Otro posible campo
            'FechaCreacion'   # Otro posible nombre
        ]
        
        # Buscar el primer campo de fecha que exista
        for field in date_fields:
            date_value = item.properties.get(field)
            if date_value:
                try:
                    # Parsear la fecha considerando zona horaria
                    item_datetime = parse(date_value)
                    
                    # Si no tiene zona horaria, asumir UTC (valor por defecto SharePoint)
                    if item_datetime.tzinfo is None:
                        item_datetime = pytz.utc.localize(item_datetime)
                    
                    # Convertir a zona horaria local (Bogotá)
                    local_tz = timezone('America/Bogota')
                    local_date = item_datetime.astimezone(local_tz).date()
                    
                    print(f"[DEBUG] Campo {field}: {date_value} -> Fecha local: {local_date}")
                    
                    return local_date == target_date
                    
                except Exception as e:
                    print(f"[WARNING] Error procesando campo {field} ({date_value}): {str(e)}")
                    continue
        
        print(f"[WARNING] No se pudo determinar fecha para item ID: {item.properties.get('ID')}")
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

    def diagnosticar_fechas(self):
        """Método de diagnóstico para verificar fechas"""
        try:
            items = self.get_all_items()
            print("\n[DIAGNÓSTICO] Análisis de fechas en SharePoint:")
            
            date_fields_found = set()
            
            # Analizar los primeros 20 items
            for item in items[:20]:
                print(f"\nItem ID: {item.properties.get('ID')}")
                
                # Mostrar todos los campos que contengan 'date' o 'fecha'
                for field, value in item.properties.items():
                    if 'date' in field.lower() or 'fecha' in field.lower():
                        date_fields_found.add(field)
                        print(f"  {field}: {value}")
            
            print("\nCampos de fecha encontrados:", date_fields_found)
            
        except Exception as e:
            print(f"Error en diagnóstico: {str(e)}")