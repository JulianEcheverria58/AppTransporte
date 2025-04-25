from flask import Flask, request
from sharepoint_client import SharePointClient
from datetime import datetime
import os
from config import SHAREPOINT_SITE, LIST_NAME  # Asegúrate de importar las configuraciones

app = Flask(__name__)

@app.route('/approve', methods=['GET'])
def approve_items():
    try:
        # 1. Validar parámetros
        date_str = request.args.get('date')
        if not date_str:
            return "Falta parámetro 'date'", 400
        
        # 2. Conectar a SharePoint
        sp_client = SharePointClient()  # Asegúrate que usa las credenciales correctas
        
        # 3. Obtener items por fecha
        target_date = datetime.strptime(date_str, '%Y-%m-%d').date()
        df = sp_client.get_items_by_date(target_date)
        
        if df.empty:
            return "No hay registros para esta fecha", 404
        
        # 4. Aprobar items (DEBUG: imprimir IDs primero)
        print(f"Intentando aprobar {len(df)} registros...")
        print("IDs de registros:", df['ID'].tolist())
        
        success, msg = sp_client.approve_items(df['ID'].tolist())
        
        # 5. Verificar resultado
        if not success:
            print(f"Error en aprobación: {msg}")
            return f"Error al aprobar: {msg}", 500
        
        print("Aprobación exitosa!")
        return f"""
        <h1 style='color: green;'>✓ Aprobación Exitosa</h1>
        <p>{len(df)} registros actualizados en SharePoint.</p>
        <p>Campo <strong>RESIDENTE</strong> cambiado a 'Si'.</p>
        """
        
    except Exception as e:
        print(f"Error crítico: {str(e)}")
        return f"Error inesperado: {str(e)}", 500

def run_service():
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=True)  # Modo debug activado

if __name__ == '__main__':
    run_service()