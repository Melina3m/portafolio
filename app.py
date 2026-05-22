from flask import Flask, render_template, request, jsonify
import logging

import pathlib

# Configurar el logger
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Inicializar Flask (usa las carpetas por defecto: templates y static)
app = Flask(__name__)

PROYECTOS_METADATA = {
    "buscarean": {
        "display_name": "Servicio de Extracción y Consulta de EAN/SKU (UnoEE)",
        "description": "Microservicio web con Flask y SQLAlchemy que automatiza la extracción de datos de códigos de barra (EAN) y SKUs de la base de datos Siesa UnoEE. Limpia y concatena datos de inventario para exportar a Excel mediante un clic.",
        "badges": ["Flask", "SQLAlchemy", "pyodbc", "Pandas", "Bootstrap"]
    },
    "conciliacionbancaria": {
        "display_name": "Conciliación Bancaria Automatizada (Efectivo & Tarjetas)",
        "description": "Aplicación web Flask y motor de procesamiento analítico que automatiza la conciliación contable de cajas y transacciones con tarjetas. Integra un resolvedor combinatorio en Python que detecta correspondencias múltiples y genera registros listos para Siesa ERP.",
        "badges": ["Flask", "SQLAlchemy", "Pandas", "openpyxl", "xlsxwriter", "Combinatoria"]
    },
    "contabilidad": {
        "display_name": "Comparador de Cortes Contables & Variaciones",
        "description": "Dashboard interactivo con Streamlit conectado a SQL Server (Siesa ERP). Permite comparar cortes financieros de diferentes periodos, agregando saldos por cuenta auxiliar y tercero para analizar variaciones porcentuales y absolutas en cuentas de resultado.",
        "badges": ["Streamlit", "SQLAlchemy", "pyodbc", "Pandas", "Numpy", "Análisis Financiero"]
    },
    "controlg": {
        "display_name": "Gestor de Gastos y Finanzas Personales (Web/Desktop)",
        "description": "Aplicación híbrida que incluye un dashboard web en Flask con PostgreSQL y una interfaz de escritorio en Tkinter. Cuenta con control de presupuestos quincenales, categorización inteligente, alertas de vencimiento y un módulo de seguimiento de deudas insolutas.",
        "badges": ["Flask", "PostgreSQL", "Tkinter", "JSON", "Autenticación"]
    },
    "interfazgconocimiento": {
        "display_name": "Procesador de Indicadores UPT/TKT & ETL de Ventas",
        "description": "Plataforma de inteligencia de negocios con Streamlit y SQL Server para el cálculo automatizado de KPIs comerciales (Unidades por Transacción - UPT y Ticket Promedio - TKT). Realiza transformaciones de datos complejas (ETL) y genera rankings por vendedor a nivel regional.",
        "badges": ["Streamlit", "SQLAlchemy", "pyodbc", "Pandas", "Numpy", "ETL / BI"]
    },
    "magnus": {
        "display_name": "Sistema de Gestión Empresarial & ERP Magnus Parfum",
        "description": "Plataforma ERP completa y panel de administración en Streamlit integrada con Supabase (Base de datos relacional en la nube). Permite gestionar inventarios en tiempo real, flujos de compras/ventas, cuentas por cobrar/pagar, abonos y control de aportes/retornos para inversionistas.",
        "badges": ["Streamlit", "Supabase", "st_supabase_connection", "Pandas", "Gestión ERP"]
    },
    "presupuestobtb": {
        "display_name": "Consolidador y Asignador de Presupuestos (B2B / B2C)",
        "description": "Plataforma web corporativa en Flask conectada a SQL Server para gestionar presupuestos de ventas multicanal (B2B, B2C, Online, Internacional). Incluye endpoints de persistencia de zonas de vendedores y un generador automatizado de libros contables complejos en Excel.",
        "badges": ["Flask", "SQL Server", "Pandas", "openpyxl", "Gestión de Ventas"]
    },
    "pronosticador": {
        "display_name": "Polla Mundialista & Sincronizador de Marcadores en Vivo",
        "description": "Plataforma web modular en Flask para la gestión de pronósticos deportivos y rankings. Integra APScheduler para tareas en segundo plano que sincronizan marcadores de fútbol en tiempo real mediante TheSportsDB API y calcula puntajes acumulados.",
        "badges": ["Flask", "SQLite", "APScheduler", "REST API", "Autenticación", "Background Tasks"]
    },
    "reabastecimiento_online": {
        "display_name": "Optimizador de Reabastecimiento Online & Redistribución",
        "description": "Sistema web con Flask y algoritmos de optimización logística para la redistribución y reabastecimiento de inventario de la Tienda Online. Integra cálculo de mercancía en tránsito, límites de stock por bodega y generación automatizada de órdenes de traslado.",
        "badges": ["Flask", "Pandas", "Logística / SCM", "Excel"]
    }
}

def obtener_info_proyecto(folder_name):
    name_lower = folder_name.lower()
    meta = PROYECTOS_METADATA.get(name_lower)
    
    if meta:
        display_name = meta["display_name"]
        description = meta["description"]
        badges = meta["badges"]
    else:
        # Valores por defecto de respaldo si se añade una nueva carpeta en el futuro
        display_name = folder_name.replace('_', ' ').replace('-', ' ').title()
        description = f"Desarrollo e infraestructura de backend alojado en la carpeta '{folder_name}'."
        badges = ["Python", "Backend"]
        
    return {
        "folder_name": folder_name,
        "display_name": display_name,
        "description": description,
        "badges": badges
    }

@app.route('/')
def home():
    """Renderiza el index.html del portafolio, leyendo dinámicamente la carpeta PROYECTOS"""
    proyectos_dir = pathlib.Path("PROYECTOS")
    lista_proyectos = []
    
    if proyectos_dir.exists() and proyectos_dir.is_dir():
        # Escanear subcarpetas evitando ocultas o entornos virtuales
        subdirs = sorted([
            d.name for d in proyectos_dir.iterdir() 
            if d.is_dir() and not d.name.startswith('.') and d.name.lower() != 'venv'
        ])
        for s in subdirs:
            lista_proyectos.append(obtener_info_proyecto(s))
            
    return render_template('index.html', proyectos=lista_proyectos)

@app.route('/contacto', methods=['POST'])
@app.route('/api/contacto', methods=['POST'])
def contacto():
    """
    Ruta POST que procesa los datos del formulario de contacto.
    Simula la inserción y persistencia en una base de datos SQL Server.
    """
    try:
        # Obtener los datos como JSON o de formulario clásico
        if request.is_json:
            datos = request.get_json()
        else:
            datos = request.form

        nombre = datos.get('nombre')
        email = datos.get('email')
        mensaje = datos.get('mensaje')

        # Validación simple de presencia de datos
        if not nombre or not email or not mensaje:
            return jsonify({
                "status": "error",
                "message": "Faltan campos obligatorios (nombre, email, mensaje)."
            }), 400

        # Log simulando inserción en SQL Server
        logging.info("==================================================")
        logging.info("SIMULANDO CONEXIÓN A SQL SERVER...")
        logging.info("INSERT INTO Contactos (Nombre, Email, Mensaje, Fecha) VALUES (%s, %s, %s, GETDATE())", nombre, email, mensaje)
        logging.info("¡Datos guardados con éxito en la base de datos empresarial!")
        logging.info("==================================================")

        return jsonify({
            "status": "success",
            "message": "Mensaje recibido correctamente. Nos pondremos en contacto pronto."
        }), 200

    except Exception as e:
        logging.error("Error al procesar el mensaje: %s", str(e))
        return jsonify({
            "status": "error",
            "message": "Ocurrió un error interno en el servidor."
        }), 500

if __name__ == '__main__':
    # Ejecuta el servidor de desarrollo
    app.run(debug=True, port=5000)
