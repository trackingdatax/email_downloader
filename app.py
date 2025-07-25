#!/usr/bin/env python3
"""
app.py - Interfaz Streamlit para el Descargador de Emails
Versión mejorada con rango de fechas específico y validación de contraseñas
"""
import streamlit as st
import os
from pathlib import Path
import pandas as pd
from datetime import datetime, timedelta, date
import zipfile
from unicodedata import normalize

# Importar la clase desde functions.py
from functions import EmailImageDownloader


def normalizar_palabra(palabra):
    palabra = palabra.strip().lower()
    palabra = normalize('NFKD', palabra).encode('ascii', 'ignore').decode('utf-8')
    return palabra


def is_real_password(password):
    """Verifica si la contraseña es real o un placeholder"""
    if not password or not password.strip():
        return False
    
    # Placeholders comunes que no son contraseñas reales
    placeholders = [
        "tu_contraseña_aqui",
        "tu_contraseña_de_app",
        "tu_contraseña_de_app_gmail",
        "tu_contraseña_de_app_gmail_consultorio_aqui",
        "tu_contraseña_de_app_hotmail_aqui",
        "your_password_here",
        "password_here",
        "contraseña",
        "password",
        "example",
        "ejemplo"
    ]
    
    password_lower = password.lower().strip()
    
    # Verificar si es un placeholder
    for placeholder in placeholders:
        if placeholder in password_lower:
            return False
    
    # Verificar que tenga al menos 10 caracteres (las contraseñas de app suelen ser largas)
    if len(password.strip()) < 10:
        return False
    
    return True

    
# Cargar variables de entorno
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# Configuración de la página
st.set_page_config(
    page_title="📧 SmartExtract Dr.Lucero",
    page_icon="📧",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS personalizado
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
        padding: 1rem;
        border-radius: 10px;
        color: white;
        text-align: center;
        margin-bottom: 2rem;
    }
    .status-success {
        background-color: #d4edda;
        color: #155724;
        padding: 0.75rem;
        border-radius: 5px;
        border: 1px solid #c3e6cb;
    }
    .status-error {
        background-color: #f8d7da;
        color: #721c24;
        padding: 0.75rem;
        border-radius: 5px;
        border: 1px solid #f5c6cb;
    }
    .debug-info {
        background-color: #fff3cd;
        color: #856404;
        padding: 0.75rem;
        border-radius: 5px;
        border: 1px solid #ffeaa7;
    }
    .date-range-info {
        background-color: #e8f4fd;
        color: #0c5460;
        padding: 0.75rem;
        border-radius: 5px;
        border: 1px solid #bee5eb;
        margin: 0.5rem 0;
    }
</style>
""", unsafe_allow_html=True)

# Header principal
st.markdown("""
<div class="main-header">
    <h1>📧 SmartExtract Dr.Lucero - CON REPORTE</h1>
    <p>Configura y descarga archivos desde tus emails con reporte detallado</p>
</div>
""", unsafe_allow_html=True)

# Cuentas de email predefinidas
CUENTAS_EMAIL = {
    "victorlucero1981@gmail.com": {
        "email": "victorlucero1981@gmail.com",
        "password": os.getenv("GMAIL_VICTOR_PERSONAL", ""),
        "server": "imap.gmail.com",
        "port": 993,
        "use_ssl": True,
        "tipo": "Gmail Personal"
    },
    "victorLucero.consultorio@gmail.com": {
        "email": "victorLucero.consultorio@gmail.com", 
        "password": os.getenv("GMAIL_VICTOR_CONSULTORIO", ""),
        "server": "imap.gmail.com",
        "port": 993,
        "use_ssl": True,
        "tipo": "Gmail Consultorio"
    },
    "victorlucero1981@hotmail.com": {
        "email": "victorlucero1981@hotmail.com",
        "password": os.getenv("HOTMAIL_VICTOR_PERSONAL", ""),
        "server": "outlook.office365.com", 
        "port": 993,
        "use_ssl": True,
        "tipo": "Hotmail Personal"
    }
}

# Sidebar para configuración
with st.sidebar:
    st.header("⚙️ Configuración")
    
    cuenta_seleccionada = st.selectbox(
        "🔧 Cuenta de Email",
        options=list(CUENTAS_EMAIL.keys()),
        help="Selecciona una de tus cuentas de email predefinidas"
    )
    
    config_cuenta = CUENTAS_EMAIL[cuenta_seleccionada]
    
    # Mostrar información de la cuenta
    st.info(f"""
    **Email:** {config_cuenta['email']}  
    **Tipo:** {config_cuenta['tipo']}  
    **Servidor:** {config_cuenta['server']}:{config_cuenta['port']}
    """)
    
    # Configurar valores
    email_usuario = config_cuenta['email']
    password_usuario = config_cuenta['password']
    servidor = config_cuenta['server']
    puerto = config_cuenta['port']
    usar_ssl = config_cuenta['use_ssl']
    
    # Mostrar estado de la contraseña con validación mejorada
    if is_real_password(password_usuario):
        st.success("✅ Contraseña configurada")
    else:
        st.error("❌ Falta configurar contraseña válida")
        
        # Mostrar ayuda específica según el tipo de cuenta
        if "gmail" in email_usuario.lower():
            st.info("""
            📋 **Para Gmail:**
            1. Ve a [myaccount.google.com](https://myaccount.google.com)
            2. Seguridad → Verificación en 2 pasos
            3. Contraseñas de aplicaciones
            4. Genera una nueva para 'Mail'
            5. Actualiza tu archivo .env
            """)
        elif "hotmail" in email_usuario.lower() or "outlook" in email_usuario.lower():
            st.info("""
            📋 **Para Hotmail/Outlook:**
            1. Ve a [account.microsoft.com](https://account.microsoft.com)
            2. Seguridad → Opciones avanzadas de seguridad
            3. Contraseñas de aplicación
            4. Genera una nueva para 'Correo'
            5. Actualiza tu archivo .env
            """)
        

# Tabs principales
tab1, tab2, tab3 = st.tabs(["🎯 Filtros", "📁 Archivos", "📊 Resultados"])

with tab1:
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.subheader("📮 Filtros de Email")
        
        remitentes_text = st.text_area(
            "👥 Emails Remitentes (uno por línea):",
            placeholder="empresa@ejemplo.com\nfacturas@proveedor.com",
            height=100
        )
        remitentes = [email.strip() for email in remitentes_text.split('\n') if email.strip()]
        
        if remitentes:
            st.success(f"✅ {len(remitentes)} remitente(s) configurado(s)")
        
        # === NUEVA SECCIÓN: RANGO DE FECHAS ===
        st.subheader("📅 Rango de Fechas")
        
        # Opción para activar/desactivar filtro de fechas
        usar_filtro_fecha = st.checkbox(
            "🔘 Filtrar por rango de fechas", 
            value=True,
            help="Si está desactivado, buscará en todos los emails (puede ser muy lento)"
        )
        
        if usar_filtro_fecha:
            col_fecha1, col_fecha2 = st.columns(2)
            
            with col_fecha1:
                fecha_inicio = st.date_input(
                    "📅 Fecha de inicio",
                    value=date(2020, 1, 1),
                    min_value=date(2000, 1, 1),
                    max_value=date.today(),
                    help="Fecha desde la cual buscar emails"
                )
            
            with col_fecha2:
                fecha_fin = st.date_input(
                    "📅 Fecha de fin",
                    value=date.today(),
                    min_value=date(2000, 1, 1),
                    max_value=date.today(),
                    help="Fecha hasta la cual buscar emails"
                )
            
            # Validación de fechas
            if fecha_inicio > fecha_fin:
                st.error("❌ **Error:** La fecha de inicio debe ser anterior a la fecha de fin")
                fecha_valida = False
            elif fecha_fin > date.today():
                st.error("❌ **Error:** La fecha de fin no puede ser futura")
                fecha_valida = False
            else:
                fecha_valida = True
                
                # Calcular duración del rango
                duracion = (fecha_fin - fecha_inicio).days + 1
                
        else:
            fecha_valida = True
            st.info("ℹ️ Se buscarán emails en **todo el historial** (puede ser muy lento)")
    
    with col2:
        st.subheader("🔍 Filtros Adicionales")
        
        palabras_clave_text = st.text_area(
            "🔍 Palabras clave en asunto (una por línea):",
            value="tcmaxonline\ntomografia\nradiografia\npanoramica\nrx",
            height=120
        )
        
        palabras_clave_raw = [p.strip() for p in palabras_clave_text.split('\n') if p.strip()]
        palabras_clave = [normalizar_palabra(p) for p in palabras_clave_raw]
        
        carpeta_email = st.selectbox(
            "📂 Carpeta de email",
            options=["INBOX", "SPAM", "SENT", "DRAFTS"],
            index=0
        )
        
        # === OPCIONES ADICIONALES DE FILTRADO ===
        st.subheader("⚙️ Opciones Avanzadas")
        
        # Presets rápidos para fechas comunes
        st.markdown("**🚀 Presets rápidos:**")
        col_preset1, col_preset2 = st.columns(2)
        
        with col_preset1:
            if st.button("📅 Último mes", help="Últimos 30 días"):
                st.session_state.fecha_inicio_preset = date.today() - timedelta(days=30)
                st.session_state.fecha_fin_preset = date.today()
                st.rerun()
            
            if st.button("📅 Últimos 3 meses", help="Últimos 90 días"):
                st.session_state.fecha_inicio_preset = date.today() - timedelta(days=90)
                st.session_state.fecha_fin_preset = date.today()
                st.rerun()
        
        with col_preset2:
            if st.button("📅 Este año", help="Desde 1 de enero"):
                st.session_state.fecha_inicio_preset = date(date.today().year, 1, 1)
                st.session_state.fecha_fin_preset = date.today()
                st.rerun()
            
            if st.button("📅 Año pasado", help="Todo el año anterior"):
                year_last = date.today().year - 1
                st.session_state.fecha_inicio_preset = date(year_last, 1, 1)
                st.session_state.fecha_fin_preset = date(year_last, 12, 31)
                st.rerun()
        
        # Aplicar presets si existen
        if hasattr(st.session_state, 'fecha_inicio_preset'):
            fecha_inicio = st.session_state.fecha_inicio_preset
            fecha_fin = st.session_state.fecha_fin_preset
            del st.session_state.fecha_inicio_preset
            del st.session_state.fecha_fin_preset

with tab2:
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.subheader("📁 Configuración de Carpetas")
        
        carpeta_base = st.text_input(
            "📂 Carpeta base de descarga",
            value="./archivos_medicos_descargados"
        )
        
        st.markdown("**🗂️ Estructura de carpetas:**")
        por_fecha = st.checkbox("📅 Organizar por fecha", value=True)
        por_remitente = st.checkbox("👤 Organizar por remitente", value=True)
        por_asunto = st.checkbox("📝 Organizar por asunto", value=False)
    
    with col2:
        st.subheader("📄 Tipos de Archivo")
        
        col_img, col_doc = st.columns(2)
        
        with col_img:
            st.markdown("*Imágenes:*")
            ext_jpg = st.checkbox("📷 JPG/JPEG", value=True)
            ext_png = st.checkbox("🖼️ PNG", value=True)
            ext_gif = st.checkbox("🎞️ GIF", value=False)
            ext_dcm = st.checkbox("🏥 DCM (DICOM)", value=True)
        
        with col_doc:
            st.markdown("*Documentos:*")
            ext_pdf = st.checkbox("📄 PDF", value=True)
            ext_docx = st.checkbox("📝 DOC/DOCX", value=False)
            ext_xlsx = st.checkbox("📊 XLS/XLSX", value=False)
            ext_txt = st.checkbox("📃 TXT", value=False)
        
        extensiones = []
        if ext_jpg: extensiones.extend([".jpg", ".jpeg"])
        if ext_png: extensiones.append(".png")
        if ext_gif: extensiones.append(".gif")
        if ext_dcm: extensiones.append(".dcm")
        if ext_pdf: extensiones.append(".pdf")
        if ext_docx: extensiones.extend([".doc", ".docx"])
        if ext_xlsx: extensiones.extend([".xls", ".xlsx"])
        if ext_txt: extensiones.append(".txt")
        
        if extensiones:
            st.success(f"✅ {len(extensiones)} tipo(s) seleccionado(s)")
        else:
            st.error("❌ Selecciona al menos un tipo de archivo")

with tab3:
    # Validaciones con validación mejorada de contraseñas
    errores = []
    if not email_usuario:
        errores.append("📧 Falta el email")
    if not is_real_password(password_usuario):
        errores.append("🔑 Falta configurar contraseña válida")
    if not extensiones:
        errores.append("📎 Falta seleccionar tipos de archivo")
    if usar_filtro_fecha and not fecha_valida:
        errores.append("📅 Rango de fechas inválido")
    
    if errores:
        st.error("❌ **Errores de configuración:**")
        for error in errores:
            st.error(f"• {error}")
        
        # Mostrar información adicional para configurar contraseñas
        if not is_real_password(password_usuario):
            st.info("""
            💡 **Para configurar las contraseñas:**
            
            1. **Crea/edita el archivo `.env` en la carpeta del proyecto:**
            ```
            # Gmail - Victor Personal (ya configurado)
            GMAIL_VICTOR_PERSONAL="tu_contraseña_de_app_gmail_aqui"
            
            # Gmail - Victor Consultorio
            GMAIL_VICTOR_CONSULTORIO="tu_contraseña_de_app_gmail_consultorio_aqui"
            
            # Hotmail - Victor Personal
            HOTMAIL_VICTOR_PERSONAL="tu_contraseña_de_app_hotmail_aqui"
            ```
            
            2. **Reemplaza los valores placeholder con las contraseñas reales**
            3. **Reinicia la aplicación** (Ctrl+C y `streamlit run app.py`)
            """)
    else:
        st.success("✅ **Configuración válida**")
        
        if st.button("🚀 **EJECUTAR ANÁLISIS COMPLETO**", type="primary", use_container_width=True):
            
            # Preparar configuración de fechas
            if usar_filtro_fecha:
                # Convertir dates a datetime para compatibilidad
                fecha_inicio_dt = datetime.combine(fecha_inicio, datetime.min.time())
                fecha_fin_dt = datetime.combine(fecha_fin, datetime.max.time())  # Final del día
                
                date_config = {
                    "enabled": True,
                    "start_date": fecha_inicio_dt,
                    "end_date": fecha_fin_dt,
                    "days_back": 0  # Ya no se usa, pero mantenemos por compatibilidad
                }
            else:
                date_config = {
                    "enabled": False,
                    "start_date": None,
                    "end_date": None,
                    "days_back": 0
                }
            
            # Crear configuración
            config = {
                "email_settings": {
                    "server": servidor,
                    "port": puerto,
                    "email": email_usuario,
                    "password": password_usuario,
                    "use_ssl": usar_ssl
                },
                "filters": {
                    "subject_keywords": palabras_clave,
                    "sender_emails": remitentes,
                    "date_range": date_config,
                    "has_attachments": True,
                    "folder": carpeta_email
                },
                "download_settings": {
                    "base_folder": carpeta_base,
                    "folder_structure": {
                        "by_date": por_fecha,
                        "by_sender": por_remitente,
                        "by_subject": por_asunto
                    },
                    "allowed_extensions": extensiones,
                    "max_file_size_mb": 0,
                    "rename_files": True,
                    "naming_pattern": "{date}_{sender}_{subject}_{index}_{original_name}",
                    "download_google_drive_links": True
                },
                "processing": {
                    "mark_as_read": False,
                    "delete_duplicates": False,
                    "max_emails_per_run": 0,
                    "delay_between_emails": 1.0
                },
                "logging": {
                    "level": "DEBUG",
                    "file": "email_downloader_con_reporte.log"
                }
            }
            
            # Ejecutar análisis
            progress_bar = st.progress(0)
            status_container = st.container()
            
            try:
                with status_container:
                    if usar_filtro_fecha:
                        st.info(f"🔄 Iniciando análisis desde {fecha_inicio.strftime('%d/%m/%Y')} hasta {fecha_fin.strftime('%d/%m/%Y')}...")
                    else:
                        st.info("🔄 Iniciando análisis de todo el historial...")
                
                downloader = EmailImageDownloader(config)
                progress_bar.progress(20)
                
                with status_container:
                    st.info("🔄 Conectando al email...")
                
                resultado = downloader.run()
                progress_bar.progress(100)
                
                if resultado:
                    with status_container:
                        st.markdown("""
                        <div class="status-success">
                            <h3>✅ ¡Análisis completado!</h3>
                            <p>Revisa el reporte CSV generado.</p>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    # Buscar archivo CSV generado
                    csv_files = list(Path('.').glob('reporte_analisis_emails_*.csv'))
                    
                    if csv_files:
                        csv_file = max(csv_files, key=lambda x: x.stat().st_mtime)
                        st.success(f"📊 Reporte: {csv_file.name}")
                        
                        try:
                            df = pd.read_csv(csv_file)
                            
                            col1, col2 = st.columns(2)
                            
                            with col1:
                                st.info(f"📈 Total emails: {len(df)}")
                                emails_descargados = len(df[df['estado'] == 'DESCARGADO'])
                                emails_descartados = len(df[df['estado'] == 'DESCARTADO'])
                                
                                st.write(f"✅ Descargados: {emails_descargados}")
                                st.write(f"❌ Descartados: {emails_descartados}")
                            
                            with col2:
                                motivos = df[df['estado'] == 'DESCARTADO']['motivo_rechazo'].value_counts()
                                if not motivos.empty:
                                    st.write("**Motivos más comunes:**")
                                    for motivo, count in motivos.head(3).items():
                                        st.write(f"• {count}: {motivo[:50]}...")
                            
                            # Preview del CSV
                            with st.expander("👁️ Preview del reporte", expanded=False):
                                st.dataframe(df.head(10))
                            
                            # Botón descarga CSV
                            with open(csv_file, 'rb') as f:
                                st.download_button(
                                    label="⬇️ Descargar Reporte CSV",
                                    data=f.read(),
                                    file_name=csv_file.name,
                                    mime="text/csv",
                                    use_container_width=True
                                )
                                
                        except Exception as e:
                            st.error(f"Error leyendo CSV: {e}")
                    
                    # Verificar archivos descargados
                    if os.path.exists(carpeta_base):
                        archivos = list(Path(carpeta_base).rglob('*'))
                        archivos_validos = [f for f in archivos if f.is_file()]
                        
                        if archivos_validos:
                            st.info(f"📦 {len(archivos_validos)} archivo(s) descargado(s)")
                            
                            # Crear ZIP
                            zip_filename = f"archivos_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
                            
                            with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                                for archivo in archivos_validos:
                                    ruta_relativa = archivo.relative_to(carpeta_base)
                                    zip_file.write(archivo, ruta_relativa)
                            
                            # Botón descarga ZIP
                            with open(zip_filename, 'rb') as zip_file:
                                st.download_button(
                                    label="⬇️ Descargar Archivos (ZIP)",
                                    data=zip_file.read(),
                                    file_name=zip_filename,
                                    mime="application/zip",
                                    use_container_width=True
                                )
                            
                            # Limpiar ZIP temporal
                            try:
                                os.unlink(zip_filename)
                            except:
                                pass
                        else:
                            st.warning("⚠️ No se descargaron archivos")
                else:
                    with status_container:
                        st.markdown("""
                        <div class="status-error">
                            <h3>❌ Error en el análisis</h3>
                        </div>
                        """, unsafe_allow_html=True)
                
            except Exception as e:
                progress_bar.progress(0)
                with status_container:
                    st.error(f"❌ Error: {str(e)}")

# Footer
st.markdown("---")
st.markdown("""
<div style="text-align: center; color: #666;">
    📧 <strong>SmartExtract Dr.Lucero</strong> | 
    Desarrollado por <strong>TrackingDatax</strong>
</div>
""", unsafe_allow_html=True)