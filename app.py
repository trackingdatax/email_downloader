#!/usr/bin/env python3
"""
Interfaz Streamlit para el Descargador de Emails
"""
import streamlit as st
import json
import os
from pathlib import Path
import pandas as pd
from datetime import datetime, timedelta
import tempfile
import threading
import time
import imaplib
import email
import re
import hashlib
import logging
import requests
import zipfile
from email.header import decode_header
from urllib.parse import urlparse, parse_qs

# Cargar variables de entorno desde .env en desarrollo local
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    # dotenv no está disponible, las variables se cargarán desde el sistema
    pass
# Configuración de la página
st.set_page_config(
    page_title="📧 SmartExtract Dr.Lucero",
    page_icon="📧",
    layout="wide",
    initial_sidebar_state="expanded"
)
# Clase EmailImageDownloader integrada
class EmailImageDownloader:
    def __init__(self, config_dict):
        self.config = config_dict
        self.setup_logging()
        self.mail = None
        
    def setup_logging(self):
        log_level = getattr(logging, self.config['logging']['level'].upper())
        
        # Limpiar handlers existentes
        for handler in logging.root.handlers[:]:
            logging.root.removeHandler(handler)
        
        logging.basicConfig(
            level=log_level,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(self.config['logging']['file'], encoding='utf-8'),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)
    
    def connect_to_email(self):
        try:
            email_config = self.config['email_settings']
            
            if email_config['use_ssl']:
                self.mail = imaplib.IMAP4_SSL(email_config['server'], email_config['port'])
            else:
                self.mail = imaplib.IMAP4(email_config['server'], email_config['port'])
            
            self.mail.login(email_config['email'], email_config['password'])
            self.logger.info(f"Conectado exitosamente a {email_config['email']}")
            return True
            
        except Exception as e:
            self.logger.error(f"Error al conectar: {e}")
            return False
    
    def search_emails(self):
        try:
            folder = self.config['filters']['folder']
            self.mail.select(folder)
            
            filters = self.config['filters']
            all_email_ids = set()
            
            # LÓGICA COMBINADA: Todos los filtros se combinan
            self.logger.info("=== INICIANDO BÚSQUEDA COMBINADA ===")
            
            # 1. Filtro base: rango de fechas
            base_criteria = []
            if filters['date_range']['enabled']:
                date_back = datetime.now() - timedelta(days=filters['date_range']['days_back'])
                date_str = date_back.strftime('%d-%b-%Y')
                base_criteria.append(f'SINCE {date_str}')
                self.logger.info(f"📅 Filtro de fecha: desde {date_str}")
            
            # 2. Si hay REMITENTES específicos, buscar por cada uno
            if filters['sender_emails']:
                self.logger.info(f"👥 Aplicando filtro de remitentes: {len(filters['sender_emails'])} remitente(s)")
                
                for sender in filters['sender_emails']:
                    # Buscar emails de este remitente
                    sender_criteria = base_criteria + [f'FROM "{sender}"']
                    
                    # 3. Si TAMBIÉN hay palabras clave, combinar con cada palabra
                    if filters['subject_keywords']:
                        self.logger.info(f"🔍 Combinando remitente '{sender}' con palabras clave en asunto")
                        
                        for keyword in filters['subject_keywords']:
                            combined_criteria = sender_criteria + [f'SUBJECT "{keyword}"']
                            search_string = ' '.join(combined_criteria)
                            
                            self.logger.info(f"  Buscando: DE '{sender}' + ASUNTO contiene '{keyword}'")
                            
                            result, messages = self.mail.search(None, search_string)
                            if result == 'OK' and messages[0]:
                                email_ids = messages[0].split()
                                all_email_ids.update(email_ids)
                                self.logger.info(f"  ✅ Encontrados {len(email_ids)} emails")
                    else:
                        # Solo remitente + fecha
                        search_string = ' '.join(sender_criteria)
                        self.logger.info(f"  Buscando: DE '{sender}' (sin filtro de asunto)")
                        
                        result, messages = self.mail.search(None, search_string)
                        if result == 'OK' and messages[0]:
                            email_ids = messages[0].split()
                            all_email_ids.update(email_ids)
                            self.logger.info(f"  ✅ Encontrados {len(email_ids)} emails")
            
            # 4. Si NO hay remitentes pero SÍ hay palabras clave
            elif filters['subject_keywords']:
                self.logger.info(f"🔍 Aplicando solo filtro de palabras clave: {len(filters['subject_keywords'])} palabra(s)")
                
                for keyword in filters['subject_keywords']:
                    keyword_criteria = base_criteria + [f'SUBJECT "{keyword}"']
                    search_string = ' '.join(keyword_criteria)
                    
                    self.logger.info(f"  Buscando: ASUNTO contiene '{keyword}'")
                    
                    result, messages = self.mail.search(None, search_string)
                    if result == 'OK' and messages[0]:
                        email_ids = messages[0].split()
                        all_email_ids.update(email_ids)
                        self.logger.info(f"  ✅ Encontrados {len(email_ids)} emails")
            
            # 5. Si NO hay remitentes NI palabras clave, buscar solo por fecha
            else:
                search_string = ' '.join(base_criteria) if base_criteria else 'ALL'
                self.logger.info(f"📧 Buscando TODOS los emails (solo filtro de fecha)")
                
                result, messages = self.mail.search(None, search_string)
                if result == 'OK' and messages[0]:
                    email_ids = messages[0].split()
                    all_email_ids.update(email_ids)
                    self.logger.info(f"✅ Encontrados {len(email_ids)} emails")
            
            final_email_list = list(all_email_ids)
            self.logger.info(f"📊 TOTAL de emails únicos encontrados (antes de filtrar por archivos): {len(final_email_list)}")
            self.logger.info("=== BÚSQUEDA COMBINADA COMPLETADA ===")
            
            return final_email_list
                
        except Exception as e:
            self.logger.error(f"Error al buscar emails: {e}")
            return []
    
    def decode_email_header(self, header):
        if header is None:
            return ""
        
        decoded_parts = decode_header(header)
        decoded_string = ""
        
        for part, encoding in decoded_parts:
            if isinstance(part, bytes):
                try:
                    decoded_string += part.decode(encoding or 'utf-8')
                except:
                    decoded_string += part.decode('utf-8', errors='ignore')
            else:
                decoded_string += part
        
        return decoded_string.strip()
    
    def clean_filename(self, filename):
        cleaned = re.sub(r'[<>:"/\\|?*]', '_', filename)
        cleaned = re.sub(r'\s+', '_', cleaned)
        cleaned = cleaned.strip('._')
        return cleaned[:100]
    
    def get_email_date(self, msg):
        try:
            date_header = msg['Date']
            if date_header:
                from email.utils import parsedate_to_datetime
                email_date = parsedate_to_datetime(date_header)
                return email_date.replace(tzinfo=None)
        except Exception as e:
            self.logger.debug(f"Error parseando fecha del email: {e}")
        
        return datetime.now()
    
    def create_folder_structure(self, msg, sender):
        base_folder = Path(self.config['download_settings']['base_folder'])
        folder_structure = self.config['download_settings']['folder_structure']
        
        current_path = base_folder
        email_date = self.get_email_date(msg)
        
        if folder_structure.get('by_date', False):
            date_folder = email_date.strftime('%Y-%m-%d')
            current_path = current_path / date_folder
        
        if folder_structure.get('by_sender', False):
            sender_clean = self.clean_filename(sender.split('@')[0])
            current_path = current_path / sender_clean
        
        if folder_structure.get('by_subject', False):
            subject = self.decode_email_header(msg['Subject'])
            subject_clean = self.clean_filename(subject)[:30]
            current_path = current_path / subject_clean
        
        current_path.mkdir(parents=True, exist_ok=True)
        return current_path, email_date
    
    def generate_filename(self, msg, sender, subject, index, original_name, email_date):
        pattern = self.config['download_settings']['naming_pattern']
        date_str = email_date.strftime('%Y%m%d_%H%M%S')
        
        clean_sender = self.clean_filename(sender.split('@')[0])
        clean_subject = self.clean_filename(subject)[:30]
        clean_original = self.clean_filename(original_name)
        
        filename = pattern.format(
            date=date_str,
            sender=clean_sender,
            subject=clean_subject,
            index=str(index).zfill(3),
            original_name=clean_original
        )
        
        return filename
    
    def get_file_hash(self, file_data):
        return hashlib.md5(file_data).hexdigest()
    
    def is_duplicate_in_day(self, file_path, file_data, email_date):
        if not self.config['processing']['delete_duplicates']:
            return False
        
        day_folder = file_path.parent
        new_hash = self.get_file_hash(file_data)
        
        if day_folder.exists():
            for existing_file in day_folder.rglob('*'):
                if existing_file.is_file() and existing_file != file_path:
                    try:
                        existing_hash = self.get_file_hash(existing_file.read_bytes())
                        if existing_hash == new_hash:
                            return True
                    except:
                        pass
        
        return False
    
    def extract_google_drive_links(self, html_content):
        drive_links = []
        
        patterns = [
            r'https://drive\.google\.com/file/d/([a-zA-Z0-9_-]+)',
            r'https://docs\.google\.com/document/d/([a-zA-Z0-9_-]+)',
            r'https://docs\.google\.com/spreadsheets/d/([a-zA-Z0-9_-]+)',
            r'https://docs\.google\.com/presentation/d/([a-zA-Z0-9_-]+)',
        ]
        
        for pattern in patterns:
            matches = re.findall(pattern, html_content)
            for match in matches:
                download_url = f"https://drive.google.com/uc?id={match}&export=download"
                drive_links.append({
                    'file_id': match,
                    'download_url': download_url,
                    'original_url': f"https://drive.google.com/file/d/{match}"
                })
        
        return drive_links
    
    def download_from_google_drive(self, drive_info, target_folder, base_filename, index, email_date):
        try:
            session = requests.Session()
            response = session.get(drive_info['download_url'], stream=True)
            
            if 'confirm=' in response.text:
                confirm_pattern = r'confirm=([^&]+)'
                confirm_match = re.search(confirm_pattern, response.text)
                if confirm_match:
                    confirm_token = confirm_match.group(1)
                    confirm_url = f"{drive_info['download_url']}&confirm={confirm_token}"
                    response = session.get(confirm_url, stream=True)
            
            filename = "archivo_drive"
            if 'Content-Disposition' in response.headers:
                content_disp = response.headers['Content-Disposition']
                filename_match = re.search(r'filename="([^"]+)"', content_disp)
                if filename_match:
                    filename = filename_match.group(1)
            
            if '.' not in filename:
                content_type = response.headers.get('Content-Type', '')
                if 'pdf' in content_type:
                    filename += '.pdf'
                elif 'document' in content_type:
                    filename += '.docx'
                elif 'spreadsheet' in content_type:
                    filename += '.xlsx'
                elif 'image' in content_type:
                    filename += '.jpg'
            
            if response.status_code != 200:
                return None
            
            file_data = b''
            for chunk in response.iter_content(chunk_size=8192):
                if chunk:
                    file_data += chunk
            
            file_ext = Path(filename).suffix.lower()
            allowed_extensions = self.config['download_settings']['allowed_extensions']
            
            if file_ext not in allowed_extensions:
                return None
            
            if self.config['download_settings']['rename_files']:
                final_filename = f"{base_filename}_{str(index).zfill(3)}_GDrive_{self.clean_filename(filename)}"
                if not final_filename.endswith(file_ext):
                    final_filename += file_ext
            else:
                final_filename = f"GDrive_{filename}"
            
            file_path = target_folder / final_filename
            
            if self.is_duplicate_in_day(file_path, file_data, email_date):
                return None
            
            counter = 1
            original_path = file_path
            while file_path.exists():
                stem = original_path.stem
                suffix = original_path.suffix
                file_path = original_path.parent / f"{stem}_{counter}{suffix}"
                counter += 1
            
            file_path.write_bytes(file_data)
            
            if file_path.exists() and file_path.stat().st_size > 0:
                self.logger.info(f"Descargado desde Google Drive: {file_path}")
                return file_path
            else:
                if file_path.exists():
                    file_path.unlink()
                return None
                
        except Exception as e:
            self.logger.error(f"Error descargando desde Google Drive: {e}")
            return None
    
    def has_relevant_attachments(self, msg):
        """Verifica si el email tiene archivos adjuntos de los tipos permitidos"""
        allowed_extensions = self.config['download_settings']['allowed_extensions']
        
        for part in msg.walk():
            if part.get_filename():
                filename = part.get_filename()
                file_ext = Path(filename).suffix.lower()
                
                if file_ext in allowed_extensions:
                    return True
        
        return False
    
    def download_images_from_email(self, email_id):
        try:
            result, msg_data = self.mail.fetch(email_id, '(RFC822)')
            
            if result != 'OK':
                return 0
            
            msg = email.message_from_bytes(msg_data[0][1])
            sender = self.decode_email_header(msg['From'])
            subject = self.decode_email_header(msg['Subject'])
            
            # Verificar si el email tiene archivos adjuntos relevantes
            if not self.has_relevant_attachments(msg):
                self.logger.debug(f"Email sin archivos relevantes - De: {sender}, Asunto: {subject}")
                return 0
            
            self.logger.info(f"Procesando email de {sender}: {subject}")
            
            email_date = self.get_email_date(msg)
            downloaded_count = 0
            
            # FASE 1: Recopilar archivos adjuntos válidos
            attachments_to_download = []
            
            for part in msg.walk():
                if part.get_filename():
                    filename = part.get_filename()
                    content_type = part.get_content_type()
                    
                    is_attachment = (
                        part.get_content_disposition() == 'attachment' or
                        (part.get_content_disposition() == 'inline' and filename) or
                        (filename and content_type in ['application/pdf', 'application/msword', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'])
                    )
                    
                    if is_attachment:
                        filename = self.decode_email_header(filename)
                        file_ext = Path(filename).suffix.lower()
                        allowed_extensions = self.config['download_settings']['allowed_extensions']
                        
                        if file_ext in allowed_extensions:
                            file_data = part.get_payload(decode=True)
                            if file_data:
                                max_size = self.config['download_settings']['max_file_size_mb'] * 1024 * 1024
                                if len(file_data) <= max_size:
                                    attachments_to_download.append({
                                        'filename': filename,
                                        'data': file_data,
                                        'ext': file_ext
                                    })
            
            # FASE 2: Buscar enlaces de Google Drive
            drive_links_to_download = []
            
            if self.config['download_settings'].get('download_google_drive_links', False):
                for part in msg.walk():
                    if part.get_content_type() == 'text/html':
                        try:
                            html_content = part.get_payload(decode=True).decode('utf-8')
                            if 'drive.google.com' in html_content or 'docs.google.com' in html_content:
                                drive_links = self.extract_google_drive_links(html_content)
                                if drive_links:
                                    drive_links_to_download.extend(drive_links)
                                break
                        except:
                            pass
            
            # FASE 3: Solo crear carpeta y descargar si hay archivos
            if attachments_to_download or drive_links_to_download:
                target_folder, email_date = self.create_folder_structure(msg, sender)
                
                # Descargar archivos adjuntos
                for idx, attachment in enumerate(attachments_to_download):
                    if self.config['download_settings']['rename_files']:
                        new_filename = self.generate_filename(
                            msg, sender, subject, idx, attachment['filename'], email_date
                        )
                        if not new_filename.endswith(attachment['ext']):
                            new_filename += attachment['ext']
                    else:
                        new_filename = attachment['filename']
                    
                    file_path = target_folder / new_filename
                    
                    if self.is_duplicate_in_day(file_path, attachment['data'], email_date):
                        self.logger.info(f"Archivo duplicado omitido: {new_filename}")
                        continue
                    
                    counter = 1
                    original_path = file_path
                    while file_path.exists():
                        stem = original_path.stem
                        suffix = original_path.suffix
                        file_path = original_path.parent / f"{stem}_{counter}{suffix}"
                        counter += 1
                    
                    file_path.write_bytes(attachment['data'])
                    downloaded_count += 1
                    self.logger.info(f"Descargado: {file_path}")
                
                # Descargar archivos de Google Drive
                if drive_links_to_download:
                    self.logger.info(f"  Encontrados {len(drive_links_to_download)} enlaces de Google Drive")
                    
                    drive_base_filename = self.generate_filename(
                        msg, sender, subject, 0, "drive_file", email_date
                    ).replace("_drive_file", "")
                    
                    for idx, drive_info in enumerate(drive_links_to_download):
                        result = self.download_from_google_drive(
                            drive_info, target_folder, drive_base_filename, idx + 100, email_date
                        )
                        if result:
                            downloaded_count += 1
            else:
                self.logger.debug(f"  No hay archivos descargables en este email")
            
            if self.config['processing']['mark_as_read']:
                self.mail.store(email_id, '+FLAGS', '\\Seen')
            
            return downloaded_count
            
        except Exception as e:
            self.logger.error(f"Error procesando email {email_id}: {e}")
            return 0
    
    def run(self):
        try:
            self.logger.info("Iniciando descarga de archivos desde emails")
            
            if not self.connect_to_email():
                return False
            
            email_ids = self.search_emails()
            
            if not email_ids:
                self.logger.info("No se encontraron emails que coincidan con los criterios")
                return True
            
            max_emails = self.config['processing']['max_emails_per_run']
            if max_emails > 0:
                email_ids = email_ids[:max_emails]
            
            total_downloaded = 0
            processed_emails = 0
            emails_with_attachments = 0
            
            for email_id in email_ids:
                downloaded = self.download_images_from_email(email_id)
                if downloaded > 0:
                    emails_with_attachments += 1
                total_downloaded += downloaded
                processed_emails += 1
                
                delay = self.config['processing']['delay_between_emails']
                if delay > 0:
                    time.sleep(delay)
            
            self.logger.info(f"Proceso completado: {processed_emails} emails procesados, "
                           f"{emails_with_attachments} emails con archivos relevantes, "
                           f"{total_downloaded} archivos descargados")
            
            return True
            
        except Exception as e:
            self.logger.error(f"Error en ejecución: {e}")
            return False
        
        finally:
            if self.mail:
                try:
                    self.mail.close()
                    self.mail.logout()
                except:
                    pass

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
    .config-section {
        background-color: #f8f9fa;
        padding: 1rem;
        border-radius: 8px;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

# Header principal
st.markdown("""
<div class="main-header">
    <h1>📧 SmartExtract Dr.Lucero</h1>
    <p>Configura y descarga archivos desde tus emails de forma automática</p>
</div>
""", unsafe_allow_html=True)

# Configuraciones predefinidas de usuarios
USUARIOS_PREDEFINIDOS = {
    "Gmail - Personal": {
        "server": "imap.gmail.com",
        "port": 993,
        "use_ssl": True,
        "tipo": "Gmail",
        "instrucciones": "⚠️ **IMPORTANTE:** Gmail requiere configuración especial:\n\n1. Activar verificación en 2 pasos\n\n2. Ir a: **https://myaccount.google.com/apppasswords**\n\n3. Generar contraseña para 'Correo'\n\n4. Usar esa contraseña de 16 caracteres",
        "tip": "💡 La contraseña se ve así: 'abcd efgh ijkl mnop'"
    },
    "Outlook/Hotmail": {
        "server": "outlook.office365.com",
        "port": 993,
        "use_ssl": True,
        "tipo": "Outlook",
        "instrucciones": "⚠️ **IMPORTANTE:** Outlook/Hotmail requiere:\n1. Activar verificación en 2 pasos en tu cuenta Microsoft\n2. Ir a Seguridad → Opciones de seguridad avanzadas\n3. Generar una 'Contraseña de aplicación'\n4. Usar esa contraseña generada (no tu contraseña normal)",
        "tip": "💡 Alternativamente, algunos usuarios pueden usar la contraseña normal si tienen IMAP habilitado"
    },
    "Yahoo Mail": {
        "server": "imap.mail.yahoo.com",
        "port": 993,
        "use_ssl": True,
        "tipo": "Yahoo",
        "instrucciones": "⚠️ **IMPORTANTE:** Yahoo requiere:\n1. Ir a Configuración de cuenta → Seguridad de la cuenta\n2. Activar verificación en 2 pasos\n3. Generar una 'Contraseña de aplicación'\n4. Usar esa contraseña (no tu contraseña normal)",
        "tip": "💡 Yahoo también requiere tener IMAP activado en la configuración de correo"
    },
    "iCloud (Apple)": {
        "server": "imap.mail.me.com",
        "port": 993,
        "use_ssl": True,
        "tipo": "iCloud",
        "instrucciones": "⚠️ **IMPORTANTE:** iCloud requiere:\n1. Activar autenticación de dos factores en tu Apple ID\n2. Ir a appleid.apple.com → Seguridad\n3. Generar una 'Contraseña específica de la app'\n4. Usar esa contraseña de 16 caracteres",
        "tip": "💡 Solo funciona con cuentas @icloud.com, @me.com, @mac.com"
    },
    "Gmail Corporativo": {
        "server": "imap.gmail.com",
        "port": 993,
        "use_ssl": True,
        "tipo": "Gmail Corporativo",
        "instrucciones": "⚠️ **IMPORTANTE:** Gmail corporativo puede requerir:\n1. Verificar con tu administrador IT si IMAP está habilitado\n2. Algunos requieren OAuth2 (más complejo)\n3. Otros permiten contraseñas de aplicación como Gmail personal\n4. Contacta a tu IT para confirmar el método",
        "tip": "💡 Cada empresa tiene políticas diferentes"
    },
    "Personalizado": {
        "server": "",
        "port": 993,
        "use_ssl": True,
        "tipo": "Personalizado",
        "instrucciones": "ℹ️ **Configuración manual:** Ingresa los datos IMAP de tu proveedor.\nConsulta la documentación de tu proveedor de email para obtener:\n- Servidor IMAP\n- Puerto (usualmente 993 para SSL)\n- Si requiere contraseña de aplicación",
        "tip": "💡 La mayoría de proveedores modernos requieren contraseñas de aplicación"
    }
}

# Cuentas de email predefinidas con variables de entorno
CUENTAS_EMAIL = {
    "victorlucero1981@gmail.com": {
        "email": "victorlucero1981@gmail.com",
        "password": os.getenv("GMAIL_VICTOR_PERSONAL", ""),
        "server": "imap.gmail.com",
        "port": 993,
        "use_ssl": True,
        "tipo": "Gmail Personal",
        "env_var": "GMAIL_VICTOR_PERSONAL"
    },
    "victorLucero.consultorio@gmail.com": {
        "email": "victorLucero.consultorio@gmail.com", 
        "password": os.getenv("GMAIL_VICTOR_CONSULTORIO", ""),
        "server": "imap.gmail.com",
        "port": 993,
        "use_ssl": True,
        "tipo": "Gmail Consultorio",
        "env_var": "GMAIL_VICTOR_CONSULTORIO"
    },
    "victorlucero1981@hotmail.com": {
        "email": "victorlucero1981@hotmail.com",
        "password": os.getenv("HOTMAIL_VICTOR_PERSONAL", ""),
        "server": "outlook.office365.com", 
        "port": 993,
        "use_ssl": True,
        "tipo": "Hotmail Personal",
        "env_var": "HOTMAIL_VICTOR_PERSONAL"
    }
}

# Sidebar para configuración
with st.sidebar:
    st.header("⚙️ Configuración")
    
    # Selección de cuenta de email predefinida
    st.subheader("📧 Seleccionar Cuenta de Email")
    
    cuenta_seleccionada = st.selectbox(
        "🔧 Cuenta de Email",
        options=list(CUENTAS_EMAIL.keys()),
        help="Selecciona una de tus cuentas de email predefinidas"
    )
    
    config_cuenta = CUENTAS_EMAIL[cuenta_seleccionada]
    
    # Mostrar información de la cuenta seleccionada
    st.info(f"""
    **Email:** {config_cuenta['email']}  
    **Tipo:** {config_cuenta['tipo']}  
    **Servidor:** {config_cuenta['server']}:{config_cuenta['port']}  
    **SSL:** {'Sí' if config_cuenta['use_ssl'] else 'No'}
    """)
    
    # Configurar valores automáticamente
    email_usuario = config_cuenta['email']
    password_usuario = config_cuenta['password']
    servidor = config_cuenta['server']
    puerto = config_cuenta['port']
    usar_ssl = config_cuenta['use_ssl']
    
    # Mostrar estado de la contraseña
    # Verificar si es una contraseña real o placeholder
    is_placeholder = (
        not password_usuario or 
        password_usuario.strip() == "" or
        "tu_contraseña" in password_usuario.lower() or
        "aqui" in password_usuario.lower()
    )
    
    if password_usuario and not is_placeholder:
        st.success("✅ Contraseña de aplicación configurada")
    else:
        st.error("❌ Falta configurar contraseña de aplicación para esta cuenta")
    
    st.markdown("---")
    
    # Mostrar instrucciones específicas del tipo de cuenta
    tipo_cuenta = "Gmail - Personal" if "gmail.com" in email_usuario else "Outlook/Hotmail"
    config_servidor = USUARIOS_PREDEFINIDOS[tipo_cuenta]
    
    st.subheader("📋 Instrucciones de Configuración")
    
    # Mostrar instrucciones en un container expandible
    with st.expander(f"🔧 Cómo configurar {config_servidor['tipo']}", expanded=False):
        st.markdown(config_servidor["instrucciones"])
        if "tip" in config_servidor:
            st.info(config_servidor["tip"])
        
        # Enlaces directos para los principales proveedores
        if config_servidor["tipo"] == "Gmail":
            col1, col2 = st.columns(2)
            with col1:
                st.markdown("**[Setup de Seguridad](https://myaccount.google.com/security)**", unsafe_allow_html=True)
            with col2:
                st.markdown("**[Crear App Password](https://myaccount.google.com/apppasswords)**", unsafe_allow_html=True)
        
        elif config_servidor["tipo"] == "Outlook":
            st.markdown("**[Configuración de Microsoft](https://account.live.com/proofs/manage/additional)**", unsafe_allow_html=True)

# Tabs principales
tab1, tab2, tab3 = st.tabs(["🎯 Filtros y Descarga", "📁 Configuración de Archivos", "📊 Resultados"])

with tab1:
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.subheader("📮 Filtros de Email")
        
        # Remitentes
        st.markdown("**👥 Emails Remitentes a buscar (uno por línea):**")
        remitentes_text = st.text_area(
            "",
            placeholder="empresa@ejemplo.com\nfacturas@proveedor.com\nrrhh@miempresa.com",
            height=100,
            help="Ingresa un email por línea"
        )
        
        remitentes = [email.strip() for email in remitentes_text.split('\n') if email.strip()]
        
        if remitentes:
            st.success(f"✅ {len(remitentes)} remitente(s) configurado(s)")
        
        # Rango de fechas
        st.markdown("**📅 Rango de fechas:**")
        
        # Calcular días desde 01-01-2020 hasta hoy
        fecha_inicio = datetime(2020, 1, 1)
        fecha_actual = datetime.now()
        dias_maximos = (fecha_actual - fecha_inicio).days
        
        dias_atras = st.slider(
            "Buscar emails de los últimos X días",
            min_value=1,
            max_value=dias_maximos,
            value=30,
            help=f"El script buscará emails desde hace X días hasta hoy (máximo desde 01-01-2020: {dias_maximos} días)"
        )
        
        # Mostrar fecha de inicio calculada
        fecha_busqueda = fecha_actual - timedelta(days=dias_atras)
        st.info(f"📅 Buscando emails desde: **{fecha_busqueda.strftime('%d/%m/%Y')}** hasta hoy")
    
    with col2:
        st.subheader("🔍 Filtros Adicionales")
        
        # Palabras clave en asunto (opcional)
        st.markdown("**🔍 Palabras clave en asunto (opcional y una por línea):**")
        palabras_clave_text = st.text_area(
            "",
            value="tcmaxonline\ntomografia\nradiografia\npanoramica\nrx",
            placeholder="tcmaxonline\ntomografia\nradiografia\npanoramica\nrx",
            height=120,
            help="Opcional: buscar solo emails que contengan estas palabras en el asunto"
        )
        
        palabras_clave = [palabra.strip() for palabra in palabras_clave_text.split('\n') if palabra.strip()]
        
        # Valores fijos ocultos al usuario
        carpeta_email = "INBOX"  # Valor por defecto, pero podrías cambiarlo para buscar en todas
        max_emails = 100
        delay_emails = 0.5


with tab2:
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.subheader("📁 Configuración de Carpetas")
        
        carpeta_base = st.text_input(
            "📂 Carpeta base de descarga",
            value="./archivos_medicos_descargados",
            help="Carpeta donde se guardarán todos los archivos"
        )
        
        st.markdown("**🗂️ Estructura de carpetas:**")
        por_fecha = st.checkbox("📅 Organizar por fecha", value=True)
        por_remitente = st.checkbox("👤 Organizar por remitente", value=True)
        por_asunto = st.checkbox("📝 Organizar por asunto", value=False)
        
        if por_fecha and por_remitente:
            st.info("📁 Estructura: `carpeta_base/2025-01-20/remitente/archivos`")
        elif por_fecha:
            st.info("📁 Estructura: `carpeta_base/2025-01-20/archivos`")
        elif por_remitente:
            st.info("📁 Estructura: `carpeta_base/remitente/archivos`")
    
    with col2:
        st.subheader("📄 Configuración de Archivos")
        
        # Extensiones permitidas - VALORES POR DEFECTO MODIFICADOS
        st.markdown("**📎 Tipos de archivo a descargar:**")
        
        col_img, col_doc = st.columns(2)
        
        with col_img:
            st.markdown("*Imágenes:*")
            ext_jpg = st.checkbox("📷 JPG/JPEG", value=True)  # Por defecto: True
            ext_png = st.checkbox("🖼️ PNG", value=False)      # Por defecto: False
            ext_gif = st.checkbox("🎞️ GIF", value=False)      # Por defecto: False
            ext_webp = st.checkbox("🌐 WEBP", value=False)    # Por defecto: False
            ext_dcm = st.checkbox("🏥 DCM (DICOM)", value=True)  # Por defecto: True - renombrado para claridad
        
        with col_doc:
            st.markdown("*Documentos:*")
            ext_pdf = st.checkbox("📄 PDF", value=True)       # Por defecto: True
            ext_docx = st.checkbox("📝 DOC/DOCX", value=False) # Por defecto: False
            ext_xlsx = st.checkbox("📊 XLS/XLSX", value=False) # Por defecto: False
            ext_txt = st.checkbox("📃 TXT", value=False)       # Por defecto: False
        
        extensiones = []
        if ext_jpg: extensiones.extend([".jpg", ".jpeg"])
        if ext_png: extensiones.append(".png")
        if ext_gif: extensiones.append(".gif")
        if ext_webp: extensiones.append(".webp")
        if ext_dcm: extensiones.append(".dcm")
        if ext_pdf: extensiones.append(".pdf")
        if ext_docx: extensiones.extend([".doc", ".docx"])
        if ext_xlsx: extensiones.extend([".xls", ".xlsx"])
        if ext_txt: extensiones.append(".txt")
        
        if extensiones:
            st.success(f"✅ {len(extensiones)} tipo(s) de archivo seleccionado(s)")
            st.info(f"📋 Tipos: {', '.join(extensiones)}")
        else:
            st.error("❌ Debes seleccionar al menos un tipo de archivo")
        
        # Otras opciones
        tamaño_max = st.slider(
            "📏 Tamaño máximo por archivo (MB)",
            min_value=1,
            max_value=100,
            value=25,  # Aumentado para archivos médicos
            help="Archivos más grandes serán omitidos"
        )
        
        renombrar_archivos = st.checkbox(
            "🏷️ Renombrar archivos automáticamente",
            value=True,
            help="Renombra archivos con fecha, remitente y asunto"
        )
        
        descargar_google_drive = st.checkbox(
            "☁️ Descargar enlaces de Google Drive",
            value=True,
            help="También descarga archivos compartidos por Google Drive en los emails"
        )

with tab3:
    st.subheader("📊 Ejecutar Descarga")
    
    # Validaciones
    errores = []
    if not email_usuario:
        errores.append("📧 Falta el email del usuario")
    if not password_usuario:
        errores.append("🔑 Falta la contraseña")
    if not extensiones:
        errores.append("📎 Falta seleccionar tipos de archivo")
    
    if errores:
        st.error("❌ **Errores de configuración:**\n" + "\n".join([f"• {error}" for error in errores]))
    else:
        st.success("✅ **Configuración válida** - Lista para ejecutar")
        
        # Mostrar resumen de configuración
        with st.expander("📋 Ver resumen de configuración", expanded=False):
            st.json({
                "email": email_usuario,
                "servidor": f"{servidor}:{puerto}",
                "remitentes": remitentes if remitentes else "TODOS",
                "palabras_clave": palabras_clave if palabras_clave else "TODAS",
                "dias_busqueda": dias_atras,
                "tipos_archivo": extensiones
            })
        
        # Botón para ejecutar
        if st.button("🚀 **EJECUTAR DESCARGA**", type="primary", use_container_width=True):
            
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
                    "date_range": {
                        "enabled": True,
                        "days_back": dias_atras
                    },
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
                    "max_file_size_mb": tamaño_max,
                    "rename_files": renombrar_archivos,
                    "naming_pattern": "{date}_{sender}_{subject}_{index}_{original_name}",
                    "download_google_drive_links": descargar_google_drive
                },
                "processing": {
                    "mark_as_read": False,
                    "delete_duplicates": True,
                    "max_emails_per_run": max_emails,
                    "delay_between_emails": delay_emails
                },
                "logging": {
                    "level": "INFO",
                    "file": "email_downloader.log"
                }
            }
            
            # Ejecutar descarga
            progress_bar = st.progress(0)
            status_container = st.container()
            
            try:
                with status_container:
                    st.info("🔄 Iniciando descarga...")
                
                # Crear instancia del descargador
                downloader = EmailImageDownloader(config)
                
                # Actualizar progreso
                progress_bar.progress(20)
                
                with status_container:
                    st.info("🔄 Conectando al servidor de email...")
                
                # Ejecutar
                resultado = downloader.run()
                
                progress_bar.progress(100)
                
                if resultado:
                    with status_container:
                        st.markdown("""
                        <div class="status-success">
                            <h3>✅ ¡Descarga completada exitosamente!</h3>
                            <p>Revisa los archivos descargados a continuación.</p>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    # Crear ZIP con estructura de carpetas para descarga
                    if os.path.exists(carpeta_base):
                        archivos_encontrados = list(Path(carpeta_base).rglob('*'))
                        archivos_validos = [f for f in archivos_encontrados if f.is_file()]
                        
                        if archivos_validos:
                            # Crear ZIP manteniendo estructura
                            zip_filename = f"emails_descargados_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
                            
                            with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                                for archivo in archivos_validos:
                                    # Mantener la ruta relativa dentro del ZIP
                                    ruta_relativa = archivo.relative_to(carpeta_base)
                                    zip_file.write(archivo, ruta_relativa)
                            
                            # Mostrar información de archivos
                            st.info(f"📦 Se creó un archivo ZIP con {len(archivos_validos)} archivo(s) manteniendo la estructura de carpetas")
                            
                            # Botón de descarga del ZIP
                            with open(zip_filename, 'rb') as zip_file:
                                st.download_button(
                                    label="⬇️ Descargar todos los archivos (ZIP)",
                                    data=zip_file.read(),
                                    file_name=zip_filename,
                                    mime="application/zip",
                                    use_container_width=True
                                )
                                    
                            # Limpiar archivo ZIP temporal
                            try:
                                os.unlink(zip_filename)
                            except:
                                pass
                        else:
                            st.warning("No se encontraron archivos descargados para crear el ZIP")
                else:
                    with status_container:
                        st.markdown("""
                        <div class="status-error">
                            <h3>❌ Error en la descarga</h3>
                            <p>Revisa el log para más detalles.</p>
                        </div>
                        """, unsafe_allow_html=True)
                
            except Exception as e:
                progress_bar.progress(0)
                with status_container:
                    st.error(f"❌ Error durante la ejecución: {str(e)}")

# Footer
st.markdown("---")
st.markdown("""
<div style="text-align: center; color: #666; padding: 1rem;">
    📧 <strong>SmartExtract Dr.Lucero</strong> | 
    Desarrollado por <strong>TrackingDatax</strong>
</div>
""", unsafe_allow_html=True)