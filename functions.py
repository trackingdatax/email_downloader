import email
import email.header
import imaplib
import logging
from datetime import datetime, timedelta
import re
from pathlib import Path
import os
import hashlib
import csv
import requests
from urllib.parse import urlparse, parse_qs

class EmailImageDownloader:
    def __init__(self, config):
        self.config = config
        self.mail = None
        self.logger = self.setup_logger()
        self.report_data = []
        self.duplicates_cache = {}
        
    def setup_logger(self):
        logger = logging.getLogger('EmailDownloader')
        logger.setLevel(logging.DEBUG)
        
        if not logger.handlers:
            handler = logging.StreamHandler()
            formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
            handler.setFormatter(formatter)
            logger.addHandler(handler)
        
        return logger
    
    def connect_to_email(self):
        """Conecta al servidor de email"""
        try:
            # Debug: mostrar configuración
            self.logger.debug(f"🔧 Configuración recibida:")
            self.logger.debug(f"  • Config keys: {list(self.config.keys())}")
            
            # Verificar si tenemos la estructura nueva (email_settings) o la antigua (imap + credentials)
            if 'email_settings' in self.config:
                # Estructura nueva de app.py
                email_config = self.config['email_settings']
                server = email_config.get('server')
                port = email_config.get('port')
                email_addr = email_config.get('email')
                password = email_config.get('password')
                use_ssl = email_config.get('use_ssl', True)
                
                self.logger.debug(f"🔧 Usando estructura 'email_settings':")
                self.logger.debug(f"  • Server: {server}")
                self.logger.debug(f"  • Port: {port}")
                self.logger.debug(f"  • Email: {email_addr}")
                self.logger.debug(f"  • Use SSL: {use_ssl}")
                
            elif 'imap' in self.config and 'credentials' in self.config:
                # Estructura antigua
                imap_config = self.config['imap']
                credentials_config = self.config['credentials']
                server = imap_config.get('server')
                port = imap_config.get('port')
                email_addr = credentials_config.get('email')
                password = credentials_config.get('password')
                use_ssl = True  # Por defecto SSL para IMAP
                
                self.logger.debug(f"🔧 Usando estructura 'imap' + 'credentials':")
                self.logger.debug(f"  • Server: {server}")
                self.logger.debug(f"  • Port: {port}")
                self.logger.debug(f"  • Email: {email_addr}")
                
            else:
                self.logger.error("❌ No se encontró configuración de email válida")
                self.logger.error(f"   Claves disponibles: {list(self.config.keys())}")
                return False
            
            # Validar que tenemos todos los datos necesarios
            if not all([server, port, email_addr, password]):
                self.logger.error("❌ Faltan datos de configuración de email")
                self.logger.error(f"   Server: {'✅' if server else '❌'}")
                self.logger.error(f"   Port: {'✅' if port else '❌'}")
                self.logger.error(f"   Email: {'✅' if email_addr else '❌'}")
                self.logger.error(f"   Password: {'✅' if password else '❌'}")
                return False
            
            self.logger.debug(f"🔌 Conectando a {server}:{port}")
            
            if use_ssl:
                self.mail = imaplib.IMAP4_SSL(server, port)
            else:
                self.mail = imaplib.IMAP4(server, port)
            
            self.logger.debug(f"🔐 Autenticando con {email_addr}")
            self.mail.login(email_addr, password)
            
            self.logger.debug(f"📬 Seleccionando INBOX")
            self.mail.select('INBOX')
            
            self.logger.info(f"Conectado exitosamente a {email_addr}")
            return True
            
        except Exception as e:
            self.logger.error(f"Error conectando: {e}")
            self.logger.error(f"Tipo de error: {type(e).__name__}")
            
            # Información adicional de debug
            self.logger.debug(f"🔧 Config completo recibido: {list(self.config.keys())}")
            
            return False
    # AGREGA ESTA FUNCIÓN TEMPORAL A TU functions.py para investigar




    # Añadir estas funciones adicionales
    def extract_base64_images(html_content):
        """Extrae imágenes en base64 del contenido HTML"""
        import re
        base64_images = []
        pattern = r'src="data:image/(\w+);base64,([a-zA-Z0-9+/=]+)"'
        
        matches = re.findall(pattern, html_content)
        for ext, data in matches:
            base64_images.append({
                'type': 'base64',
                'extension': f".{ext}",
                'data': data
            })
        
        return base64_images

    def extract_embedded_attachments(msg):
        """Extrae archivos embebidos usando Content-ID"""
        embedded = {}
        for part in msg.walk():
            content_id = part.get('Content-ID', '')
            if content_id:
                content_id = content_id.strip('<>')
                filename = part.get_filename()
                payload = part.get_payload(decode=True)
                if payload:
                    embedded[content_id] = {
                        'filename': filename,
                        'data': payload,
                        'content_type': part.get_content_type()
                    }
        return embedded

    def investigar_emails_patricia(self):
        """
        Función de investigación: Busca TODOS los emails de Patricia con adjuntos
        """
        if not self.connect_to_email():
            return
        
        try:
            self.logger.info("🔍 INVESTIGANDO TODOS LOS EMAILS DE PATRICIA...")
            
            # Buscar emails de Patricia en un rango amplio
            status, messages = self.mail.search(None, 'FROM "ferrari_patricia@yahoo.com"')
            
            if status != 'OK':
                self.logger.error("Error buscando emails de Patricia")
                return
            
            email_ids = messages[0].split()
            self.logger.info(f"📊 Encontrados {len(email_ids)} emails de Patricia")
            
            for email_id in email_ids:
                try:
                    result, msg_data = self.mail.fetch(email_id, '(RFC822)')
                    if result != 'OK':
                        continue
                    
                    msg = email.message_from_bytes(msg_data[0][1])
                    sender = self.decode_email_header(msg['From'])
                    subject = self.decode_email_header(msg['Subject'])
                    email_date = self.get_email_date(msg)
                    
                    # Contar partes del email
                    parts = list(msg.walk())
                    total_parts = len(parts)
                    
                    # Buscar archivos en cada parte
                    attachments_found = []
                    for i, part in enumerate(parts):
                        filename = part.get_filename()
                        content_type = part.get_content_type()
                        
                        if filename:
                            attachments_found.append(f"Parte {i}: {filename} ({content_type})")
                        
                        # Verificar si tiene contenido binario grande
                        try:
                            payload = part.get_payload(decode=True)
                            if payload and len(payload) > 1000:
                                if isinstance(payload, bytes):
                                    # Magic bytes check
                                    magic = payload[:10].hex() if len(payload) >= 10 else "N/A"
                                    if payload.startswith(b'\xff\xd8\xff'):
                                        attachments_found.append(f"Parte {i}: JPEG detectado por magic bytes! ({len(payload)} bytes)")
                                    elif len(payload) > 10000:  # Archivo grande
                                        attachments_found.append(f"Parte {i}: Archivo binario grande ({len(payload)} bytes, magic: {magic})")
                        except:
                            pass
                    
                    # Reportar este email
                    self.logger.info(f"📧 EMAIL #{email_id.decode()}")
                    self.logger.info(f"   📅 Fecha: {email_date.strftime('%Y-%m-%d %H:%M:%S')}")
                    self.logger.info(f"   👤 De: {sender}")
                    self.logger.info(f"   📝 Asunto: {subject}")
                    self.logger.info(f"   🔢 Total partes: {total_parts}")
                    
                    if attachments_found:
                        self.logger.info(f"   📎 ARCHIVOS ENCONTRADOS:")
                        for attachment in attachments_found:
                            self.logger.info(f"      • {attachment}")
                    else:
                        self.logger.info(f"   ❌ Sin archivos adjuntos detectados")
                    
                    self.logger.info(f"   ---")
                    
                    # Si este email es del 4 de abril de 2023, hacer análisis extra
                    if email_date.date().strftime('%Y-%m-%d') == '2023-04-04':
                        self.logger.info(f"🎯 ESTE ES EL EMAIL DEL 4 DE ABRIL - ANÁLISIS EXTRA:")
                        
                        # Mostrar contenido HTML completo para buscar links
                        for part in msg.walk():
                            if part.get_content_type() == 'text/html':
                                try:
                                    html_content = part.get_payload(decode=True).decode('utf-8')
                                    self.logger.info(f"   📄 CONTENIDO HTML COMPLETO:")
                                    self.logger.info(f"   {html_content}")
                                    
                                    # Buscar URLs de imágenes en el HTML
                                    import re
                                    image_urls = re.findall(r'https?://[^\s<>"\']+\.(?:jpg|jpeg|png|gif)', html_content, re.IGNORECASE)
                                    if image_urls:
                                        self.logger.info(f"   🔗 URLs DE IMÁGENES ENCONTRADAS:")
                                        for url in image_urls:
                                            self.logger.info(f"      • {url}")
                                    
                                except Exception as e:
                                    self.logger.info(f"   Error analizando HTML: {e}")
                    
                except Exception as e:
                    self.logger.error(f"Error analizando email {email_id}: {e}")
        
        finally:
            if self.mail:
                self.mail.close()
                self.mail.logout()


    # TAMBIÉN AGREGA ESTA FUNCIÓN PARA BUSCAR POR FECHA Y ASUNTO
    def buscar_emails_rx_abril_2023(self):
        """
        Busca específicamente emails con 'rx' en abril 2023
        """
        if not self.connect_to_email():
            return
        
        try:
            self.logger.info("🔍 BUSCANDO EMAILS CON 'RX' EN ABRIL 2023...")
            
            # Buscar emails en abril 2023 con "rx" en el asunto
            search_criteria = 'SINCE 01-Apr-2023 BEFORE 01-May-2023 SUBJECT "rx"'
            status, messages = self.mail.search(None, search_criteria)
            
            if status != 'OK':
                self.logger.error("Error en búsqueda")
                return
            
            email_ids = messages[0].split()
            self.logger.info(f"📊 Encontrados {len(email_ids)} emails con 'rx' en abril 2023")
            
            for email_id in email_ids:
                try:
                    result, msg_data = self.mail.fetch(email_id, '(RFC822)')
                    if result != 'OK':
                        continue
                    
                    msg = email.message_from_bytes(msg_data[0][1])
                    sender = self.decode_email_header(msg['From'])
                    subject = self.decode_email_header(msg['Subject'])
                    email_date = self.get_email_date(msg)
                    
                    self.logger.info(f"📧 EMAIL: {sender} - {subject} ({email_date.strftime('%Y-%m-%d')})")
                    
                    # Verificar si tiene adjuntos reales
                    parts = list(msg.walk())
                    self.logger.info(f"   Partes: {len(parts)}")
                    
                    for i, part in enumerate(parts):
                        filename = part.get_filename()
                        content_type = part.get_content_type()
                        content_disposition = part.get('Content-Disposition', '')
                        
                        self.logger.info(f"   Parte {i}: {content_type}")
                        if filename:
                            self.logger.info(f"      📎 Filename: {filename}")
                        if 'attachment' in content_disposition.lower():
                            self.logger.info(f"      📎 Es attachment!")
                        
                        # Verificar payload binario
                        try:
                            payload = part.get_payload(decode=True)
                            if payload and isinstance(payload, bytes) and len(payload) > 1000:
                                magic = payload[:10].hex()
                                self.logger.info(f"      📊 Binario: {len(payload)} bytes, magic: {magic}")
                                
                                if payload.startswith(b'\xff\xd8\xff'):
                                    self.logger.info(f"      🎯 ¡JPEG ENCONTRADO!")
                        except:
                            pass
                    
                except Exception as e:
                    self.logger.error(f"Error: {e}")
        
        finally:
            if self.mail:
                self.mail.close()
                self.mail.logout()    




    def normalize_text_for_search(self, text):
        """Normaliza texto para búsqueda (elimina acentos, convierte a minúsculas)"""
        if not text:
            return ""
        
        # Convertir a minúsculas
        text = text.lower()
        
        # Eliminar acentos comunes en español
        replacements = {
            'á': 'a', 'é': 'e', 'í': 'i', 'ó': 'o', 'ú': 'u', 'ü': 'u',
            'ñ': 'n', 'ç': 'c'
        }
        
        for accented, plain in replacements.items():
            text = text.replace(accented, plain)
        
        return text
    
    def decode_email_header(self, header):
        """Decodifica headers de email que pueden estar en diferentes encodings"""
        if not header:
            return ""
        
        try:
            decoded_parts = email.header.decode_header(header)
            decoded_header = ""
            
            for part, encoding in decoded_parts:
                if isinstance(part, bytes):
                    if encoding:
                        try:
                            decoded_header += part.decode(encoding)
                        except (UnicodeDecodeError, LookupError):
                            decoded_header += part.decode('utf-8', errors='ignore')
                    else:
                        decoded_header += part.decode('utf-8', errors='ignore')
                else:
                    decoded_header += str(part)
            
            return decoded_header
        except Exception as e:
            self.logger.debug(f"Error decodificando header: {e}")
            return str(header)
    
    def get_email_date(self, msg):
        """Obtiene la fecha del email"""
        date_header = msg['Date']
        if date_header:
            try:
                return email.utils.parsedate_to_datetime(date_header)
            except:
                pass
        return datetime.now()
    
    def search_emails_by_date_range(self):
        """Busca emails en el rango de fechas configurado usando el nuevo sistema mejorado"""
        self.logger.info("=== BUSCANDO EMAILS CON NUEVO SISTEMA DE FECHAS ===")
        
        filters = self.config['filters']
        self.logger.info(f"🔧 DEBUG - Configuración de filtros:")
        self.logger.info(f"  • date_range enabled: {filters['date_range']['enabled']}")
        self.logger.info(f"  • start_date: {filters['date_range'].get('start_date')}")
        self.logger.info(f"  • end_date: {filters['date_range'].get('end_date')}")
        
        if not filters['date_range']['enabled']:
            self.logger.warning("⚠️ Filtro de fecha deshabilitado - buscando TODOS los emails")
            status, messages = self.mail.search(None, 'ALL')
        else:
            start_date = filters['date_range'].get('start_date')
            end_date = filters['date_range'].get('end_date')
            
            if not start_date or not end_date:
                self.logger.error("❌ Fechas no configuradas correctamente")
                return []
            
            # Convertir fechas al formato IMAP
            start_imap = start_date.strftime('%d-%b-%Y')
            end_imap = end_date.strftime('%d-%b-%Y')
            
            self.logger.info(f"📅 Formato IMAP inicio: {start_imap}")
            self.logger.info(f"📅 Formato IMAP fin: {end_imap}")
            
            # Configurar criterios de búsqueda
            self.logger.info(f"📅 Buscando emails desde: {start_imap}")
            
            # Calcular duración del rango
            duration = (end_date - start_date).days + 1
            self.logger.info(f"📊 Duración del rango: {duration} días")
            
            if duration == 1:
                # Un solo día: usar SINCE y BEFORE del día siguiente
                next_day = end_date + timedelta(days=1)
                next_day_imap = next_day.strftime('%d-%b-%Y')
                search_criteria = f'SINCE {start_imap} BEFORE {next_day_imap}'
                self.logger.info(f"📅 Buscando emails hasta: {end_imap} (usando BEFORE {next_day_imap})")
            else:
                # Múltiples días: usar SINCE y BEFORE
                next_day = end_date + timedelta(days=1)
                next_day_imap = next_day.strftime('%d-%b-%Y')
                search_criteria = f'SINCE {start_imap} BEFORE {next_day_imap}'
                self.logger.info(f"📅 Buscando emails hasta: {end_imap} (usando BEFORE {next_day_imap})")
            
            self.logger.info(f"🔍 Criterio de búsqueda FINAL: '{search_criteria}'")
            
            try:
                status, messages = self.mail.search(None, search_criteria)
                self.logger.info(f"🔍 Resultado de búsqueda IMAP: {status}")
            except Exception as e:
                self.logger.error(f"❌ Error en búsqueda IMAP: {e}")
                return []
        
        if status != 'OK':
            self.logger.error(f"❌ Error en búsqueda: {status}")
            return []
        
        email_ids = messages[0].split()
        self.logger.info(f"📊 TOTAL de emails encontrados: {len(email_ids)}")
        
        if email_ids:
            # Obtener fechas del primer y último email para verificación
            try:
                first_email_result, first_email_data = self.mail.fetch(email_ids[0], '(RFC822)')
                if first_email_result == 'OK':
                    first_msg = email.message_from_bytes(first_email_data[0][1])
                    first_date = self.get_email_date(first_msg)
                    self.logger.info(f"📅 Fecha del primer email: {first_date.strftime('%Y-%m-%d %H:%M:%S')}")
                
                last_email_result, last_email_data = self.mail.fetch(email_ids[-1], '(RFC822)')
                if last_email_result == 'OK':
                    last_msg = email.message_from_bytes(last_email_data[0][1])
                    last_date = self.get_email_date(last_msg)
                    self.logger.info(f"📅 Fecha del último email: {last_date.strftime('%Y-%m-%d %H:%M:%S')}")
                    
            except Exception as e:
                self.logger.debug(f"Error obteniendo fechas de verificación: {e}")
        
        self.logger.info("=== BÚSQUEDA DE EMAILS COMPLETADA ===")
        return email_ids
    
    def check_email_matches_filters(self, msg, sender, subject):
        """Verifica si un email cumple con los filtros configurados"""
        filters = self.config['filters']
        rejection_reasons = []
        
        self.logger.debug(f"🔍 === VERIFICANDO FILTROS PARA: {sender} - {subject} ===")
        
        # 1. Verificar rango de fechas (validación adicional a nivel de email)
        if filters['date_range']['enabled']:
            start_date = filters['date_range'].get('start_date')
            end_date = filters['date_range'].get('end_date')
            
            if start_date and end_date:
                email_date = self.get_email_date(msg)
                
                # Verificar que la fecha del email esté en el rango
                if email_date.date() < start_date.date():
                    rejection_reasons.append(f"Email anterior al rango ({email_date.date()} < {start_date.date()})")
                    self.logger.debug(f"❌ FILTRO FECHA: Email anterior al rango")
                elif email_date.date() > end_date.date():
                    rejection_reasons.append(f"Email posterior al rango ({email_date.date()} > {end_date.date()})")
                    self.logger.debug(f"❌ FILTRO FECHA: Email posterior al rango")
                else:
                    self.logger.debug(f"✅ FILTRO FECHA: OK ({email_date.date()})")
        
        # 2. Verificar remitentes
        if filters['sender_emails'] and len(filters['sender_emails']) > 0:
            # HAY remitentes específicos configurados
            self.logger.debug(f"🔍 FILTRO REMITENTES: Verificando contra lista: {filters['sender_emails']}")
            sender_match = False
            for filter_sender in filters['sender_emails']:
                if filter_sender.lower() in sender.lower():
                    sender_match = True
                    self.logger.debug(f"✅ FILTRO REMITENTES: Match encontrado con {filter_sender}")
                    break
            
            if not sender_match:
                rejection_reasons.append(f"Remitente no coincide con los configurados ({', '.join(filters['sender_emails'])})")
                self.logger.debug(f"❌ FILTRO REMITENTES: No coincide con ninguno de la lista")
        else:
            # NO hay remitentes específicos = permitir cualquier remitente
            self.logger.debug(f"✅ FILTRO REMITENTES: Lista vacía - permitiendo cualquier remitente")
        
        # 3. Verificar palabras clave en asunto
        if filters['subject_keywords'] and len(filters['subject_keywords']) > 0:
            # HAY palabras clave configuradas
            self.logger.debug(f"🔍 FILTRO PALABRAS CLAVE: Verificando contra: {filters['subject_keywords']}")
            subject_match = False
            subject_normalized = self.normalize_text_for_search(subject)
            
            # Debug: mostrar asunto normalizado
            self.logger.debug(f"🔍 Asunto original: '{subject}'")
            self.logger.debug(f"🔍 Asunto normalizado: '{subject_normalized}'")
            
            for keyword in filters['subject_keywords']:
                keyword_normalized = self.normalize_text_for_search(keyword)
                self.logger.debug(f"🔍 Buscando palabra clave: '{keyword}' (normalizada: '{keyword_normalized}')")
                
                if keyword_normalized in subject_normalized:
                    subject_match = True
                    self.logger.debug(f"✅ FILTRO PALABRAS CLAVE: MATCH encontrado con palabra clave: '{keyword}'")
                    break
                else:
                    self.logger.debug(f"❌ No match con: '{keyword_normalized}' en '{subject_normalized}'")
            
            if not subject_match:
                rejection_reasons.append(f"Asunto no contiene palabras clave esperadas ({', '.join(filters['subject_keywords'])})")
                self.logger.debug(f"❌ FILTRO PALABRAS CLAVE: No se encontró ninguna coincidencia")
        else:
            # NO hay palabras clave = permitir cualquier asunto
            self.logger.debug(f"✅ FILTRO PALABRAS CLAVE: Lista vacía - permitiendo cualquier asunto")
        
        # 4. Verificar archivos adjuntos
        self.logger.debug(f"🔍 FILTRO ARCHIVOS ADJUNTOS: Verificando...")
        has_attachments = self.has_relevant_attachments(msg)
        if not has_attachments:
            rejection_reasons.append("Sin archivos adjuntos de tipos permitidos")
            self.logger.debug(f"❌ FILTRO ARCHIVOS ADJUNTOS: No se encontraron archivos válidos")
        else:
            self.logger.debug(f"✅ FILTRO ARCHIVOS ADJUNTOS: Archivos válidos encontrados")
        
        # Resultado final
        passes_all_filters = len(rejection_reasons) == 0
        
        if passes_all_filters:
            self.logger.debug(f"🎉 RESULTADO: EMAIL APROBADO - Pasa todos los filtros")
        else:
            self.logger.debug(f"💥 RESULTADO: EMAIL RECHAZADO - Motivos: {'; '.join(rejection_reasons)}")
        
        self.logger.debug(f"🔍 === FIN VERIFICACIÓN FILTROS ===")
        
        return passes_all_filters, rejection_reasons
    
    def has_relevant_attachments(self, msg):
        allowed_extensions = self.config['download_settings']['allowed_extensions']
        
        self.logger.debug(f"🔍 ANÁLISIS EXTREMO - Extensiones permitidas: {allowed_extensions}")
        
        attachment_found = False
        attachment_details = []
        html_content = None
        
        # DEBUG: Mostrar estructura completa del mensaje
        self.logger.debug(f"🏗️ ESTRUCTURA DEL MENSAJE:")
        self.logger.debug(f"   Es multipart: {msg.is_multipart()}")
        self.logger.debug(f"   Content-Type principal: {msg.get_content_type()}")
        self.logger.debug(f"   Headers principales: {list(msg.keys())}")
        
        # ANÁLISIS PARTE POR PARTE CON DEBUG EXTREMO
        part_index = 0
        for part in msg.walk():
            self.logger.debug(f"\n🔎 ═══ ANÁLISIS DETALLADO PARTE {part_index} ═══")
            
            # Obtener TODOS los headers de esta parte
            filename = part.get_filename()
            content_type = part.get_content_type()
            content_disposition = part.get('Content-Disposition', '')
            content_id = part.get('Content-ID', '')
            content_transfer_encoding = part.get('Content-Transfer-Encoding', '')
            
            # DEBUG: Mostrar TODOS los headers de esta parte
            self.logger.debug(f"   📋 HEADERS COMPLETOS:")
            for header_name, header_value in part.items():
                self.logger.debug(f"      {header_name}: {header_value}")
            
            self.logger.debug(f"   📊 ANÁLISIS BÁSICO:")
            self.logger.debug(f"      Content-Type: {content_type}")
            self.logger.debug(f"      Content-Disposition: {content_disposition}")
            self.logger.debug(f"      Content-Transfer-Encoding: {content_transfer_encoding}")
            self.logger.debug(f"      Filename: {filename}")
            self.logger.debug(f"      Content-ID: {content_id}")
            
            # Guardar contenido HTML para análisis posterior
            if content_type == 'text/html' and not html_content:
                try:
                    html_content = part.get_payload(decode=True).decode('utf-8', errors='replace')
                    self.logger.debug(f"   📝 CONTENIDO HTML OBTENIDO ({len(html_content)} caracteres)")
                except Exception as e:
                    self.logger.debug(f"   ❌ Error obteniendo HTML: {e}")
            
            # Verificar si tiene payload
            try:
                payload = part.get_payload(decode=False)
                has_payload = payload is not None
                payload_type = type(payload).__name__
                
                if isinstance(payload, str):
                    payload_length = len(payload)
                    payload_preview = payload[:100].replace('\n', '\\n').replace('\r', '\\r')
                elif isinstance(payload, list):
                    payload_length = len(payload)
                    payload_preview = f"Lista con {payload_length} elementos"
                else:
                    payload_length = "unknown"
                    payload_preview = str(payload)[:100]
                
                self.logger.debug(f"      Payload: {has_payload} (tipo: {payload_type})")
                self.logger.debug(f"      Payload length: {payload_length}")
                self.logger.debug(f"      Payload preview: {payload_preview}")
                
            except Exception as e:
                self.logger.debug(f"      Error obteniendo payload: {e}")
            
            # Intentar decodificar payload
            try:
                decoded_payload = part.get_payload(decode=True)
                if decoded_payload:
                    decoded_length = len(decoded_payload)
                    self.logger.debug(f"      Payload decodificado: {decoded_length} bytes")
                    
                    # Verificar magic bytes SOLO si es binario
                    if decoded_length > 10 and isinstance(decoded_payload, bytes):
                        magic_bytes = decoded_payload[:10]
                        magic_hex = magic_bytes.hex()
                        self.logger.debug(f"      Magic bytes: {magic_hex}")
                        
                        # Detectar tipo de archivo por magic bytes
                        if magic_bytes.startswith(b'\xff\xd8\xff'):
                            self.logger.debug(f"      🎯 JPEG DETECTADO POR MAGIC BYTES!")
                            attachment_details.append(f"JPEG por magic bytes en parte {part_index}")
                            if '.jpg' in allowed_extensions or '.jpeg' in allowed_extensions:
                                attachment_found = True
                        elif magic_bytes.startswith(b'\x89PNG'):
                            self.logger.debug(f"      🎯 PNG DETECTADO POR MAGIC BYTES!")
                            attachment_details.append(f"PNG por magic bytes en parte {part_index}")
                            if '.png' in allowed_extensions:
                                attachment_found = True
                        elif magic_bytes.startswith(b'%PDF'):
                            self.logger.debug(f"      🎯 PDF DETECTADO POR MAGIC BYTES!")
                            attachment_details.append(f"PDF por magic bytes en parte {part_index}")
                            if '.pdf' in allowed_extensions:
                                attachment_found = True
                        elif magic_bytes.startswith(b'GIF8'):
                            self.logger.debug(f"      🎯 GIF DETECTADO POR MAGIC BYTES!")
                            attachment_details.append(f"GIF por magic bytes en parte {part_index}")
                            if '.gif' in allowed_extensions:
                                attachment_found = True
                        elif magic_bytes.startswith(b'DICM'):
                            self.logger.debug(f"      🏥 DICOM DETECTADO POR MAGIC BYTES!")
                            attachment_details.append(f"DICOM por magic bytes en parte {part_index}")
                            if '.dcm' in allowed_extensions:
                                attachment_found = True
                else:
                    self.logger.debug(f"      Payload decodificado: None o vacío")
                    
            except Exception as e:
                self.logger.debug(f"      Error decodificando payload: {e}")
            
            # VERIFICACIONES ESPECÍFICAS
            
            # 1. Archivo con filename
            if filename:
                try:
                    decoded_filename = self.decode_email_header(filename)
                    file_ext = Path(decoded_filename).suffix.lower()
                    self.logger.debug(f"      🔍 FILENAME ANALYSIS:")
                    self.logger.debug(f"         Original: {filename}")
                    self.logger.debug(f"         Decoded: {decoded_filename}")
                    self.logger.debug(f"         Extension: {file_ext}")
                    
                    if file_ext in allowed_extensions:
                        self.logger.debug(f"      ✅ MATCH POR FILENAME: {decoded_filename}")
                        attachment_details.append(f"Archivo válido: {decoded_filename}")
                        attachment_found = True
                    else:
                        self.logger.debug(f"      ❌ Extensión no permitida: {file_ext}")
                except Exception as e:
                    self.logger.debug(f"      Error procesando filename: {e}")
            
            # 2. Content-Type específico
            if content_type:
                self.logger.debug(f"      🔍 CONTENT-TYPE ANALYSIS: {content_type}")
                
                # Mapeo detallado
                type_matches = {
                    'image/jpeg': ['.jpg', '.jpeg'],
                    'image/jpg': ['.jpg'],
                    'image/png': ['.png'],
                    'image/gif': ['.gif'],
                    'image/bmp': ['.bmp'],
                    'application/pdf': ['.pdf'],
                    'application/dicom': ['.dcm'],
                    'application/octet-stream': ['.jpg', '.jpeg', '.png', '.pdf', '.dcm']  # Genérico
                }
                
                if content_type in type_matches:
                    possible_exts = type_matches[content_type]
                    matching_exts = [ext for ext in possible_exts if ext in allowed_extensions]
                    
                    if matching_exts:
                        self.logger.debug(f"      ✅ MATCH POR CONTENT-TYPE: {content_type} -> {matching_exts}")
                        attachment_details.append(f"Content-type válido: {content_type}")
                        attachment_found = True
                    else:
                        self.logger.debug(f"      ❌ Content-type no tiene extensiones permitidas")
            
            # 3. Content-Disposition
            if content_disposition:
                self.logger.debug(f"      🔍 CONTENT-DISPOSITION ANALYSIS: {content_disposition}")
                if 'attachment' in content_disposition.lower():
                    self.logger.debug(f"      ✅ MATCH POR CONTENT-DISPOSITION: attachment")
                    attachment_details.append(f"Marcado como attachment")
                    attachment_found = True
            
            part_index += 1
            self.logger.debug(f"═══ FIN PARTE {part_index-1} ═══\n")
        
        # ANÁLISIS FINAL
        self.logger.debug(f"📊 RESUMEN FINAL:")
        self.logger.debug(f"   Total partes analizadas: {part_index}")
        self.logger.debug(f"   Detalles encontrados: {attachment_details}")
        self.logger.debug(f"   ¿Encontró archivos relevantes? {'SÍ' if attachment_found else 'NO'}")
        
        # Si no encontró nada, hacer análisis adicional
        if not attachment_found:
            self.logger.debug(f"🔍 ÚLTIMO INTENTO - Verificando estructura completa...")
            
            # 1. Verificar si hay partes con contenido binario significativo
            for i, part in enumerate(msg.walk()):
                try:
                    payload = part.get_payload(decode=True)
                    if payload and len(payload) > 1000:  # Archivos > 1KB
                        self.logger.debug(f"   Parte {i}: {len(payload)} bytes de contenido binario")
                        
                        # Si es binario grande y no es texto, probablemente es un archivo
                        if isinstance(payload, bytes):
                            # Verificar si no es solo texto
                            try:
                                # Intento de decodificar como texto
                                payload.decode('utf-8')
                                self.logger.debug(f"   Parte {i}: Es texto UTF-8")
                            except UnicodeDecodeError:
                                self.logger.debug(f"   Parte {i}: Es contenido binario - ¡POSIBLE ARCHIVO!")
                                attachment_details.append(f"Contenido binario grande en parte {i}")
                                
                                # Verificar magic bytes si no se hizo antes
                                if len(payload) > 10:
                                    magic_bytes = payload[:10]
                                    if magic_bytes.startswith(b'\xff\xd8\xff'):
                                        self.logger.debug(f"   🎯 ¡JPEG DETECTADO EN BINARIO!")
                                        if '.jpg' in allowed_extensions or '.jpeg' in allowed_extensions:
                                            attachment_found = True
                                    elif magic_bytes.startswith(b'\x89PNG'):
                                        self.logger.debug(f"   🎯 ¡PNG DETECTADO EN BINARIO!")
                                        if '.png' in allowed_extensions:
                                            attachment_found = True
                                    else:
                                        # Considerar como archivo válido genérico
                                        attachment_found = True
                except:
                    pass
            
            # 2. Analizar contenido HTML para detectar imágenes embebidas y enlaces
            if html_content:
                self.logger.debug(f"🔍 ANALIZANDO CONTENIDO HTML ({len(html_content)} caracteres)")
                
                # Patrones mejorados para detectar contenido
                patterns = {
                    'base64_images': r'src="data:image/(\w+);base64,([a-zA-Z0-9+/=]+)"',
                    'image_links': r'https?://[^\s<>"\']+\.(jpg|jpeg|png|gif|bmp|tiff|tif|webp|dcm|pdf)\b',
                    'cid_references': r'src=["\']cid:([^"\'\s]+)["\']',
                    'img_tags': r'<img[^>]+src=["\']([^"\']+)["\']'
                }
                
                # Buscar imágenes en base64
                base64_images = re.findall(patterns['base64_images'], html_content, re.IGNORECASE)
                if base64_images:
                    self.logger.debug(f"   🖼️ Encontradas {len(base64_images)} imágenes base64 en HTML")
                    attachment_details.append(f"Imágenes base64 en HTML")
                    attachment_found = True
                
                # Buscar enlaces a imágenes/archivos (versión mejorada)
                image_links = re.findall(patterns['image_links'], html_content, re.IGNORECASE)
                if image_links:
                    self.logger.debug(f"   🔗 Encontrados {len(image_links)} enlaces a archivos en HTML")
                    attachment_details.append(f"Enlaces a archivos en HTML")
                    attachment_found = True
                
                # Buscar referencias CID (archivos embebidos)
                cid_references = re.findall(patterns['cid_references'], html_content, re.IGNORECASE)
                if cid_references:
                    self.logger.debug(f"   📎 Encontradas {len(cid_references)} referencias CID en HTML")
                    attachment_details.append(f"Referencias CID en HTML")
                    attachment_found = True
                
                # Buscar cualquier etiqueta img (útil para imágenes sin extensión)
                img_tags = re.findall(patterns['img_tags'], html_content, re.IGNORECASE)
                if img_tags:
                    self.logger.debug(f"   🖼️ Encontradas {len(img_tags)} etiquetas img en HTML")
                    for img_src in img_tags:
                        if img_src.startswith('http') and any(ext in img_src for ext in allowed_extensions):
                            attachment_details.append(f"Imagen en etiqueta img: {img_src[:50]}...")
                            attachment_found = True
        
        return attachment_found
    
    def extract_image_links_from_content(self, content):
        """Extrae enlaces de imágenes del contenido HTML/texto del email"""
        image_links = []
        allowed_extensions = self.config['download_settings']['allowed_extensions']
        
        # Patrones para encontrar URLs de imágenes
        patterns = [
            r'https?://[^\s<>"\']+\.(?:jpg|jpeg|png|gif|bmp|tiff|tif|webp|dcm|pdf)\b',  # URLs directas
            r'src=["\']([^"\']+\.(?:jpg|jpeg|png|gif|bmp|tiff|tif|webp))["\']',  # Atributos src
            r'href=["\']([^"\']+\.(?:jpg|jpeg|png|gif|bmp|tiff|tif|webp|dcm|pdf))["\']',  # Enlaces
            r'cid:([^<>\s"\']+)',  # Content-ID references para imágenes embebidas
        ]
        
        import re
        for pattern in patterns:
            matches = re.findall(pattern, content, re.IGNORECASE)
            for match in matches:
                # Limpiar y validar URL
                url = match.strip()
                if url.startswith('cid:'):
                    # Content-ID reference - imagen embebida
                    image_links.append({'type': 'cid', 'url': url, 'source': 'embedded'})
                else:
                    # URL externa - verificar extensión
                    for ext in allowed_extensions:
                        if url.lower().endswith(ext.lower()):
                            image_links.append({'type': 'url', 'url': url, 'source': 'link'})
                            break
        
        return image_links
    
    def analyze_email_for_report(self, email_id):
        """
        VERSIÓN CON DEBUG EXTREMO - Para el caso específico de Patricia
        """
        try:
            result, msg_data = self.mail.fetch(email_id, '(RFC822)')
            
            if result != 'OK':
                self.add_email_to_report(email_id, None, "ERROR", 0, "Error obteniendo datos del email", "N/A")
                return False
            
            msg = email.message_from_bytes(msg_data[0][1])
            sender = self.decode_email_header(msg['From'])
            subject = self.decode_email_header(msg['Subject'])
            
            self.logger.debug(f"📧 ANÁLISIS DETALLADO: {sender} - {subject}")
            
            # DEBUG ESPECIAL PARA PATRICIA
            if 'ferrari_patricia@yahoo.com' in sender.lower():
                self.logger.debug(f"🚨 EMAIL DE PATRICIA DETECTADO - ANÁLISIS SÚPER DETALLADO:")
                self.logger.debug(f"   Full sender: {sender}")
                self.logger.debug(f"   Full subject: {subject}")
                
                # Mostrar TODOS los headers del email
                self.logger.debug(f"   📋 TODOS LOS HEADERS DEL EMAIL:")
                for header_name, header_value in msg.items():
                    self.logger.debug(f"      {header_name}: {header_value}")
            
            # Verificar filtros con logging extra para Patricia
            passes_filters, rejection_reasons = self.check_email_matches_filters(msg, sender, subject)
            
            if passes_filters:
                self.logger.info(f"✅ EMAIL APROBADO: {sender} - {subject}")
                self.add_email_to_report(email_id, msg, "PENDIENTE_DESCARGA", 0, "", "")
                return True
            else:
                motivo_completo = "; ".join(rejection_reasons)
                self.logger.warning(f"❌ EMAIL RECHAZADO: {sender} - {subject}")
                self.logger.warning(f"   Motivos: {motivo_completo}")
                
                # DEBUG EXTRA PARA PATRICIA
                if 'ferrari_patricia@yahoo.com' in sender.lower():
                    self.logger.debug(f"🚨 PATRICIA RECHAZADA - INVESTIGANDO...")
                    
                    # Forzar re-análisis de adjuntos solo para debug
                    self.logger.debug(f"   🔍 RE-ANALIZANDO ADJUNTOS...")
                    has_attachments = self.has_relevant_attachments(msg)
                    self.logger.debug(f"   Resultado re-análisis: {has_attachments}")
                
                self.add_email_to_report(email_id, msg, "DESCARTADO", 0, motivo_completo, "N/A")
                return False
                    
        except Exception as e:
            self.logger.error(f"Error analizando email {email_id}: {e}")
            self.add_email_to_report(email_id, None, "ERROR", 0, f"Error: {str(e)}", "N/A")
            return False    




    def add_email_to_report(self, email_id, msg, estado, archivos_descargados, motivo_rechazo, ruta_descarga):
        """Agrega un email al reporte CSV"""
        if msg:
            sender = self.decode_email_header(msg['From'])
            subject = self.decode_email_header(msg['Subject'])
            email_date = self.get_email_date(msg)
            
            # Contar archivos adjuntos
            total_attachments = 0
            attachment_types = []
            
            for part in msg.walk():
                filename = part.get_filename()
                content_type = part.get_content_type()
                
                if filename or content_type.startswith('image/') or content_type == 'application/pdf':
                    total_attachments += 1
                    if content_type not in attachment_types:
                        attachment_types.append(content_type)
        else:
            sender = "Error"
            subject = "Error obteniendo email"
            email_date = datetime.now()
            total_attachments = 0
            attachment_types = []
        
        self.report_data.append({
            'email_id': int(email_id),
            'fecha': email_date.strftime('%Y-%m-%d %H:%M:%S'),
            'remitente': sender,
            'asunto': subject,
            'archivos_adjuntos_total': total_attachments,
            'tipos_archivos': '; '.join(attachment_types) if attachment_types else 'Ninguno',
            'estado': estado,
            'archivos_descargados': archivos_descargados,
            'motivo_rechazo': motivo_rechazo,
            'ruta_descarga': ruta_descarga
        })
    
    def update_email_report_status(self, email_id, new_status, files_downloaded, download_path):
        """Actualiza el estado de un email en el reporte"""
        for entry in self.report_data:
            if entry['email_id'] == int(email_id):
                entry['estado'] = new_status
                entry['archivos_descargados'] = files_downloaded
                entry['ruta_descarga'] = download_path
                break
    
    def generate_report_csv(self):
        """Genera el reporte CSV con todos los emails analizados"""
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        report_filename = f'reporte_analisis_emails_{timestamp}.csv'
        
        with open(report_filename, 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = ['email_id', 'fecha', 'remitente', 'asunto', 'archivos_adjuntos_total', 
                         'tipos_archivos', 'estado', 'archivos_descargados', 'motivo_rechazo', 'ruta_descarga']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            
            writer.writeheader()
            for row in self.report_data:
                writer.writerow(row)
        
        self.logger.info(f"📊 REPORTE GENERADO: {report_filename}")
        return report_filename
    
    def create_folder_structure(self, msg, sender):
        """Crea la estructura de carpetas para organizar los archivos"""
        base_path = Path(self.config['download_settings']['base_folder'])
        
        # Obtener fecha del email
        email_date = self.get_email_date(msg)
        
        # Limpiar nombre del remitente para usar como carpeta
        clean_sender = re.sub(r'[<>:"/\\|?*]', '_', sender)
        if len(clean_sender) > 50:
            clean_sender = clean_sender[:50]
        
        # Estructura: base_folder/año/mes/día/remitente
        folder_path = base_path / str(email_date.year) / f"{email_date.month:02d}" / f"{email_date.day:02d}" / clean_sender
        folder_path.mkdir(parents=True, exist_ok=True)
        
        return folder_path, email_date
    
    def generate_filename(self, msg, sender, subject, index, original_filename, email_date):
        """Genera un nombre de archivo único basado en configuraciones"""
        try:
            # Limpiar caracteres especiales del asunto
            clean_subject = re.sub(r'[<>:"/\\|?*]', '_', subject)[:30]
            
            # Timestamp
            timestamp = email_date.strftime('%Y%m%d_%H%M%S')
            
            # Crear nombre base
            filename_parts = []
            
            if self.config['download_settings'].get('include_timestamp', True):
                filename_parts.append(timestamp)
            
            if self.config['download_settings'].get('include_subject', True) and clean_subject:
                filename_parts.append(clean_subject)
            
            if self.config['download_settings'].get('include_index', True):
                filename_parts.append(f"_{index:03d}")
            
            if not filename_parts:
                filename_parts = [timestamp, f"_{index:03d}"]
            
            base_name = "_".join(filename_parts)
            
            # Obtener extensión del archivo original
            original_ext = Path(original_filename).suffix if original_filename else ''
            
            return f"{base_name}{original_ext}"
            
        except Exception as e:
            self.logger.debug(f"Error generando nombre de archivo: {e}")
            timestamp = email_date.strftime('%Y%m%d_%H%M%S')
            return f"{timestamp}_{index:03d}"
    
    def is_duplicate_in_day(self, file_path, file_data, email_date):
        """Verifica si un archivo es duplicado dentro del mismo día"""
        if not self.config['download_settings'].get('skip_duplicates', True):
            return False
        
        # Crear hash del contenido del archivo
        file_hash = hashlib.md5(file_data).hexdigest()
        
        # Clave para el cache basada en fecha
        date_key = email_date.strftime('%Y-%m-%d')
        
        if date_key not in self.duplicates_cache:
            self.duplicates_cache[date_key] = set()
        
        if file_hash in self.duplicates_cache[date_key]:
            return True
        
        self.duplicates_cache[date_key].add(file_hash)
        return False
    
    def extract_google_drive_links(self, html_content):
        """Extrae enlaces de Google Drive del contenido HTML"""
        drive_patterns = [
            r'https://drive\.google\.com/file/d/([a-zA-Z0-9_-]+)',
            r'https://docs\.google\.com/document/d/([a-zA-Z0-9_-]+)',
            r'https://docs\.google\.com/spreadsheets/d/([a-zA-Z0-9_-]+)',
            r'https://docs\.google\.com/presentation/d/([a-zA-Z0-9_-]+)'
        ]
        
        drive_links = []
        for pattern in drive_patterns:
            matches = re.findall(pattern, html_content)
            for match in matches:
                drive_links.append({
                    'file_id': match,
                    'type': 'google_drive',
                    'download_url': f'https://drive.google.com/uc?export=download&id={match}'
                })
        
        return drive_links
    
    def download_from_google_drive(self, drive_info, target_folder, base_filename, index, email_date):
        """Descarga archivo desde Google Drive"""
        try:
            import requests
            
            url = drive_info['download_url']
            self.logger.debug(f"☁️ Descargando desde Google Drive: {url}")
            
            session = requests.Session()
            response = session.get(url, stream=True)
            
            # Google Drive a veces requiere confirmación para archivos grandes
            if 'virus scan warning' in response.text.lower():
                # Buscar el enlace de confirmación
                confirm_pattern = r'confirm=([0-9A-Za-z_]+)'
                confirm_match = re.search(confirm_pattern, response.text)
                if confirm_match:
                    confirm_token = confirm_match.group(1)
                    confirm_url = f"{url}&confirm={confirm_token}"
                    response = session.get(confirm_url, stream=True)
            
            response.raise_for_status()
            
            # Determinar extensión del archivo
            content_disposition = response.headers.get('Content-Disposition', '')
            if 'filename=' in content_disposition:
                filename_match = re.search(r'filename="?([^"]+)"?', content_disposition)
                if filename_match:
                    original_filename = filename_match.group(1)
                    file_ext = Path(original_filename).suffix
                else:
                    file_ext = '.bin'
            else:
                content_type = response.headers.get('Content-Type', '')
                if 'pdf' in content_type:
                    file_ext = '.pdf'
                elif 'image' in content_type:
                    file_ext = '.jpg'
                else:
                    file_ext = '.bin'
            
            # Generar nombre de archivo
            filename = f"{base_filename}_drive_{index}{file_ext}"
            file_path = target_folder / filename
            
            # Descargar contenido
            file_data = b''
            for chunk in response.iter_content(chunk_size=8192):
                if chunk:
                    file_data += chunk
            
            if len(file_data) == 0:
                return None
            
            # Verificar duplicados
            if self.is_duplicate_in_day(file_path, file_data, email_date):
                self.logger.info(f"⏭️ Duplicado omitido: {filename}")
                return None
            
            # Resolver conflictos de nombres
            counter = 1
            original_path = file_path
            while file_path.exists():
                stem = original_path.stem
                suffix = original_path.suffix
                file_path = original_path.parent / f"{stem}_{counter}{suffix}"
                counter += 1
            
            # Guardar archivo
            file_path.write_bytes(file_data)
            
            if file_path.exists() and file_path.stat().st_size > 0:
                self.logger.info(f"✅ Descargado desde Google Drive: {file_path}")
                return file_path
            else:
                if file_path.exists():
                    file_path.unlink()
                return None
                
        except Exception as e:
            self.logger.error(f"Error descargando desde Google Drive: {e}")
            return None
    
    def download_from_image_link(self, link_info, target_folder, sender, subject, index, email_date):
        """Descarga imagen desde un enlace encontrado en el contenido del email"""
        try:
            # Saltar Content-ID references por ahora (requieren procesamiento especial)
            if link_info['type'] == 'cid':
                self.logger.debug(f"⏭️ Saltando Content-ID reference: {link_info['url']}")
                return None
            
            # Descargar desde URL
            url = link_info['url']
            self.logger.debug(f"🔗 Descargando imagen desde URL: {url}")
            
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
            }
            
            response = requests.get(url, headers=headers, timeout=30, stream=True)
            response.raise_for_status()
            
            # Determinar extensión desde URL o Content-Type
            file_ext = None
            if '.' in url:
                file_ext = '.' + url.split('.')[-1].split('?')[0].lower()
            
            if not file_ext or file_ext not in self.config['download_settings']['allowed_extensions']:
                content_type = response.headers.get('Content-Type', '')
                if 'image/jpeg' in content_type:
                    file_ext = '.jpg'
                elif 'image/png' in content_type:
                    file_ext = '.png'
                elif 'image/gif' in content_type:
                    file_ext = '.gif'
                elif 'application/pdf' in content_type:
                    file_ext = '.pdf'
                else:
                    file_ext = '.jpg'  # Default
            
            # Generar nombre de archivo
            if self.config['download_settings']['rename_files']:
                base_filename = self.generate_filename(
                    None, sender, subject, index, f"imagen_link", email_date
                ).replace("_imagen_link", "")
                filename = f"{base_filename}_link_{index}{file_ext}"
            else:
                filename = f"imagen_link_{index}{file_ext}"
            
            file_path = target_folder / filename
            
            # Descargar contenido
            file_data = b''
            for chunk in response.iter_content(chunk_size=8192):
                if chunk:
                    file_data += chunk
            
            if len(file_data) == 0:
                return None
            
            # Verificar duplicados
            if self.is_duplicate_in_day(file_path, file_data, email_date):
                self.logger.info(f"⏭️ Duplicado omitido: {filename}")
                return None
            
            # Resolver conflictos de nombres
            counter = 1
            original_path = file_path
            while file_path.exists():
                stem = original_path.stem
                suffix = original_path.suffix
                file_path = original_path.parent / f"{stem}_{counter}{suffix}"
                counter += 1
            
            # Guardar archivo
            file_path.write_bytes(file_data)
            
            if file_path.exists() and file_path.stat().st_size > 0:
                self.logger.info(f"✅ Descargado desde enlace: {file_path}")
                return file_path
            else:
                if file_path.exists():
                    file_path.unlink()
                return None
                
        except Exception as e:
            self.logger.error(f"Error descargando imagen desde enlace {link_info.get('url', 'unknown')}: {e}")
            return None
    
    def download_images_from_email(self, email_id):
        """
        Versión MEJORADA: Descarga archivos usando detección más agresiva
        """
        try:
            result, msg_data = self.mail.fetch(email_id, '(RFC822)')
            
            if result != 'OK':
                self.update_email_report_status(email_id, "ERROR", 0, "Error obteniendo email para descarga")
                return 0
            
            msg = email.message_from_bytes(msg_data[0][1])
            sender = self.decode_email_header(msg['From'])
            subject = self.decode_email_header(msg['Subject'])
            
            self.logger.info(f"🔍 DESCARGA MEJORADA: De: {sender}, Asunto: {subject}")
            
            # Usar la detección mejorada
            if not self.has_relevant_attachments(msg):
                self.logger.debug(f"❌ Sin archivos relevantes - De: {sender}, Asunto: {subject}")
                self.update_email_report_status(email_id, "SIN_ARCHIVOS", 0, "N/A")
                return 0
            
            self.logger.info(f"✅ PROCESANDO CON DETECCIÓN MEJORADA: De: {sender}, Asunto: {subject}")
            
            email_date = self.get_email_date(msg)
            downloaded_count = 0
            
            # EXTRACCIÓN AGRESIVA DE ARCHIVOS
            attachments_to_download = []
            allowed_extensions = self.config['download_settings']['allowed_extensions']
            
            for i, part in enumerate(msg.walk()):
                filename = part.get_filename()
                content_type = part.get_content_type()
                content_disposition = part.get('Content-Disposition', '')
                
                should_download = False
                generated_filename = None
                file_data = None
                
                # Obtener datos del archivo
                try:
                    file_data = part.get_payload(decode=True)
                    if not file_data or len(file_data) < 10:  # Muy pequeño
                        continue
                except:
                    continue
                
                # ESTRATEGIA 1: Archivo con filename
                if filename:
                    try:
                        decoded_filename = self.decode_email_header(filename)
                        file_ext = Path(decoded_filename).suffix.lower()
                        
                        if file_ext in allowed_extensions:
                            should_download = True
                            generated_filename = decoded_filename
                            self.logger.debug(f"✅ ARCHIVO ENCONTRADO (filename): {decoded_filename}")
                    except Exception as e:
                        self.logger.debug(f"Error con filename: {e}")
                
                # ESTRATEGIA 2: Sin filename pero content-type conocido
                elif content_type and not should_download:
                    ext_map = {
                        'image/jpeg': '.jpg',
                        'image/jpg': '.jpg',
                        'image/png': '.png',
                        'image/gif': '.gif',
                        'image/bmp': '.bmp',
                        'image/tiff': '.tiff',
                        'image/webp': '.webp',
                        'application/pdf': '.pdf',
                        'application/msword': '.doc',
                        'application/vnd.openxmlformats-officedocument.wordprocessingml.document': '.docx'
                    }
                    
                    if content_type in ext_map:
                        file_ext = ext_map[content_type]
                        if file_ext in allowed_extensions:
                            should_download = True
                            generated_filename = f"archivo_parte_{i}{file_ext}"
                            self.logger.debug(f"✅ ARCHIVO ENCONTRADO (content-type): {content_type} -> {generated_filename}")
                
                # ESTRATEGIA 3: Detección por magic bytes
                elif not should_download and file_data:
                    detected_ext = None
                    
                    if file_data.startswith(b'\xff\xd8\xff'):  # JPEG
                        detected_ext = '.jpg'
                    elif file_data.startswith(b'\x89PNG'):  # PNG
                        detected_ext = '.png'
                    elif file_data.startswith(b'%PDF'):  # PDF
                        detected_ext = '.pdf'
                    elif file_data.startswith(b'GIF8'):  # GIF
                        detected_ext = '.gif'
                    
                    if detected_ext and detected_ext in allowed_extensions:
                        should_download = True
                        generated_filename = f"archivo_magic_{i}{detected_ext}"
                        self.logger.debug(f"✅ ARCHIVO ENCONTRADO (magic bytes): {detected_ext} -> {generated_filename}")
                
                # ESTRATEGIA 4: Si es attachment, intentar descargar
                elif 'attachment' in content_disposition.lower() and not should_download:
                    # Asumir extensión basada en content-type o usar .bin
                    if 'image' in content_type:
                        detected_ext = '.jpg'  # Default para imágenes
                    elif 'pdf' in content_type:
                        detected_ext = '.pdf'
                    else:
                        detected_ext = '.bin'
                    
                    if detected_ext in allowed_extensions:
                        should_download = True
                        generated_filename = f"attachment_{i}{detected_ext}"
                        self.logger.debug(f"✅ ARCHIVO ENCONTRADO (attachment): {generated_filename}")
                
                # Agregar a la lista de descarga
                if should_download and generated_filename and file_data:
                    attachments_to_download.append({
                        'filename': generated_filename,
                        'data': file_data,
                        'content_type': content_type,
                        'source': f'estrategia_mejorada_parte_{i}'
                    })
                    self.logger.info(f"📎 AGREGADO PARA DESCARGA: {generated_filename} ({len(file_data)} bytes)")
            
            # Descargar archivos encontrados
            if attachments_to_download:
                target_folder, email_date = self.create_folder_structure(msg, sender)
                download_path = str(target_folder)
                
                self.logger.info(f"📁 Descargando {len(attachments_to_download)} archivos en: {target_folder}")
                
                for idx, attachment in enumerate(attachments_to_download):
                    try:
                        if self.config['download_settings']['rename_files']:
                            new_filename = self.generate_filename(
                                msg, sender, subject, idx, attachment['filename'], email_date
                            )
                        else:
                            new_filename = attachment['filename']
                        
                        file_path = target_folder / new_filename
                        
                        # Resolver conflictos de nombres
                        counter = 1
                        original_path = file_path
                        while file_path.exists():
                            stem = original_path.stem
                            suffix = original_path.suffix
                            file_path = original_path.parent / f"{stem}_{counter}{suffix}"
                            counter += 1
                        
                        file_path.write_bytes(attachment['data'])
                        downloaded_count += 1
                        self.logger.info(f"✅ DESCARGADO: {file_path}")
                        
                    except Exception as e:
                        self.logger.error(f"❌ Error descargando {attachment['filename']}: {e}")
                
                self.update_email_report_status(email_id, "DESCARGADO", downloaded_count, download_path)
            else:
                self.logger.warning(f"⚠️ No se pudieron extraer archivos de: {sender} - {subject}")
                self.update_email_report_status(email_id, "SIN_ARCHIVOS", 0, "N/A")
            
            return downloaded_count
            
        except Exception as e:
            self.logger.error(f"❌ Error procesando email {email_id}: {e}")
            self.update_email_report_status(email_id, "ERROR", 0, f"Error: {str(e)}")
            return 0
    
    def run(self):
        """Método principal para compatibilidad con la interfaz - llama a run_complete_analysis"""
        return self.run_complete_analysis()
    
    def run_complete_analysis(self):
        """Ejecuta el análisis completo: busca emails, filtra y genera reporte"""
        self.logger.info("=== INICIANDO ANÁLISIS COMPLETO DE EMAILS CON SISTEMA DE FECHAS MEJORADO ===")
        
        # Mostrar configuración de fechas
        filters = self.config['filters']
        if filters['date_range']['enabled']:
            start_date = filters['date_range'].get('start_date')
            end_date = filters['date_range'].get('end_date')
            duration = (end_date - start_date).days + 1
            
            self.logger.info("📅 RANGO DE FECHAS CONFIGURADO:")
            self.logger.info(f"  • Desde: {start_date.strftime('%d/%m/%Y %H:%M:%S')} ")
            self.logger.info(f"  • Hasta: {end_date.strftime('%d/%m/%Y %H:%M:%S')} ")
            self.logger.info(f"  • Duración: {duration} días")
        
        # Conectar al email
        if not self.connect_to_email():
            return None
        
        try:
            # PASO 1: Buscar emails en el rango de fechas
            self.logger.info("🔍 PASO 1: Buscando TODOS los emails en el rango configurado...")
            email_ids = self.search_emails_by_date_range()
            
            if not email_ids:
                self.logger.warning("⚠️ No se encontraron emails en el rango de fechas especificado")
                return None
            
            self.logger.info(f"📊 Se analizarán {len(email_ids)} emails en total")
            
            # PASO 2: Analizar cada email contra los filtros
            self.logger.info("📋 PASO 2: Analizando cada email contra los filtros configurados...")
            
            valid_emails = []
            
            for email_id in email_ids:
                if self.analyze_email_for_report(email_id):
                    valid_emails.append(email_id)
            
            self.logger.info("✅ PASO 2 COMPLETADO:")
            self.logger.info(f"  • Total emails analizados: {len(email_ids)}")
            self.logger.info(f"  • Emails que cumplen criterios: {len(valid_emails)}")
            self.logger.info(f"  • Emails descartados: {len(email_ids) - len(valid_emails)}")
            
            # PASO 3: Descargar archivos de emails válidos
            if valid_emails:
                self.logger.info(f"📥 PASO 3: Descargando archivos de {len(valid_emails)} emails válidos...")
                
                total_files_downloaded = 0
                
                for email_id in valid_emails:
                    try:
                        files_downloaded = self.download_images_from_email(email_id)
                        total_files_downloaded += files_downloaded
                    except Exception as e:
                        self.logger.error(f"Error descargando email {email_id}: {e}")
                
                self.logger.info("✅ PASO 3 COMPLETADO:")
                self.logger.info(f"  • Total archivos descargados: {total_files_downloaded}")
            else:
                self.logger.info("⚠️ No hay emails que cumplan los criterios para descarga")
            
            # PASO 4: Generar reporte
            self.logger.info("📊 PASO 4: Generando reporte detallado...")
            report_filename = self.generate_report_csv()
            
            # Estadísticas finales
            descargados = len([email for email in self.report_data if email['estado'] == 'DESCARGADO'])
            sin_archivos = len([email for email in self.report_data if email['estado'] == 'SIN_ARCHIVOS'])
            descartados = len([email for email in self.report_data if email['estado'] == 'DESCARTADO'])
            errores = len([email for email in self.report_data if email['estado'] == 'ERROR'])
            total_archivos = sum([email['archivos_descargados'] for email in self.report_data])
            
            self.logger.info(f"  • Total emails analizados: {len(self.report_data)}")
            self.logger.info(f"  • Emails con archivos descargados: {descargados}")
            self.logger.info(f"  • Emails descartados: {descartados}")
            self.logger.info(f"  • Emails con error: {errores}")
            self.logger.info(f"  • Total archivos descargados: {total_archivos}")
            
            self.logger.info("=== PROCESO COMPLETADO ===")
            self.logger.info("📈 RESUMEN FINAL:")
            self.logger.info(f"  • Total emails en rango de fechas: {len(email_ids)}")
            self.logger.info(f"  • Emails que cumplen criterios: {len(valid_emails)}")
            self.logger.info(f"  • Emails descartados: {len(email_ids) - len(valid_emails)}")
            self.logger.info(f"  • Reporte CSV generado: {report_filename}")
            
            if filters['date_range']['enabled']:
                start_date_str = filters['date_range']['start_date'].strftime('%d/%m/%Y')
                end_date_str = filters['date_range']['end_date'].strftime('%d/%m/%Y')
                self.logger.info(f"  • Rango procesado: {start_date_str} - {end_date_str}")
            
            self.logger.info("=" * 50)
            
            return {
                'total_emails': len(email_ids),
                'valid_emails': len(valid_emails),
                'total_files': total_archivos,
                'report_file': report_filename
            }
            
        finally:
            if self.mail:
                self.mail.close()
                self.mail.logout()
    
    def disconnect(self):
        """Desconecta del servidor de email"""
        if self.mail:
            try:
                self.mail.close()
                self.mail.logout()
            except:
                pass