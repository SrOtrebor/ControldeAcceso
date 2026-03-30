# =============================================================================
# app.py — Sistema de Control de Acceso Alcorta Shopping
# =============================================================================
# Aplicación web desarrollada con Flask que gestiona el control de ingreso
# de personal (FAP, FAO y excepciones) al centro comercial.
# Verifica DNI contra listas Excel, registra cada evento y permite enviar
# reportes diarios por correo electrónico.
#
# Autor: Roberto Laforcada
# =============================================================================

import os
import smtplib
import sys
import time
import webbrowser
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

import pandas as pd
from dotenv import load_dotenv
from flask import Flask, jsonify, request, render_template, redirect, url_for, session

# =============================================================================
# CARGA DE VARIABLES DE ENTORNO
# =============================================================================
# Lee el archivo .env ubicado junto al ejecutable (modo .exe) o junto al
# script (modo desarrollo). Nunca se guardan credenciales en el código.

if getattr(sys, 'frozen', False):
    # Modo ejecutable PyInstaller
    BASE_DIR         = os.path.dirname(sys.executable)
    template_folder  = os.path.join(sys._MEIPASS, 'templates')
    static_folder    = os.path.join(sys._MEIPASS, 'static')
    app = Flask(__name__, template_folder=template_folder, static_folder=static_folder)
else:
    # Modo desarrollo
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    app = Flask(__name__)

# Carga el .env desde el mismo directorio que el script / ejecutable
load_dotenv(os.path.join(BASE_DIR, '.env'))

# =============================================================================
# CONFIGURACIÓN GENERAL (desde variables de entorno)
# =============================================================================

app.secret_key  = os.environ['SECRET_KEY']

ADMIN_USERNAME  = os.environ['ADMIN_USERNAME']
ADMIN_PASSWORD  = os.environ['ADMIN_PASSWORD']

EMAIL_SENDER    = os.environ['EMAIL_SENDER']
EMAIL_PASSWORD  = os.environ['EMAIL_PASSWORD']
EMAIL_RECEIVER  = os.environ['EMAIL_RECEIVER']
SMTP_SERVER     = os.environ['SMTP_SERVER']
SMTP_PORT       = int(os.environ['SMTP_PORT'])

# Límites del sistema de bloqueo de intentos de login
MAX_LOGIN_ATTEMPTS   = int(os.environ.get('MAX_LOGIN_ATTEMPTS', 5))
LOGIN_LOCKOUT_SECONDS = int(os.environ.get('LOGIN_LOCKOUT_SECONDS', 300))

# Rutas a los archivos de datos
REGISTROS_DIARIOS_DIR = os.path.join(BASE_DIR, 'registros_diarios')
EXCEL_FAP              = os.path.join(BASE_DIR, 'ListadoFAPs.xlsx')
EXCEL_FAO              = os.path.join(BASE_DIR, 'ListadoFAOs.xlsx')
EXCEL_EXCEPCIONES      = os.path.join(BASE_DIR, 'excepciones.xlsx')

# =============================================================================
# NOMBRES DE COLUMNAS — FORMATO INTERNO (normalizado)
# =============================================================================
COL_DNI             = 'DNI'
COL_NOMBRE_APELLIDO = 'Nombre y Apellido'
COL_NUM_PERMISO     = 'Num_Permiso'
COL_VENCE           = 'Vence'
COL_LOCAL           = 'Local'
COL_TAREA           = 'Tarea'
COL_TIPO_PERMISO    = 'Tipo de Permiso'

# Columnas originales en el Excel FAP
COL_DNI_FAP_ORIGINAL          = 'Numero'
COL_NOMBRE_FAP_ORIGINAL        = 'Nombre'
COL_APELLIDO_FAP_ORIGINAL      = 'Apellido'
COL_NUM_PERMISO_FAP_ORIGINAL   = 'FAP'
COL_VENCE_FAP_ORIGINAL         = 'Fecha Fin'
COL_LOCAL_FAP_ORIGINAL         = 'Marca'

# Columnas originales en el Excel FAO
COL_DNI_FAO_ORIGINAL          = 'Numero'
COL_NOMBRE_FAO_ORIGINAL        = 'Nombre'
COL_APELLIDO_FAO_ORIGINAL      = 'Apellido'
COL_NUM_PERMISO_FAO_ORIGINAL   = 'FAO'
COL_VENCE_FAO_ORIGINAL         = 'Fecha Fin'
COL_LOCAL_FAO_ORIGINAL         = 'Marca'
COL_TAREA_FAO_ORIGINAL         = 'Tarea'

# =============================================================================
# CACHÉ EN MEMORIA — DataFrames globales con los datos cargados
# =============================================================================
# Solo se recargan si el archivo fue modificado desde la última lectura.
df_fap              = pd.DataFrame()
df_fao              = pd.DataFrame()
df_excepciones      = pd.DataFrame()
ult_mod_fap         = 0
ult_mod_fao         = 0
ult_mod_excepciones = 0

# =============================================================================
# SISTEMA DE PROTECCIÓN CONTRA FUERZA BRUTA EN EL LOGIN
# =============================================================================
# Registra, por dirección IP, cuántos intentos fallidos se hicieron y
# cuándo fue el último. Si se supera el máximo, bloquea por X segundos.
_login_intentos = {}   # { ip: {'count': N, 'last_attempt': timestamp} }

def registrar_intento_fallido(ip):
    """Incrementa el contador de intentos fallidos para la IP dada."""
    now = time.time()
    if ip not in _login_intentos:
        _login_intentos[ip] = {'count': 0, 'last_attempt': now}
    _login_intentos[ip]['count']        += 1
    _login_intentos[ip]['last_attempt']  = now

def resetear_intentos(ip):
    """Limpia el contador de intentos fallidos tras un login exitoso."""
    _login_intentos.pop(ip, None)

def esta_bloqueado(ip):
    """
    Indica si la IP está bloqueada por exceder el límite de intentos.

    Si el período de bloqueo ya venció, resetea el contador automáticamente.

    Retorna:
        (bool bloqueado, int segundos_restantes)
    """
    if ip not in _login_intentos:
        return False, 0

    datos       = _login_intentos[ip]
    transcurrido = time.time() - datos['last_attempt']

    if datos['count'] >= MAX_LOGIN_ATTEMPTS:
        restante = int(LOGIN_LOCKOUT_SECONDS - transcurrido)
        if restante > 0:
            return True, restante
        # El bloqueo venció: resetear
        resetear_intentos(ip)

    return False, 0

# Crear el directorio de registros diarios si no existe
os.makedirs(REGISTROS_DIARIOS_DIR, exist_ok=True)


# =============================================================================
# FUNCIONES AUXILIARES
# =============================================================================

def extraer_dni_de_cuil(valor):
    """
    Extrae los 8 dígitos del DNI a partir de un número de CUIL de 11 dígitos.
    Si el valor ya es un DNI normal (7 u 8 dígitos), lo devuelve tal cual.

    Ejemplo: '20123456789' (CUIL) → '12345678' (DNI)
    """
    valor_str = str(valor).strip()
    if len(valor_str) == 11 and valor_str.isdigit():
        return valor_str[2:10]
    return valor_str


def cargar_y_procesar_excel(filepath, ult_mod_global, tipo_permiso, cols_rename_map, df_global):
    """
    Carga y normaliza un archivo Excel de autorizaciones.

    Implementa una caché por fecha de modificación: si el archivo no cambió
    desde la última carga, devuelve el DataFrame ya en memoria sin releerlo.

    Parámetros:
        filepath        : Ruta al archivo Excel.
        ult_mod_global  : Marca de tiempo de la última carga conocida.
        tipo_permiso    : Etiqueta del tipo ('FAP', 'FAO', 'Excepcion').
        cols_rename_map : Diccionario para renombrar columnas al formato interno.
        df_global       : DataFrame actualmente en caché.

    Retorna:
        (DataFrame normalizado, nueva marca de tiempo de modificación)
    """
    try:
        mod_time = os.path.getmtime(filepath)

        # Si el archivo no cambió y ya hay datos cargados, usar la caché
        if mod_time <= ult_mod_global and not df_global.empty:
            return df_global, ult_mod_global

        header_row = 1 if tipo_permiso in ['FAP', 'FAO'] else 0
        df_temp    = pd.read_excel(filepath, header=header_row)
        df_temp    = df_temp.rename(columns=cols_rename_map)

        # Normalizar DNI: numérico → entero → string → extraer de CUIL si aplica
        if COL_DNI in df_temp.columns:
            df_temp[COL_DNI] = pd.to_numeric(df_temp[COL_DNI], errors='coerce')
            df_temp.dropna(subset=[COL_DNI], inplace=True)
            df_temp[COL_DNI] = (
                df_temp[COL_DNI]
                .astype('Int64')
                .astype(str)
                .apply(extraer_dni_de_cuil)
            )

        # Normalizar fechas de vencimiento
        if COL_VENCE in df_temp.columns:
            df_temp[COL_VENCE] = pd.to_datetime(df_temp[COL_VENCE], dayfirst=True, errors='coerce')

        # Construir el campo "Nombre y Apellido" unificado
        if 'Nombre Completo' in df_temp.columns and COL_NOMBRE_APELLIDO not in df_temp.columns:
            df_temp[COL_NOMBRE_APELLIDO] = df_temp['Nombre Completo']
        elif 'Nombre' in df_temp.columns and 'Apellido' in df_temp.columns:
            df_temp[COL_NOMBRE_APELLIDO] = (
                df_temp['Nombre'].fillna('') + ' ' + df_temp['Apellido'].fillna('')
            ).str.strip()

        df_temp[COL_TIPO_PERMISO] = tipo_permiso

        # Asegurar columnas necesarias (rellenar con N/A si no existen)
        for col in [COL_LOCAL, COL_TAREA, COL_NUM_PERMISO, COL_NOMBRE_APELLIDO, 'Quien_Autoriza']:
            if col not in df_temp.columns:
                df_temp[col] = 'N/A'

        return df_temp, mod_time

    except FileNotFoundError:
        return pd.DataFrame(), 0
    except Exception as e:
        print(f"Error al cargar '{filepath}': {e}")
        return pd.DataFrame(), ult_mod_global


def cargar_autorizaciones():
    """
    Recarga los tres DataFrames de autorizaciones desde sus Excel
    (solo si alguno fue modificado desde la última carga).
    """
    global df_fap, ult_mod_fap, df_fao, ult_mod_fao, df_excepciones, ult_mod_excepciones

    mapa_cols_fap = {
        COL_DNI_FAP_ORIGINAL:         COL_DNI,
        COL_NOMBRE_FAP_ORIGINAL:      'Nombre',
        COL_APELLIDO_FAP_ORIGINAL:    'Apellido',
        COL_NUM_PERMISO_FAP_ORIGINAL: COL_NUM_PERMISO,
        COL_VENCE_FAP_ORIGINAL:       COL_VENCE,
        COL_LOCAL_FAP_ORIGINAL:       COL_LOCAL,
    }
    df_fap, ult_mod_fap = cargar_y_procesar_excel(
        EXCEL_FAP, ult_mod_fap, 'FAP', mapa_cols_fap, df_fap
    )

    mapa_cols_fao = {
        COL_DNI_FAO_ORIGINAL:         COL_DNI,
        COL_NOMBRE_FAO_ORIGINAL:      'Nombre',
        COL_APELLIDO_FAO_ORIGINAL:    'Apellido',
        COL_NUM_PERMISO_FAO_ORIGINAL: COL_NUM_PERMISO,
        COL_VENCE_FAO_ORIGINAL:       COL_VENCE,
        COL_LOCAL_FAO_ORIGINAL:       COL_LOCAL,
        COL_TAREA_FAO_ORIGINAL:       COL_TAREA,
    }
    df_fao, ult_mod_fao = cargar_y_procesar_excel(
        EXCEL_FAO, ult_mod_fao, 'FAO', mapa_cols_fao, df_fao
    )

    mapa_cols_excepciones = {
        'Numero':          COL_DNI,
        'Nombre Completo': 'Nombre Completo',
        'Fecha de Alta':   COL_VENCE,
        'Local':           COL_LOCAL,
        'Quien Autoriza':  'Quien_Autoriza',
    }
    df_excepciones, ult_mod_excepciones = cargar_y_procesar_excel(
        EXCEL_EXCEPCIONES, ult_mod_excepciones, 'Excepcion', mapa_cols_excepciones, df_excepciones
    )


def guardar_registro(dni, nombre, hora_ingreso, tipo_permiso, num_permiso, local, tarea, resultado):
    """
    Agrega un nuevo evento de acceso al archivo Excel del día actual.

    Cada día genera un archivo separado con el formato:
        registros_ingreso_YYYY-MM-DD.xlsx
    """
    fecha_actual_str  = datetime.now().strftime('%Y-%m-%d')
    nombre_archivo    = os.path.join(REGISTROS_DIARIOS_DIR, f'registros_ingreso_{fecha_actual_str}.xlsx')
    columnas_registro = ['DNI', 'Nombre y Apellido', 'Hora_Ingreso', 'Tipo_Permiso',
                         'Num_Permiso', 'Local', 'Tarea', 'Resultado']

    nuevo_registro_df = pd.DataFrame([{
        'DNI':               dni,
        'Nombre y Apellido': nombre,
        'Hora_Ingreso':      hora_ingreso,
        'Tipo_Permiso':      tipo_permiso,
        'Num_Permiso':       num_permiso,
        'Local':             local,
        'Tarea':             tarea,
        'Resultado':         resultado,
    }])

    try:
        if not os.path.exists(nombre_archivo):
            df_registros = pd.DataFrame(columns=columnas_registro)
        else:
            df_registros = pd.read_excel(nombre_archivo)

        df_final = pd.concat([df_registros, nuevo_registro_df], ignore_index=True)
        df_final.to_excel(nombre_archivo, index=False)

    except Exception as e:
        print(f"Error crítico al guardar registro en Excel: {e}")


def crear_reporte_formateado(ruta_original):
    """
    Genera una versión formateada del archivo de registros del día,
    con colores por resultado y un resumen al encabezado.

    Retorna la ruta del archivo formateado, o None si hubo un error.
    """
    try:
        df = pd.read_excel(ruta_original)
        if df.empty:
            return None

        df = df.iloc[::-1].reset_index(drop=True)
        columna_resultado = df['Resultado']
        df_final_reporte  = df.drop(columns=['Resultado'])
        df_final_reporte  = df_final_reporte.rename(columns={'Num_Permiso': 'Permiso/Autoriza'})

        if 'Local' in df_final_reporte.columns:
            df_final_reporte['Local'] = (
                df_final_reporte['Local'].astype(str).apply(lambda x: x.split('\n')[0])
            )

        ruta_formateada = ruta_original.replace('.xlsx', '_formateado.xlsx')
        writer    = pd.ExcelWriter(ruta_formateada, engine='xlsxwriter')
        df_final_reporte.to_excel(writer, sheet_name='Reporte de Ingresos', index=False, startrow=3)

        workbook  = writer.book
        worksheet = writer.sheets['Reporte de Ingresos']

        header_format    = workbook.add_format({
            'bold': True, 'text_wrap': True, 'valign': 'top',
            'fg_color': '#2B2D42', 'font_color': 'white', 'border': 1
        })
        title_format     = workbook.add_format({'bold': True, 'font_size': 16, 'font_color': '#2B2D42'})
        summary_format   = workbook.add_format({'bold': True, 'font_size': 11})
        green_format     = workbook.add_format({'bg_color': '#eaf7ed'})
        red_format       = workbook.add_format({'bg_color': '#fdecea'})
        excepcion_format = workbook.add_format({'bg_color': '#e0f7fa'})
        fao_format       = workbook.add_format({'bg_color': '#fff9c4'})

        total      = len(df)
        permitidos = (columna_resultado == 'VERDE').sum()
        denegados  = total - permitidos

        worksheet.merge_range(
            'A1:D1',
            f"Reporte de Ingresos - {datetime.now().strftime('%d/%m/%Y')}",
            title_format
        )
        worksheet.write('A2', 'Resumen:', summary_format)
        worksheet.write('B2', f'Total: {total} | Permitidos: {permitidos} | Denegados: {denegados}')

        for col_num, value in enumerate(df_final_reporte.columns.values):
            worksheet.write(3, col_num, value, header_format)

        num_filas = len(df_final_reporte)
        for i, (idx, row) in enumerate(df_final_reporte.iterrows()):
            fila_excel        = i + 4
            resultado_actual  = columna_resultado.iloc[i]
            formato_a_aplicar = None

            if resultado_actual == 'VERDE':
                formato_a_aplicar = green_format
            elif resultado_actual == 'ROJO':
                formato_a_aplicar = red_format

            if row['Tipo_Permiso'] == 'Excepcion':
                formato_a_aplicar = excepcion_format
            elif row['Tipo_Permiso'] == 'FAO':
                formato_a_aplicar = fao_format

            if formato_a_aplicar:
                worksheet.set_row(fila_excel, None, formato_a_aplicar)

        for idx, col in enumerate(df_final_reporte):
            max_len = max(df_final_reporte[col].astype(str).map(len).max(), len(str(col))) + 3
            worksheet.set_column(idx, idx, max_len)

        worksheet.autofilter(3, 0, num_filas + 3, len(df_final_reporte.columns) - 1)
        worksheet.freeze_panes(4, 0)

        writer.close()
        return ruta_formateada

    except Exception as e:
        print(f"Error al crear el reporte formateado: {e}")
        return None


def enviar_email(archivo_adjunto=None):
    """
    Envía el reporte diario por correo electrónico al destinatario configurado.
    Las credenciales se obtienen de las variables de entorno.

    Retorna True si el envío fue exitoso, False en caso contrario.
    """
    msg            = MIMEMultipart()
    msg['From']    = EMAIL_SENDER
    msg['To']      = EMAIL_RECEIVER
    msg['Subject'] = f"Reporte Diario de Admisión - {datetime.now().strftime('%Y-%m-%d')}"
    msg.attach(MIMEText("Adjunto el reporte diario de admisiones.", 'plain'))

    if archivo_adjunto and os.path.exists(archivo_adjunto):
        try:
            with open(archivo_adjunto, 'rb') as f:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header(
                'Content-Disposition',
                f'attachment; filename={os.path.basename(archivo_adjunto)}'
            )
            msg.attach(part)
        except Exception as e:
            print(f"Error al adjuntar archivo: {e}")

    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(EMAIL_SENDER, EMAIL_PASSWORD)
            server.sendmail(EMAIL_SENDER, EMAIL_RECEIVER, msg.as_string())
        return True
    except Exception as e:
        print(f"Error al enviar el correo: {e}")
        return False


# =============================================================================
# RUTAS DE FLASK — Páginas y API
# =============================================================================

@app.route('/')
def home():
    """Página principal: panel de escaneo de DNI."""
    return render_template('index.html')


@app.route('/login')
def login_page():
    """Página de inicio de sesión del panel de administración."""
    return render_template('login.html')


@app.route('/admin')
def admin_page():
    """Panel de administración. Requiere sesión activa."""
    if session.get('logged_in'):
        return render_template('admin.html')
    return redirect(url_for('login_page'))


@app.route('/perform_login', methods=['POST'])
def perform_login():
    """
    Endpoint de autenticación (POST/JSON).

    Incluye protección contra fuerza bruta: bloquea la IP del solicitante
    tras superar el límite de intentos fallidos configurado en .env.

    Body esperado: { "username": "...", "password": "..." }
    """
    ip = request.remote_addr

    # Verificar si la IP está bloqueada antes de procesar
    bloqueado, segundos = esta_bloqueado(ip)
    if bloqueado:
        minutos = segundos // 60
        return jsonify({
            'success': False,
            'message': f'Demasiados intentos fallidos. Intente nuevamente en {minutos} min {segundos % 60} seg.'
        }), 429

    data = request.get_json()
    if data.get('username') == ADMIN_USERNAME and data.get('password') == ADMIN_PASSWORD:
        resetear_intentos(ip)
        session['logged_in'] = True
        return jsonify({'success': True})

    # Login incorrecto: registrar intento fallido
    registrar_intento_fallido(ip)
    datos_ip = _login_intentos.get(ip, {})
    restantes = MAX_LOGIN_ATTEMPTS - datos_ip.get('count', 0)

    if restantes <= 0:
        return jsonify({
            'success': False,
            'message': f'Acceso bloqueado por exceso de intentos. Espere {LOGIN_LOCKOUT_SECONDS // 60} minutos.'
        }), 429

    return jsonify({
        'success': False,
        'message': f'Usuario o contraseña incorrectos. Intentos restantes: {restantes}.'
    })


@app.route('/logout')
def logout():
    """Cierra la sesión del administrador y redirige al login."""
    session.pop('logged_in', None)
    return redirect(url_for('login_page'))


@app.route('/verificar_dni', methods=['POST'])
def verificar_dni():
    """
    Endpoint principal de control de acceso (POST/JSON).

    Busca el DNI en las listas de autorizaciones (Excepciones → FAP → FAO),
    evalúa la vigencia del permiso y registra el evento.

    Body esperado: { "dni": "12345678" }
    """
    data          = request.get_json()
    dni_ingresado = str(data.get('dni', '')).strip()

    cargar_autorizaciones()

    hoy             = datetime.now()
    hora_actual_str = hoy.strftime('%H:%M:%S')

    respuesta = {
        'acceso':       'DENEGADO',
        'nombre':       'No Encontrado',
        'mensaje':      f'ACCESO DENEGADO: DNI {dni_ingresado} no encontrado o permiso vencido.',
        'tipo_permiso': 'N/A',
        'num_permiso':  'N/A',
        'local':        'N/A',
        'tarea':        'N/A',
        'vence':        'N/A',
    }

    if not dni_ingresado.isdigit() or len(dni_ingresado) < 7:
        respuesta['mensaje'] = 'Formato de DNI inválido.'
        guardar_registro(dni_ingresado, 'N/A', hora_actual_str, 'N/A', 'N/A', 'N/A', 'N/A', 'ROJO')
        return jsonify(respuesta)

    for df, tipo in [(df_excepciones, 'Excepcion'), (df_fap, 'FAP'), (df_fao, 'FAO')]:
        if df.empty or COL_DNI not in df.columns:
            continue

        match = df[df[COL_DNI] == dni_ingresado]
        if not match.empty:
            persona           = match.iloc[0]
            fecha_vencimiento = persona.get(COL_VENCE)

            respuesta['nombre']       = persona.get(COL_NOMBRE_APELLIDO, 'N/A')
            respuesta['tipo_permiso'] = tipo

            if tipo == 'Excepcion':
                respuesta['num_permiso'] = str(persona.get('Quien_Autoriza', 'N/A'))
            else:
                respuesta['num_permiso'] = str(persona.get(COL_NUM_PERMISO, 'N/A'))

            local_raw         = persona.get(COL_LOCAL, 'N/A')
            respuesta['local'] = str(local_raw).split('\n')[0].replace('Local', '').strip()
            respuesta['tarea'] = str(persona.get(COL_TAREA, 'N/A'))

            if pd.isna(fecha_vencimiento) or fecha_vencimiento.date() >= hoy.date():
                respuesta['acceso']  = 'PERMITIDO'
                respuesta['mensaje'] = f"ACCESO PERMITIDO ({tipo}): {respuesta['nombre']}"
                respuesta['vence']   = (
                    'Indefinido' if pd.isna(fecha_vencimiento)
                    else fecha_vencimiento.strftime('%d/%m/%Y')
                )
            else:
                respuesta['acceso']  = 'DENEGADO'
                respuesta['vence']   = fecha_vencimiento.strftime('%d/%m/%Y')
                respuesta['mensaje'] = f"ACCESO DENEGADO: Permiso {tipo} vencido el {respuesta['vence']}."

            resultado_log = 'VERDE' if respuesta['acceso'] == 'PERMITIDO' else 'ROJO'
            guardar_registro(
                dni_ingresado, respuesta['nombre'], hora_actual_str,
                respuesta['tipo_permiso'], respuesta['num_permiso'],
                respuesta['local'], respuesta['tarea'], resultado_log
            )
            return jsonify(respuesta)

    guardar_registro(dni_ingresado, 'No Encontrado', hora_actual_str,
                     'N/A', 'N/A', 'N/A', 'N/A', 'ROJO')
    return jsonify(respuesta)


@app.route('/upload_excel', methods=['POST'])
def upload_excel():
    """
    Reemplaza los archivos Excel de listas FAP y/o FAO (POST/FormData).
    Solo accesible con sesión activa.
    """
    if not session.get('logged_in'):
        return jsonify({'success': False, 'message': 'No autorizado.'}), 403

    messages = []
    success  = True

    def procesar_archivo(file_key, target_path):
        nonlocal success
        archivo = request.files.get(file_key)
        if archivo and archivo.filename != '':
            try:
                temp_path = target_path + '.tmp'
                archivo.save(temp_path)
                pd.read_excel(temp_path, header=1 if file_key.startswith('f') else 0)
                os.replace(temp_path, target_path)
                messages.append(f'Archivo {file_key.split("_")[0].upper()} actualizado.')
            except Exception as e:
                success = False
                messages.append(f'Error al procesar {archivo.filename}: {e}')
                if os.path.exists(temp_path):
                    os.remove(temp_path)

    procesar_archivo('fap_file', EXCEL_FAP)
    procesar_archivo('fao_file', EXCEL_FAO)
    cargar_autorizaciones()

    return jsonify({'success': success, 'message': ' '.join(messages)})


@app.route('/add_excepcion', methods=['POST'])
def add_excepcion():
    """
    Agrega o actualiza una excepción temporal de acceso (POST/JSON).
    Solo accesible con sesión activa.
    """
    if not session.get('logged_in'):
        return jsonify({'success': False, 'message': 'No autorizado.'}), 403

    data           = request.get_json()
    nombre         = data.get('nombre', '')
    apellido       = data.get('apellido', '')
    dni            = data.get('dni', '')
    local          = data.get('local', '')
    quien_autoriza = data.get('quien_autoriza', '')

    if not all([nombre, apellido, dni, local, quien_autoriza]):
        return jsonify({'success': False, 'message': 'Todos los campos son obligatorios.'})

    try:
        columnas = ['Numero', 'Nombre Completo', 'Local', 'Quien Autoriza', 'Fecha de Alta']

        if os.path.exists(EXCEL_EXCEPCIONES):
            df_excepciones_actual = pd.read_excel(EXCEL_EXCEPCIONES)
        else:
            df_excepciones_actual = pd.DataFrame(columns=columnas)

        if 'Numero' in df_excepciones_actual.columns:
            df_excepciones_actual['Numero'] = df_excepciones_actual['Numero'].astype(str)

        mask      = df_excepciones_actual['Numero'] == dni
        timestamp = datetime.now().strftime('%d/%m/%Y %H:%M:%S')

        if mask.any():
            idx = df_excepciones_actual[mask].index
            df_excepciones_actual.loc[idx, ['Nombre Completo', 'Local', 'Quien Autoriza', 'Fecha de Alta']] = [
                f"{nombre} {apellido}", local, quien_autoriza, timestamp
            ]
            msg = 'Excepción actualizada.'
        else:
            nuevo_registro = pd.DataFrame([{
                'Numero':          dni,
                'Nombre Completo': f"{nombre} {apellido}",
                'Local':           local,
                'Quien Autoriza':  quien_autoriza,
                'Fecha de Alta':   timestamp,
            }])
            df_excepciones_actual = pd.concat([df_excepciones_actual, nuevo_registro], ignore_index=True)
            msg = 'Excepción agregada.'

        df_excepciones_actual.to_excel(EXCEL_EXCEPCIONES, index=False)

        global ult_mod_excepciones
        ult_mod_excepciones = 0
        cargar_autorizaciones()

        return jsonify({'success': True, 'message': msg})

    except Exception as e:
        return jsonify({'success': False, 'message': f'Error al guardar la excepción: {e}'})


@app.route('/enviar_reporte_diario', methods=['POST'])
def enviar_reporte():
    """
    Genera el reporte formateado del día y lo envía por email.
    Solo accesible con sesión activa.
    """
    if not session.get('logged_in'):
        return jsonify({'success': False, 'message': 'No autorizado.'}), 403

    fecha_hoy        = datetime.now().strftime('%Y-%m-%d')
    archivo_original = os.path.join(REGISTROS_DIARIOS_DIR, f'registros_ingreso_{fecha_hoy}.xlsx')

    if not os.path.exists(archivo_original):
        return jsonify({'success': False, 'message': 'No hay registros de entradas para hoy.'})

    archivo_formateado = crear_reporte_formateado(archivo_original)
    if not archivo_formateado:
        return jsonify({'success': False, 'message': 'Error al generar el reporte formateado.'})

    if enviar_email(archivo_formateado):
        os.remove(archivo_formateado)
        return jsonify({'success': True, 'message': 'Reporte enviado exitosamente.'})

    return jsonify({'success': False, 'message': 'Error al enviar el reporte.'})


@app.route('/get_daily_records')
def get_daily_records():
    """
    Retorna todos los registros de acceso del día actual (GET/JSON).
    Usado por el frontend para actualizar la tabla en tiempo real.
    """
    fecha_hoy      = datetime.now().strftime('%Y-%m-%d')
    nombre_archivo = os.path.join(REGISTROS_DIARIOS_DIR, f'registros_ingreso_{fecha_hoy}.xlsx')

    if not os.path.exists(nombre_archivo):
        return jsonify({'success': True, 'records': [], 'message': 'Aún no hay registros para hoy.'})

    try:
        df_registros = pd.read_excel(nombre_archivo).fillna('')
        return jsonify({'success': True, 'records': df_registros.to_dict('records')})
    except Exception as e:
        print(f"Error al leer registros diarios: {e}")
        return jsonify({'success': False, 'message': 'Error al procesar registros.', 'records': []}), 500


@app.route('/get_daily_stats')
def get_daily_stats():
    """
    Retorna el conteo de ingresos del día: total, permitidos y rechazados (GET/JSON).
    Usado por el dashboard en tiempo real del panel principal.
    """
    fecha_hoy      = datetime.now().strftime('%Y-%m-%d')
    nombre_archivo = os.path.join(REGISTROS_DIARIOS_DIR, f'registros_ingreso_{fecha_hoy}.xlsx')

    if not os.path.exists(nombre_archivo):
        return jsonify({'total': 0, 'permitidos': 0, 'rechazados': 0})

    try:
        df = pd.read_excel(nombre_archivo)
        if df.empty:
            return jsonify({'total': 0, 'permitidos': 0, 'rechazados': 0})

        total      = len(df)
        permitidos = int((df['Resultado'] == 'VERDE').sum())
        return jsonify({'total': total, 'permitidos': permitidos, 'rechazados': total - permitidos})

    except Exception as e:
        print(f"Error al calcular estadísticas: {e}")
        return jsonify({'total': 'Err', 'permitidos': 'Err', 'rechazados': 'Err'})


# =============================================================================
# PUNTO DE ENTRADA
# =============================================================================

if __name__ == '__main__':
    webbrowser.open('http://127.0.0.1:5000/')
    app.run(host='127.0.0.1', port=5000)