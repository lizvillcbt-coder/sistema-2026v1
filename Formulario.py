import datetime
import os
import re
import sqlite3
import zipfile
from io import BytesIO
import openpyxl
from openpyxl.styles import Alignment, Font
import pandas as pd
import streamlit as st

# ==========================================
# CONFIGURACIÓN INICIAL DE LA PÁGINA
# ==========================================
st.set_page_config(
    page_title="Sistema de Reinscripción 2026",
    page_icon="📝",
    layout="wide",
)

PASSWORD_ADMIN = st.secrets.get("PASSWORD_ADMIN", "admin123")
DB_FILE = "inscripciones.db"

# Mapeo de archivos de plantillas según el semestre (1ER SEMESTRE desactivado temporalmente)
PLANTILLAS_EXCEL = {
    "1ER SEMESTRE": "SOLIC INSCRIP NVO 2026.xlsx",
    "3ER SEMESTRE": "SOLICITUD REINSCRIPCION tercero.xlsx",
    "5TO SEMESTRE": "SOLICITUD REINSCRIPCION quinto.xlsx",
}

# DICCIONARIOS DE DOCUMENTOS SEGÚN EL SEMESTRE
DOCS_OPCIONES_1ER = {
    "1": "1.- Voucher de Pago Original",
    "2": "2.- Comprobante de Asignación / Folio",
    "3": "3.- Certificado de Secundaria",
    "4": "4.- Boleta de 3er Año",
    "5": "5.- CURP Alumno",
    "6": "6.- Acta de Nacimiento",
    "7": "7.- Certificado Médico",
    "8": "8.- Comprobante Domicilio",
    "9": "9.- INE Tutor",
    "10": "10.- CURP Tutor",
    "11": "11.- 3 Fotografías Infantil",
}

DOCS_OPCIONES_REINSCRIPCION = {
    "1": "1.- Voucher de Pago Original",
}


# ==========================================
# FUNCIONES DE VALIDACIÓN DE DATOS
# ==========================================
def validar_curp(curp):
    """Valida formato de CURP mexicana (18 caracteres)."""
    patron_curp = r"^[A-Z]{4}\d{6}[HM][A-Z]{5}[A-Z0-9]\d$"
    curp_clean = curp.strip().upper()
    if len(curp_clean) != 18:
        return False, "La CURP debe tener exactamente 18 caracteres."
    if not re.match(patron_curp, curp_clean):
        return False, "El formato de la CURP es inválido."
    return True, curp_clean


def validar_telefono(telefono, nombre_campo, obligatorio=True):
    """Valida números telefónicos a 10 dígitos."""
    tel_clean = re.sub(r"\D", "", telefono.strip())
    if not tel_clean:
        if obligatorio:
            return False, f"El campo **{nombre_campo}** es obligatorio.", ""
        return True, "", ""
    if len(tel_clean) != 10:
        return (
            False,
            f"El campo **{nombre_campo}** debe tener exactamente 10 dígitos numéricos.",
            tel_clean,
        )
    return True, "", tel_clean


def validar_correo(correo, obligatorio=True):
    """Valida estructura básica de correo electrónico."""
    correo_clean = correo.strip().lower()
    if not correo_clean:
        if obligatorio:
            return False, "El correo electrónico es obligatorio.", ""
        return True, "", ""
    patron_email = r"^[\w\.-]+@[\w\.-]+\.\w+$"
    if not re.match(patron_email, correo_clean):
        return False, "El correo electrónico no tiene un formato válido.", correo_clean
    return True, "", correo_clean


def validar_promedio(promedio):
    """Valida que el promedio esté entre 0 y 10."""
    prom_str = promedio.strip().replace(",", ".")
    if not prom_str:
        return False, "El promedio es obligatorio.", ""
    try:
        val = float(prom_str)
        if 0.0 <= val <= 10.0:
            return True, "", f"{val:.1f}" if val % 1 != 0 else f"{int(val)}"
        else:
            return False, "El promedio debe ser un número entre 0.0 y 10.0.", ""
    except ValueError:
        return False, "El promedio debe ser un número válido (ej. 8.5).", ""


def validar_edad(edad):
    """Valida edad numérica."""
    edad_str = re.sub(r"\D", "", edad.strip())
    if not edad_str:
        return False, "La edad es obligatoria.", ""
    try:
        val = int(edad_str)
        if 10 <= val <= 99:
            return True, "", str(val)
        return False, "La edad debe estar entre 10 y 99 años.", ""
    except ValueError:
        return False, "La edad debe ser un número entero válido.", ""


# ==========================================
# GESTIÓN BASE DE DATOS LOCAL (SQLITE)
# ==========================================
@st.cache_resource
def get_db_connection():
    """Mantiene una conexión persistente a SQLite."""
    conn = sqlite3.connect(DB_FILE, check_same_thread=False)
    return conn


def inicializar_db():
    """Crea la tabla de alumnos si no existe en la BBDD local."""
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS alumnos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha_registro TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            semestre TEXT, nombre_alumno TEXT, curp TEXT UNIQUE, fecha_nacimiento TEXT, edad TEXT,
            sexo TEXT, lugar_nacimiento TEXT, celular_alumno TEXT, correo TEXT,
            red_social TEXT, secundaria TEXT, cct TEXT, promedio TEXT, carrera TEXT,
            turno TEXT, estatus TEXT, observaciones TEXT, nombre_tutor TEXT,
            domicilio TEXT, dom_diferente INTEGER, segundo_domicilio TEXT,
            celular_tutor TEXT, tel_casa TEXT, tel_emergencia TEXT,
            ocupacion TEXT, docs_entregados TEXT
        )
    """)
    conn.commit()

    try:
        cursor.execute("ALTER TABLE alumnos ADD COLUMN semestre TEXT")
        conn.commit()
    except sqlite3.OperationalError:
        pass


def vaciar_db():
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM alumnos")
    conn.commit()


def eliminar_alumno_db(id_alumno):
    """Elimina un solo alumno por su ID de la BBDD."""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM alumnos WHERE id = ?", (id_alumno,))
    conn.commit()


def guardar_en_db(datos):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute(
        """
        INSERT INTO alumnos (
            semestre, nombre_alumno, curp, fecha_nacimiento, edad, sexo, lugar_nacimiento,
            celular_alumno, correo, red_social, secundaria, cct, promedio,
            carrera, turno, estatus, observaciones, nombre_tutor, domicilio,
            dom_diferente, segundo_domicilio, celular_tutor, tel_casa, tel_emergencia,
            ocupacion, docs_entregados
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """,
        datos,
    )
    conn.commit()


def actualizar_en_db(id_alumno, datos):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute(
        """
        UPDATE alumnos SET
            semestre=?, nombre_alumno=?, curp=?, fecha_nacimiento=?, edad=?, sexo=?, lugar_nacimiento=?,
            celular_alumno=?, correo=?, red_social=?, secundaria=?, cct=?, promedio=?,
            carrera=?, turno=?, estatus=?, observaciones=?, nombre_tutor=?, domicilio=?,
            dom_diferente=?, segundo_domicilio=?, celular_tutor=?, tel_casa=?, tel_emergencia=?,
            ocupacion=?, docs_entregados=?
        WHERE id=?
    """,
        (*datos, id_alumno),
    )
    conn.commit()


def actualizar_docs_en_db(id_alumno, cadena_docs):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute(
        "UPDATE alumnos SET docs_entregados=? WHERE id=?",
        (cadena_docs, id_alumno),
    )
    conn.commit()


def obtener_alumnos():
    conn = get_db_connection()
    return pd.read_sql_query(
        "SELECT * FROM alumnos ORDER BY fecha_registro DESC", conn
    )


def escribir_celda_segura(sheet, celda, valor, alineacion=None, fuente=None):
    try:
        cell = sheet[celda]
        cell.value = valor
        if alineacion:
            cell.alignment = alineacion
        if fuente:
            cell.font = fuente
    except AttributeError:
        pass


# ==========================================
# FUNCIÓN PARA GENERAR EXCEL RELLENADO
# ==========================================
def generar_excel_alumno(
    semestre,
    nombre_alumno,
    dia_nac,
    mes_nac,
    anio_nac,
    edad,
    lugar_nac,
    sexo,
    celular_alumno,
    correo,
    red_social,
    curp,
    secundaria,
    cct,
    promedio,
    carrera,
    turno,
    estatus,
    observaciones,
    nombre_tutor,
    domicilio,
    dom_diferente,
    segundo_domicilio,
    celular_tutor,
    tel_casa,
    tel_emergencia,
    ocupacion,
    docs_list,
):
    plantilla_target = PLANTILLAS_EXCEL.get(
        semestre, PLANTILLAS_EXCEL["3ER SEMESTRE"]
    )

    if not os.path.exists(plantilla_target):
        return None

    wb = openpyxl.load_workbook(plantilla_target)
    sheet = wb.active

    sheet.page_setup.orientation = sheet.ORIENTATION_PORTRAIT
    sheet.page_setup.paperSize = sheet.PAPERSIZE_LETTER
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 1

    font_arial_10_bold = Font(name="Arial", size=10, bold=True)
    if sheet["D25"].value:
        sheet["D25"].font = font_arial_10_bold

    font_arial_9_bold = Font(name="Arial", size=9, bold=True)
    if sheet["A31"].value:
        sheet["A31"].font = font_arial_9_bold

    escribir_celda_segura(sheet, "G2", nombre_alumno)
    escribir_celda_segura(sheet, "I5", dia_nac)
    escribir_celda_segura(sheet, "M5", mes_nac)
    escribir_celda_segura(sheet, "Q5", anio_nac)
    escribir_celda_segura(sheet, "U5", edad)
    escribir_celda_segura(sheet, "I8", lugar_nac)

    sexo_clean = str(sexo).strip().upper() if sexo else ""
    if sexo_clean == "FEMENINO":
        escribir_celda_segura(sheet, "V8", "X")
    elif sexo_clean == "MASCULINO":
        escribir_celda_segura(sheet, "V9", "X")

    escribir_celda_segura(sheet, "I10", celular_alumno)
    escribir_celda_segura(sheet, "I11", correo)
    escribir_celda_segura(sheet, "J12", red_social)

    escribir_celda_segura(sheet, "E14", curp)
    escribir_celda_segura(sheet, "I15", secundaria)
    escribir_celda_segura(sheet, "H16", cct)
    escribir_celda_segura(sheet, "T16", promedio)
    escribir_celda_segura(sheet, "H17", carrera)

    if turno == "MATUTINO":
        escribir_celda_segura(sheet, "H18", "X")
    elif turno == "VESPERTINO":
        escribir_celda_segura(sheet, "M18", "X")

    map_estatus = {
        "FOLIO DE ASIGNACIÓN": "D19",
        "ASIGNADO": "D19",
        "CAMBIO": "H19",
        "OTRO RESULTADO": "N19",
        "SIN PROCESO": "V19",
    }
    if estatus in map_estatus:
        escribir_celda_segura(sheet, map_estatus[estatus], "X")

    escribir_celda_segura(sheet, "Q18", observaciones)
    escribir_celda_segura(sheet, "G22", nombre_tutor)
    escribir_celda_segura(sheet, "H24", domicilio)

    if dom_diferente and segundo_domicilio.strip():
        texto_f26 = f"2DO DOMICILIO TUTOR: {segundo_domicilio.strip()}"
        font_f26 = Font(name="Arial", size=8, bold=True)
    else:
        texto_f26 = (
            "EL DOMICILIO DEBE SER EL MSMO DEL ALUMNO, EN CASO DE QUE SEA DIFERENTE\n"
            " ESPECIFIQUE EN ESTE ESPACIO EL 2DO Y PRESENTE AMBOS COMPROBANTES"
        )
        font_f26 = Font(name="Arial", size=8, bold=False)

    align_f26 = Alignment(horizontal="center", vertical="center", wrap_text=True)
    escribir_celda_segura(
        sheet, "D26", texto_f26, alineacion=align_f26, fuente=font_f26
    )

    escribir_celda_segura(sheet, "I27", celular_tutor)
    escribir_celda_segura(sheet, "I28", tel_casa)
    escribir_celda_segura(sheet, "I29", tel_emergencia)
    escribir_celda_segura(sheet, "I30", ocupacion)

    if semestre == "1ER SEMESTRE":
        cell_docs_map = {
            "1": "J32", "2": "J33", "3": "J34", "4": "J35", "5": "J36",
            "6": "J37", "7": "J38", "8": "V32", "9": "V34", "10": "V36", "11": "V37",
        }
    else:
        cell_docs_map = {"1": "J32"}

    for doc_num in docs_list:
        doc_num_str = str(doc_num).strip()
        if doc_num_str in cell_docs_map:
            escribir_celda_segura(sheet, cell_docs_map[doc_num_str], "X")

    hoy = datetime.date.today()
    meses_es = [
        "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
        "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"
    ]
    nombre_mes = meses_es[hoy.month - 1]
    texto_fecha_completa = f"COL. NETZAHUALCOYOTL, TEXCOCO, MEXICO. A {hoy.day:02d} DE {nombre_mes} DEL {hoy.year}"

    fuente_original = Font(name="Aptos Narrow", size=11, bold=False)
    alineacion_centrada = Alignment(
        horizontal="center", vertical="center", wrap_text=False
    )

    celda_fecha = "A41" if semestre == "1ER SEMESTRE" else "A35"
    escribir_celda_segura(
        sheet,
        celda_fecha,
        texto_fecha_completa,
        alineacion=alineacion_centrada,
        fuente=fuente_original,
    )

    excel_buffer = BytesIO()
    wb.save(excel_buffer)
    excel_buffer.seek(0)
    return excel_buffer.getvalue()


# Inicializar base de datos
inicializar_db()

# MAPEO DE TEXTOS AMIGABLES Y VALORES INTERNOS
OPCIONES_SEMESTRE_MOSTRAR = [
    "3er Semestre (2do Año)",
    "5to Semestre (3er Año)",
]

MAP_MOSTRAR_A_VALOR = {
    "3er Semestre (2do Año)": "3ER SEMESTRE",
    "5to Semestre (3er Año)": "5TO SEMESTRE",
}

MAP_VALOR_A_MOSTRAR = {
    "3ER SEMESTRE": "3er Semestre (2do Año)",
    "5TO SEMESTRE": "5to Semestre (3er Año)",
}

# ==========================================
# INTERFAZ Y NAVEGACIÓN DE STREAMLIT
# ==========================================
tab1, tab2 = st.tabs(["📝 Formulario de Reinscripción", "🔒 Panel Administrador"])

# --- TAB 1: FORMULARIO DE REINSCRIPCIÓN ---
with tab1:
    st.title("📝 Solicitud de Reinscripción 2026")

    semestre_sel_label = st.selectbox(
        "🎓 Selecciona el Semestre al que te reinscribes: *",
        OPCIONES_SEMESTRE_MOSTRAR,
        key="f_semestre",
    )
    semestre_sel = MAP_MOSTRAR_A_VALOR[semestre_sel_label]

    st.markdown("---")
    st.header("1. Datos Personales del Alumno")
    nombre_alumno = st.text_input("Nombre completo del Alumno: *", key="f_nombre")

    col_f1, col_f2 = st.columns([2, 1])
    fecha_nac = col_f1.date_input(
        "Fecha de Nacimiento: *",
        value=datetime.date(2010, 1, 1),
        min_value=datetime.date(1990, 1, 1),
        max_value=datetime.date.today(),
        format="DD/MM/YYYY",
        key="f_fnac",
    )
    edad = col_f2.text_input("Años (Edad): *", placeholder="Ej: 16", key="f_edad")

    lugar_nac = st.text_input("Lugar de Nacimiento: *", key="f_lugarnac")
    sexo = st.radio("Sexo: *", ["FEMENINO", "MASCULINO"], horizontal=True, key="f_sexo")

    col_c1, col_c2 = st.columns(2)
    celular_alumno = col_c1.text_input("Celular del Alumno (10 dígitos): *", placeholder="10 dígitos numéricos", key="f_celalumno")
    correo = col_c2.text_input("Correo electrónico: *", placeholder="ejemplo@correo.com", key="f_correo")
    red_social = st.text_input("Red Social (Facebook, Instagram, TikTok):", key="f_redsocial")

    st.header("2. Datos Académicos")
    col_a1, col_a2 = st.columns(2)
    curp = col_a1.text_input("CURP del Alumno (18 caracteres): *", placeholder="18 caracteres en mayúsculas", key="f_curp")
    promedio = col_a2.text_input("Promedio de Secundaria: *", placeholder="Ej. 8.5", key="f_promedio")

    secundaria = st.text_input("Secundaria de procedencia: *", key="f_secundaria")
    cct = st.text_input("CCT de la Secundaria: *", key="f_cct")
    carrera = st.text_input("Carrera: *", key="f_carrera")

    col_t1, col_t2 = st.columns(2)
    turno = col_t1.radio(
        "Turno: *", ["MATUTINO", "VESPERTINO"], horizontal=True, key="f_turno"
    )
    estatus = col_t2.radio(
        "Resultado: *",
        ["FOLIO DE ASIGNACIÓN", "CAMBIO", "OTRO RESULTADO", "SIN PROCESO"],
        horizontal=True,
        key="f_estatus",
    )
    observaciones = st.text_input("Observaciones:", key="f_obs")

    st.header("3. Datos del Tutor")
    nombre_tutor = st.text_input("Nombre completo del Tutor: *", key="f_tutor")
    domicilio = st.text_input(
        "Domicilio Principal del Tutor (Calle, No., Colonia, Localidad, Municipio): *",
        key="f_domicilio",
    )

    st.markdown("##### 🏡 Domicilio Secundario del Tutor")
    col_dom1, col_dom2 = st.columns([1, 2])

    dom_dif_check = col_dom1.checkbox(
        "¿El domicilio del tutor es DIFERENTE al del alumno?",
        key="chk_domicilio_diferente",
    )

    segundo_domicilio = col_dom2.text_input(
        "Especifique el 2do domicilio del tutor:",
        placeholder=(
            "Escriba aquí el segundo domicilio..."
            if dom_dif_check
            else "Active la casilla de la izquierda para habilitar este campo"
        ),
        disabled=not dom_dif_check,
        key="txt_segundo_domicilio",
    )

    col_tut1, col_tut2, col_tut3 = st.columns(3)
    celular_tutor = col_tut1.text_input("Celular del Tutor (10 dígitos): *", placeholder="10 dígitos", key="f_celtutor")
    tel_casa = col_tut2.text_input("Teléfono de Casa (Opcional):", placeholder="10 dígitos (opcional)", key="f_telcasa")
    tel_emergencia = col_tut3.text_input("Teléfono de Emergencia (10 dígitos): *", placeholder="10 dígitos", key="f_telemerg")
    ocupacion = st.text_input("Ocupación del Tutor: *", key="f_ocupacion")

    st.markdown("---")
    if st.button("💾 GUARDAR SOLICITUD LOCALMENTE", use_container_width=True, type="primary"):
        errores = []

        if not nombre_alumno.strip():
            errores.append("El **Nombre del Alumno** es obligatorio.")
        if not lugar_nac.strip():
            errores.append("El **Lugar de Nacimiento** es obligatorio.")
        if not secundaria.strip():
            errores.append("La **Secundaria de procedencia** es obligatoria.")
        if not cct.strip():
            errores.append("El **CCT de la Secundaria** es obligatorio.")
        if not carrera.strip():
            errores.append("La **Carrera** es obligatoria.")
        if not nombre_tutor.strip():
            errores.append("El **Nombre del Tutor** es obligatorio.")
        if not domicilio.strip():
            errores.append("El **Domicilio del Tutor** es obligatorio.")
        if dom_dif_check and not segundo_domicilio.strip():
            errores.append("Indicó que el domicilio es diferente pero no especificó el **2do domicilio**.")
        if not ocupacion.strip():
            errores.append("La **Ocupación del Tutor** es obligatoria.")

        curp_ok, msg_curp = validar_curp(curp)
        if not curp_ok:
            errores.append(msg_curp)
        else:
            curp = msg_curp

        edad_ok, msg_edad, edad_val = validar_edad(edad)
        if not edad_ok:
            errores.append(msg_edad)

        prom_ok, msg_prom, prom_val = validar_promedio(promedio)
        if not prom_ok:
            errores.append(msg_prom)

        mail_ok, msg_mail, mail_val = validar_correo(correo)
        if not mail_ok:
            errores.append(msg_mail)

        cel_a_ok, msg_cel_a, cel_a_val = validar_telefono(celular_alumno, "Celular del Alumno")
        if not cel_a_ok:
            errores.append(msg_cel_a)

        cel_t_ok, msg_cel_t, cel_t_val = validar_telefono(celular_tutor, "Celular del Tutor")
        if not cel_t_ok:
            errores.append(msg_cel_t)

        tel_c_ok, msg_tel_c, tel_c_val = validar_telefono(tel_casa, "Teléfono de Casa", obligatorio=False)
        if not tel_c_ok:
            errores.append(msg_tel_c)

        tel_e_ok, msg_tel_e, tel_e_val = validar_telefono(tel_emergencia, "Teléfono de Emergencia")
        if not tel_e_ok:
            errores.append(msg_tel_e)

        if errores:
            st.error("⚠️ **Corrige los siguientes errores antes de continuar:**")
            for err in errores:
                st.write(f"- {err}")
        else:
            try:
                dia_nac = f"{fecha_nac.day:02d}"
                mes_nac = f"{fecha_nac.month:02d}"
                anio_nac = str(fecha_nac.year)

                lista_docs_vacia = ""

                datos_alumno = (
                    semestre_sel,
                    nombre_alumno.strip().upper(),
                    curp,
                    f"{dia_nac}/{mes_nac}/{anio_nac}",
                    edad_val,
                    sexo,
                    lugar_nac.strip().upper(),
                    cel_a_val,
                    mail_val,
                    red_social.strip(),
                    secundaria.strip().upper(),
                    cct.strip().upper(),
                    prom_val,
                    carrera.strip().upper(),
                    turno,
                    estatus,
                    observaciones.strip(),
                    nombre_tutor.strip().upper(),
                    domicilio.strip().upper(),
                    1 if dom_dif_check else 0,
                    segundo_domicilio.strip().upper(),
                    cel_t_val,
                    tel_c_val,
                    tel_e_val,
                    ocupacion.strip().upper(),
                    lista_docs_vacia,
                )
                guardar_en_db(datos_alumno)

                st.success(f"✅ ¡Solicitud para **{semestre_sel_label}** guardada correctamente!")

                bytes_excel = generar_excel_alumno(
                    semestre_sel,
                    nombre_alumno.strip().upper(),
                    dia_nac,
                    mes_nac,
                    anio_nac,
                    edad_val,
                    lugar_nac.strip().upper(),
                    sexo,
                    cel_a_val,
                    mail_val,
                    red_social.strip(),
                    curp,
                    secundaria.strip().upper(),
                    cct.strip().upper(),
                    prom_val,
                    carrera.strip().upper(),
                    turno,
                    estatus,
                    observaciones.strip(),
                    nombre_tutor.strip().upper(),
                    domicilio.strip().upper(),
                    dom_dif_check,
                    segundo_domicilio.strip().upper(),
                    cel_t_val,
                    tel_c_val,
                    tel_e_val,
                    ocupacion.strip().upper(),
                    [],
                )

                if bytes_excel:
                    nombre_limpio = "".join(
                        c for c in nombre_alumno if c.isalnum() or c == " "
                    ).strip()
                    st.download_button(
                        label=f"📄 Descargar Solicitud en Excel ({semestre_sel_label})",
                        data=bytes_excel,
                        file_name=f"SOLICITUD_{semestre_sel.replace(' ', '_')}_{nombre_limpio}_{curp}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

            except sqlite3.IntegrityError:
                st.error(f"⚠️ La CURP **{curp}** ya se encuentra registrada en el sistema.")
            except Exception as e:
                st.error(f"Error inesperado al guardar la solicitud: {e}")

# --- TAB 2: PANEL ADMINISTRADOR ---
with tab2:
    st.title("🔒 Panel Administrador")
    pass_input = st.text_input("Contraseña de acceso:", type="password")

    if pass_input == PASSWORD_ADMIN:

        with st.expander("📤 Cargar / Reemplazar Base de Datos (.db)", expanded=False):
            uploaded_db = st.file_uploader("Selecciona un archivo .db local:", type=["db"])
            if uploaded_db is not None:
                if st.button("🔄 Cargar esta Base de Datos ahora"):
                    with open(DB_FILE, "wb") as f:
                        f.write(uploaded_db.getbuffer())
                    st.cache_resource.clear()
                    st.success("✅ Base de datos cargada y actualizada con éxito.")
                    st.rerun()

        st.markdown("---")
        df_alumnos = obtener_alumnos()
        st.write(f"**Total de alumnos registrados:** {len(df_alumnos)}")

        col_d1, col_d2 = st.columns(2)

        output_excel = BytesIO()
        with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
            df_alumnos.to_excel(writer, index=False, sheet_name="Padrón Completo")

        col_d1.download_button(
            label="📥 Exportar Padrón Completo (.xlsx)",
            data=output_excel.getvalue(),
            file_name="Padron_Inscripciones_2026.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        if os.path.exists(DB_FILE):
            with open(DB_FILE, "rb") as f_db:
                col_d2.download_button(
                    label="🗄️ Descargar Copia BBDD (.db)",
                    data=f_db.read(),
                    file_name="inscripciones.db",
                    mime="application/x-sqlite3",
                )

        st.markdown("---")
        # ==========================================
        # SECCIÓN 1: GESTIÓN DE DOCUMENTOS ENTREGADOS
        # ==========================================
        st.subheader("📋 Cotejo de Documentos Entregados (Control Escolar)")
        if not df_alumnos.empty:
            map_alumnos_docs = {
                f"[{MAP_VALOR_A_MOSTRAR.get(r.get('semestre'), r.get('semestre', '3ER SEMESTRE'))}] {r['nombre_alumno']} - CURP: {r['curp']}": r
                for _, r in df_alumnos.iterrows()
            }
            sel_alum_doc = st.selectbox(
                "Selecciona el alumno para cotejar/actualizar sus documentos:",
                list(map_alumnos_docs.keys()),
                key="sb_alum_docs",
            )
            r_doc_sel = map_alumnos_docs[sel_alum_doc]

            semestre_alumno = r_doc_sel.get("semestre", "3ER SEMESTRE") or "3ER SEMESTRE"
            docs_opciones_target = (
                DOCS_OPCIONES_1ER
                if semestre_alumno == "1ER SEMESTRE"
                else DOCS_OPCIONES_REINSCRIPCION
            )

            docs_actuales_str = str(r_doc_sel["docs_entregados"] or "")
            docs_actuales = [d.strip() for d in docs_actuales_str.split(",") if d.strip()]

            lbl_sem_actual = MAP_VALOR_A_MOSTRAR.get(semestre_alumno, semestre_alumno)
            st.write(f"Documentos requeridos para **{lbl_sem_actual}**:")

            col_doc_a, col_doc_b = st.columns(2)
            nuevos_docs_seleccionados = []

            for num_k, txt_v in docs_opciones_target.items():
                col_target = col_doc_a if int(num_k) <= 7 else col_doc_b
                marcado = col_target.checkbox(
                    txt_v,
                    value=(num_k in docs_actuales),
                    key=f"chk_doc_{r_doc_sel['id']}_{num_k}",
                )
                if marcado:
                    nuevos_docs_seleccionados.append(num_k)

            if st.button("💾 Actualizar Documentos Entregados", type="primary"):
                nuevos_docs_seleccionados.sort(key=int)
                cadena_actualizada = ",".join(nuevos_docs_seleccionados)
                actualizar_docs_en_db(r_doc_sel["id"], cadena_actualizada)
                st.success(
                    f"✅ Documentos actualizados para **{r_doc_sel['nombre_alumno']}**."
                )
                st.rerun()

        st.markdown("---")
        # ==========================================
        # SECCIÓN 2: DESCARGA MÚLTIPLE EN EXCEL AGRUPADA POR SEMESTRE (.ZIP)
        # ==========================================
        st.subheader("📦 Descarga Masiva o Individual de Solicitudes en Excel")
        if not df_alumnos.empty:

            opciones_alumnos_map = {
                f"[{MAP_VALOR_A_MOSTRAR.get(r.get('semestre'), r.get('semestre', '3ER SEMESTRE'))}] {r['nombre_alumno']} - CURP: {r['curp']}": r
                for _, r in df_alumnos.iterrows()
            }

            lista_etiquetas = list(opciones_alumnos_map.keys())

            col_btn_a, col_btn_b = st.columns(2)
            if col_btn_a.button("Select All (Seleccionar Todos)"):
                st.session_state["alumnos_seleccionados"] = lista_etiquetas
            if col_btn_b.button("Clear All (Desmarcar Todos)"):
                st.session_state["alumnos_seleccionados"] = []

            seleccionados = st.multiselect(
                "Selecciona uno o varios alumnos para empaquetar sus archivos Excel:",
                options=lista_etiquetas,
                key="alumnos_seleccionados",
            )

            if seleccionados:
                col_info, col_dl = st.columns([2, 1])
                col_info.info(f"📋 **{len(seleccionados)}** alumno(s) seleccionado(s).")

                if col_dl.button("📦 Generar Paquete ZIP con Excels"):
                    zip_buffer = BytesIO()

                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                        for etq in seleccionados:
                            r_al = opciones_alumnos_map[etq]

                            fnac_parts = (
                                str(r_al["fecha_nacimiento"]).split("/")
                                if r_al["fecha_nacimiento"]
                                else ["", "", ""]
                            )
                            d_nac = fnac_parts[0] if len(fnac_parts) > 0 else ""
                            m_nac = fnac_parts[1] if len(fnac_parts) > 1 else ""
                            a_nac = fnac_parts[2] if len(fnac_parts) > 2 else ""

                            docs_entregados_list = (
                                str(r_al["docs_entregados"]).split(",")
                                if r_al["docs_entregados"]
                                else []
                            )

                            sem_alumno = (
                                r_al.get("semestre", "3ER SEMESTRE") or "3ER SEMESTRE"
                            )

                            bytes_excel_ind = generar_excel_alumno(
                                sem_alumno,
                                r_al["nombre_alumno"],
                                d_nac,
                                m_nac,
                                a_nac,
                                r_al["edad"],
                                r_al["lugar_nacimiento"],
                                r_al["sexo"],
                                r_al["celular_alumno"],
                                r_al["correo"],
                                r_al["red_social"],
                                r_al["curp"],
                                r_al["secundaria"],
                                r_al["cct"],
                                r_al["promedio"],
                                r_al["carrera"],
                                r_al["turno"],
                                r_al["estatus"],
                                r_al["observaciones"],
                                r_al["nombre_tutor"],
                                r_al["domicilio"],
                                bool(r_al.get("dom_diferente", False)),
                                r_al.get("segundo_domicilio", ""),
                                r_al["celular_tutor"],
                                r_al["tel_casa"],
                                r_al["tel_emergencia"],
                                r_al["ocupacion"],
                                docs_entregados_list,
                            )

                            if bytes_excel_ind:
                                nom_clean = "".join(
                                    c for c in r_al["nombre_alumno"] if c.isalnum() or c == " "
                                ).strip()
                                nom_archivo = f"SOLICITUD_{sem_alumno.replace(' ', '_')}_{nom_clean}_{r_al['curp']}.xlsx"
                                ruta_dentro_zip = f"{sem_alumno}/{nom_archivo}"
                                
                                zip_file.writestr(ruta_dentro_zip, bytes_excel_ind)

                    zip_buffer.seek(0)

                    st.download_button(
                        label="⬇️ DESCARGAR ARCHIVO ZIP (AGRUPADO POR SEMESTRE)",
                        data=zip_buffer.getvalue(),
                        file_name=f"SOLICITUDES_POR_SEMESTRE_{datetime.date.today().strftime('%Y%m%d')}.zip",
                        mime="application/zip",
                        use_container_width=True,
                    )

        # ==========================================
        # SECCIÓN 3: ELIMINAR UN ALUMNO
        # ==========================================
        st.markdown("---")
        st.subheader("🗑️ Eliminar un Alumno Específico")
        if not df_alumnos.empty:
            map_eliminar = {
                f"[{MAP_VALOR_A_MOSTRAR.get(r.get('semestre'), r.get('semestre', '3ER SEMESTRE'))}] {r['nombre_alumno']} (CURP: {r['curp']})": r["id"]
                for _, r in df_alumnos.iterrows()
            }
            sel_eliminar = st.selectbox(
                "Selecciona el alumno que deseas eliminar:",
                list(map_eliminar.keys()),
                key="sb_eliminar_individual",
            )
            id_eliminar = map_eliminar[sel_eliminar]

            col_del1, col_del2 = st.columns([2, 1])
            confirmar_borrado_ind = col_del1.checkbox(
                f"Confirmar eliminación permanente del registro seleccionado",
                key="chk_confirmar_borrado_ind",
            )
            if col_del2.button("🗑️ ELIMINAR ALUMNO", type="primary") and confirmar_borrado_ind:
                eliminar_alumno_db(id_eliminar)
                st.success("✅ Alumno eliminado correctamente de la base de datos.")
                st.rerun()

        with st.expander("⚠️ Opción Masiva: Vaciar Toda la Base de Datos"):
            confirmar_vaciar = st.checkbox("Entiendo que esta acción es irreversible y borrará TODOS los registros")
            if st.button("🗑️ VACIAR BASE DE DATOS COMPLETA") and confirmar_vaciar:
                vaciar_db()
                st.success("✅ Base de datos vaciada con éxito.")
                st.rerun()

        st.markdown("---")
        st.dataframe(df_alumnos, use_container_width=True)

        st.markdown("---")
        # ==========================================
        # SECCIÓN 4: EDITAR COMPLETO EN BBDD
        # ==========================================
        st.subheader("✏️ Editar Cualquier Campo de un Alumno Registrado")
        if not df_alumnos.empty:
            opciones = {
                f"[{MAP_VALOR_A_MOSTRAR.get(r.get('semestre'), r.get('semestre', '3ER SEMESTRE'))}] {r['nombre_alumno']} (CURP: {r['curp']})": r["id"]
                for _, r in df_alumnos.iterrows()
            }
            sel_alumno = st.selectbox(
                "Selecciona alumno a modificar:", list(opciones.keys()), key="sb_editar_alumno"
            )
            id_sel = opciones[sel_alumno]
            row_sel = df_alumnos[df_alumnos["id"] == id_sel].iloc[0]

            with st.form("form_edit_admin_completo"):
                st.markdown("##### Datos Principales")
                col_esem, col_e1, col_e2 = st.columns(3)

                val_sem = row_sel.get("semestre", "3ER SEMESTRE") or "3ER SEMESTRE"
                lbl_sem_actual = MAP_VALOR_A_MOSTRAR.get(val_sem, OPCIONES_SEMESTRE_MOSTRAR[0])
                idx_sem = (
                    OPCIONES_SEMESTRE_MOSTRAR.index(lbl_sem_actual)
                    if lbl_sem_actual in OPCIONES_SEMESTRE_MOSTRAR
                    else 0
                )

                e_semestre_lbl = col_esem.selectbox("Semestre:", OPCIONES_SEMESTRE_MOSTRAR, index=idx_sem)
                e_semestre = MAP_MOSTRAR_A_VALOR[e_semestre_lbl]

                e_nombre = col_e1.text_input("Nombre:", value=row_sel["nombre_alumno"])
                e_curp = col_e2.text_input("CURP:", value=row_sel["curp"])

                col_e3, col_e4, col_e5 = st.columns(3)
                e_fnac = col_e3.text_input(
                    "Fecha Nac. (DD/MM/AAAA):",
                    value=str(row_sel["fecha_nacimiento"] or ""),
                )
                e_edad = col_e4.text_input("Edad:", value=str(row_sel["edad"] or ""))
                e_sexo = col_e5.selectbox(
                    "Sexo:",
                    ["FEMENINO", "MASCULINO"],
                    index=0 if row_sel["sexo"] == "FEMENINO" else 1,
                )

                col_e6, col_e7, col_e8 = st.columns(3)
                e_lugar_nac = col_e6.text_input(
                    "Lugar Nacimiento:",
                    value=str(row_sel["lugar_nacimiento"] or ""),
                )
                e_cel_alum = col_e7.text_input(
                    "Celular Alumno:",
                    value=str(row_sel["celular_alumno"] or ""),
                )
                e_correo = col_e8.text_input(
                    "Correo:", value=str(row_sel["correo"] or "")
                )

                e_red_social = st.text_input(
                    "Red Social:", value=str(row_sel["red_social"] or "")
                )

                st.markdown("##### Datos Académicos")
                col_ea1, col_ea2, col_ea3 = st.columns(3)
                e_secundaria = col_ea1.text_input(
                    "Secundaria:", value=str(row_sel["secundaria"] or "")
                )
                e_cct = col_ea2.text_input("CCT:", value=str(row_sel["cct"] or ""))
                e_promedio = col_ea3.text_input(
                    "Promedio:", value=str(row_sel["promedio"] or "")
                )

                col_ea4, col_ea5, col_ea6 = st.columns(3)
                e_carrera = col_ea4.text_input(
                    "Carrera:", value=str(row_sel["carrera"] or "")
                )

                opt_turno = ["MATUTINO", "VESPERTINO"]
                idx_turno = (
                    opt_turno.index(row_sel["turno"])
                    if row_sel["turno"] in opt_turno
                    else 0
                )
                e_turno = col_ea5.selectbox("Turno:", opt_turno, index=idx_turno)

                opt_estatus = [
                    "FOLIO DE ASIGNACIÓN",
                    "CAMBIO",
                    "OTRO RESULTADO",
                    "SIN PROCESO",
                ]
                idx_estatus = (
                    opt_estatus.index(row_sel["estatus"])
                    if row_sel["estatus"] in opt_estatus
                    else 0
                )
                e_estatus = col_ea6.selectbox(
                    "Resultado / Estatus:", opt_estatus, index=idx_estatus
                )

                e_obs = st.text_input(
                    "Observaciones:", value=str(row_sel["observaciones"] or "")
                )

                st.markdown("##### Datos del Tutor")
                col_et1, col_et2 = st.columns(2)
                e_tutor = col_et1.text_input(
                    "Nombre Tutor:", value=str(row_sel["nombre_tutor"] or "")
                )
                e_domicilio = col_et2.text_input(
                    "Domicilio Principal:",
                    value=str(row_sel["domicilio"] or ""),
                )

                col_edom1, col_edom2 = st.columns([1, 2])
                e_dom_dif = col_edom1.checkbox(
                    "¿Domicilio diferente?",
                    value=bool(row_sel.get("dom_diferente", 0)),
                )
                e_seg_dom = col_edom2.text_input(
                    "Segundo Domicilio Tutor:",
                    value=str(row_sel.get("segundo_domicilio", "") or ""),
                )

                col_et3, col_et4, col_et5, col_et6 = st.columns(4)
                e_cel_tut = col_et3.text_input(
                    "Celular Tutor:", value=str(row_sel["celular_tutor"] or "")
                )
                e_tel_casa = col_et4.text_input(
                    "Tel. Casa:", value=str(row_sel["tel_casa"] or "")
                )
                e_tel_emerg = col_et5.text_input(
                    "Tel. Emergencia:",
                    value=str(row_sel["tel_emergencia"] or ""),
                )
                e_ocupacion = col_et6.text_input(
                    "Ocupación:", value=str(row_sel["ocupacion"] or "")
                )

                e_docs = st.text_input(
                    "Documentos Entregados (números separados por coma):",
                    value=str(row_sel["docs_entregados"] or ""),
                )

                btn_actualizar_todo = st.form_submit_button(
                    "🔄 Guardar Cambios en BBDD Local"
                )

            if btn_actualizar_todo:
                datos_actualizados = (
                    e_semestre,
                    e_nombre.strip().upper(),
                    e_curp.strip().upper(),
                    e_fnac,
                    e_edad,
                    e_sexo,
                    e_lugar_nac.strip().upper(),
                    e_cel_alum,
                    e_correo.strip().lower(),
                    e_red_social,
                    e_secundaria.strip().upper(),
                    e_cct.strip().upper(),
                    e_promedio,
                    e_carrera.strip().upper(),
                    e_turno,
                    e_estatus,
                    e_obs,
                    e_tutor.strip().upper(),
                    e_domicilio.strip().upper(),
                    1 if e_dom_dif else 0,
                    e_seg_dom.strip().upper(),
                    e_cel_tut,
                    e_tel_casa,
                    e_tel_emerg,
                    e_ocupacion.strip().upper(),
                    e_docs,
                )
                actualizar_en_db(id_sel, datos_actualizados)
                st.success("✅ Registro actualizado correctamente.")
                st.rerun()