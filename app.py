import os
import re
import sqlite3
from datetime import datetime, timedelta
import io
import streamlit as st

try:
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    OPENPYXL_DISPONIBLE = True
except ImportError:
    OPENPYXL_DISPONIBLE = False

# =====================================================================
# CONFIGURACIÓN Y ESTILOS INSPIRADOS EN SUNAT
# =====================================================================
st.set_page_config(
    page_title="Sistema de Gestión - Estilo SUNAT",
    layout="wide"
)

# Estilos CSS con paleta institucional SUNAT (Azul #003366 y Rojo #D91A2A)
st.markdown("""
    <style>
    /* Fondo principal y tipografía general */
    .main {
        background-color: #F4F6F9;
        font-family: 'Segoe UI', Arial, sans-serif;
    }
    
    /* Encabezado Principal al estilo SUNAT */
    .titulo-corporativo {
        color: #003366;
        font-size: 24px;
        font-weight: 700;
        letter-spacing: -0.3px;
        margin-bottom: 20px;
        padding-bottom: 10px;
        border-bottom: 3px solid #D91A2A;
    }
    
    /* Subtítulos de sección */
    .subtitulo-corporativo {
        color: #003366;
        font-size: 15px;
        font-weight: 700;
        margin-top: 15px;
        margin-bottom: 10px;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }

    /* Fondo de la Barra Lateral */
    [data-testid="stSidebar"] {
        background-color: #FFFFFF;
        border-right: 1px solid #E0E0E0;
    }
    [data-testid="stSidebar"] p, [data-testid="stSidebar"] label, [data-testid="stSidebar"] span {
        color: #333333 !important;
    }
    [data-testid="stSidebar"] input, [data-testid="stSidebar"] select {
        color: #000000 !important;
        background-color: #F8F9FA !important;
        border: 1px solid #CCCCCC !important;
    }

    /* Pestañas (Tabs) estilo SUNAT */
    .stTabs [data-baseweb="tab-list"] {
        gap: 6px;
        background-color: #E9ECEF;
        padding: 6px;
        border-radius: 4px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 38px;
        white-space: pre-wrap;
        border-radius: 4px;
        font-weight: 600;
        color: #003366;
    }
    .stTabs [aria-selected="true"] {
        background-color: #D91A2A !important;
        color: #FFFFFF !important;
    }

    /* Botones estilo Primario SUNAT (Aplica globalmente y en Sidebar) */
    div.stButton > button {
        background-color: #003366 !important;
        color: #FFFFFF !important;
        font-weight: 700 !important;
        font-size: 15px !important;
        border-radius: 4px !important;
        border: none !important;
    }
    div.stButton > button * {
        color: #FFFFFF !important;
    }
    div.stButton > button:hover {
        background-color: #D91A2A !important;
        color: #FFFFFF !important;
    }
    </style>
""", unsafe_allow_html=True)

CARPETA_VOUCHERS = "vouchers"
if not os.path.exists(CARPETA_VOUCHERS):
    os.makedirs(CARPETA_VOUCHERS)

# =====================================================================
# 1. BASE DE DATOS Y UTILIDADES
# =====================================================================
def conectar_bd():
    conn = sqlite3.connect("control_deudas.db", check_same_thread=False)
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn

def inicializar_bd():
    conn = conectar_bd()
    cursor = conn.cursor()
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS clientes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        ruc TEXT UNIQUE NOT NULL,
        razon_social TEXT NOT NULL,
        contacto TEXT,
        regimen TEXT DEFAULT 'MYPE Tributario'
    )
    """)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS obligaciones (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        cliente_id INTEGER,
        tipo TEXT NOT NULL,
        periodo TEXT NOT NULL,
        monto REAL NOT NULL,
        fecha_vencimiento TEXT,
        estado TEXT DEFAULT 'PENDIENTE',
        fecha_pago TEXT,
        situacion TEXT DEFAULT 'Declarado',
        ruta_voucher TEXT,
        descripcion_tributos TEXT,
        FOREIGN KEY (cliente_id) REFERENCES clientes(id) ON DELETE CASCADE
    )
    """)
    conn.commit()
    conn.close()

def validar_periodo(periodo_str):
    if not periodo_str:
        return False
    return bool(re.match(r"^\d{4}-(0[1-9]|1[0-2])$", periodo_str))

def calcular_vencimiento_sunat_aproximado(ruc, periodo):
    if not ruc or len(ruc) != 11 or not validar_periodo(periodo):
        return ""
    try:
        ultimo_digito = int(ruc[-1])
        año, mes = map(int, periodo.split("-"))
        if mes == 12:
            mes_venc = 1
            año_venc = año + 1
        else:
            mes_venc = mes + 1
            año_venc = año

        dias_offset = {0: 14, 1: 15, 2: 16, 3: 17, 4: 18, 5: 19, 6: 20, 7: 21, 8: 22, 9: 23}
        dia = dias_offset.get(ultimo_digito, 15)
        return f"{año_venc:04d}-{mes_venc:02d}-{dia:02d}"
    except Exception:
        return ""

def guardar_voucher_local(uploaded_file, id_obligacion):
    if uploaded_file is None:
        return None
    ext = os.path.splitext(uploaded_file.name)[1]
    nombre_destino = f"voucher_{id_obligacion}_{datetime.now().strftime('%Y%m%d_%H%M%S')}{ext}"
    ruta_destino = os.path.join(CARPETA_VOUCHERS, nombre_destino)
    with open(ruta_destino, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return ruta_destino

def generar_excel_bytes(es_pago, texto_busqueda, id_cliente_filtro, filtro_periodo):
    conn = conectar_bd()
    cursor = conn.cursor()
    estado_filtro = "PAGADO" if es_pago else "PENDIENTE"

    query = f"""
    SELECT o.id, c.razon_social, c.ruc, c.regimen, o.tipo, o.situacion, o.periodo, o.monto, o.fecha_vencimiento
    {", o.fecha_pago" if es_pago else ""}, o.descripcion_tributos
    FROM obligaciones o JOIN clientes c ON o.cliente_id = c.id WHERE o.estado = '{estado_filtro}'
    """
    params = []
    if texto_busqueda:
        query += " AND (c.razon_social LIKE ? OR c.ruc LIKE ?)"
        params.extend([f"%{texto_busqueda}%", f"%{texto_busqueda}%"])
    if id_cliente_filtro and id_cliente_filtro != 0:
        query += " AND o.cliente_id = ?"
        params.append(id_cliente_filtro)
    if filtro_periodo:
        query += " AND o.periodo LIKE ?"
        params.append(f"%{filtro_periodo}%")

    query += " ORDER BY o.fecha_vencimiento ASC" if not es_pago else " ORDER BY o.fecha_pago DESC"
    cursor.execute(query, params)
    datos = cursor.fetchall()
    conn.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Pagados" if es_pago else "Pendientes"
    ws.views.sheetView[0].showGridLines = True

    color_navy = "003366"
    fill_titulo = PatternFill(start_color=color_navy, end_color=color_navy, fill_type="solid")
    fill_cabecera = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    fill_zebra = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

    font_titulo = Font(name="Segoe UI", size=13, bold=True, color="FFFFFF")
    font_cabecera = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
    font_datos = Font(name="Segoe UI", size=10)
    font_total = Font(name="Segoe UI", size=10, bold=True)

    align_centro = Alignment(horizontal="center", vertical="center")
    align_derecha = Alignment(horizontal="right", vertical="center")
    align_izq = Alignment(horizontal="left", vertical="center")

    border_cell = Border(left=Side(style="thin", color="E2E8F0"), right=Side(style="thin", color="E2E8F0"),
                         top=Side(style="thin", color="E2E8F0"), bottom=Side(style="thin", color="E2E8F0"))
    border_total = Border(top=Side(style="thin", color="003366"), bottom=Side(style="double", color="003366"))

    headers = ["ID", "Razón Social", "RUC", "Régimen", "Tipo", "Situación", "Periodo", "Monto", "Vencimiento"]
    if es_pago:
        headers.append("Fecha Pago")
    headers.append("Detalle Tributos")

    max_letra = openpyxl.utils.get_column_letter(len(headers))
    ws.merge_cells(f"A1:{max_letra}1")
    ws["A1"] = f"REPORTE DE OBLIGACIONES TRIBUTARIAS - STATUS: {ws.title.upper()}"
    ws["A1"].font = font_titulo
    ws["A1"].fill = fill_titulo
    ws["A1"].alignment = align_centro
    ws.row_dimensions[1].height = 35

    ws.append([])
    ws.append(headers)
    ws.row_dimensions[3].height = 24

    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=3, column=col_idx)
        cell.font = font_cabecera
        cell.fill = fill_cabecera
        cell.alignment = align_centro
        cell.border = border_cell

    for idx, fila in enumerate(datos):
        ws.append(list(fila))
        r_idx = idx + 4
        ws.row_dimensions[r_idx].height = 20
        for c_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=r_idx, column=c_idx)
            cell.font = font_datos
            cell.border = border_cell
            if idx % 2 == 1:
                cell.fill = fill_zebra

            if c_idx in [1, 3, 4, 5, 6, 7, 9] or (c_idx == 10 and es_pago):
                cell.alignment = align_centro
            elif c_idx == 8:
                cell.alignment = align_derecha
                cell.number_format = '"S/"#,##0.00'
            else:
                cell.alignment = align_izq

    tot_row = len(datos) + 4
    ws.cell(row=tot_row, column=7, value="TOTAL GENERAL:").font = font_total
    ws.cell(row=tot_row, column=7).alignment = align_derecha

    c_sum = ws.cell(row=tot_row, column=8, value=f"=SUM(H4:H{tot_row-1})")
    c_sum.font = font_total
    c_sum.alignment = align_derecha
    c_sum.number_format = '"S/"#,##0.00'
    c_sum.border = border_total

    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col if cell.row > 2)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 11)

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

inicializar_bd()

# =====================================================================
# 2. INTERFAZ PRINCIPAL
# =====================================================================
st.markdown('<div class="titulo-corporativo">SUNAT - SISTEMA DE CONTROL DE OBLIGACIONES Y COMPROBANTES</div>', unsafe_allow_html=True)

# Cargar catálogo de clientes
conn = conectar_bd()
cursor = conn.cursor()
cursor.execute("SELECT id, ruc, razon_social FROM clientes ORDER BY razon_social ASC")
clientes_db = cursor.fetchall()
conn.close()

dic_clientes = {cl[0]: cl[2] for cl in clientes_db}

# BARRA LATERAL ESTILO SUNAT
st.sidebar.markdown('<div style="font-size: 15px; font-weight: 700; color: #003366; margin-bottom: 15px; border-bottom: 2px solid #D91A2A; padding-bottom: 5px;">MÓDULO DE REGISTRO</div>', unsafe_allow_html=True)

opcion_sidebar = st.sidebar.radio("Seleccione Operación:", ["Registro de Contribuyente", "Registro de Obligación"])

if opcion_sidebar == "Registro de Contribuyente":
    st.sidebar.markdown('<div class="subtitulo-corporativo">Datos del Contribuyente</div>', unsafe_allow_html=True)
    reg_ruc = st.sidebar.text_input("Número de RUC (11 dígitos):", max_chars=11)
    reg_razon = st.sidebar.text_input("Razón Social / Nombre Comercial:")
    reg_regimen = st.sidebar.selectbox("Régimen Tributario:", ["MYPE Tributario", "Régimen General", "RER (Especial)", "NRUS"])
    reg_contacto = st.sidebar.text_input("Contacto Administrativo:")

    if st.sidebar.button("Guardar Contribuyente", use_container_width=True):
        if len(reg_ruc) != 11 or not reg_ruc.isdigit():
            st.sidebar.error("El RUC debe contar con exactamente 11 dígitos numéricos.")
        elif not reg_razon.strip():
            st.sidebar.error("El campo Razón Social es obligatorio.")
        else:
            try:
                conn = conectar_bd()
                c = conn.cursor()
                c.execute("INSERT INTO clientes (ruc, razon_social, contacto, regimen) VALUES (?, ?, ?, ?)",
                          (reg_ruc, reg_razon, reg_contacto, reg_regimen))
                conn.commit()
                conn.close()
                st.sidebar.success(f"Contribuyente '{reg_razon}' registrado correctamente.")
                st.rerun()
            except sqlite3.IntegrityError:
                st.sidebar.error("El número de RUC ingresado ya se encuentra registrado.")

elif opcion_sidebar == "Registro de Obligación":
    st.sidebar.markdown('<div class="subtitulo-corporativo">Detalle de la Obligación</div>', unsafe_allow_html=True)
    if not clientes_db:
        st.sidebar.warning("No existen contribuyentes registrados. Ingrese uno antes de continuar.")
    else:
        cliente_sel = st.sidebar.selectbox("Contribuyente:", options=list(dic_clientes.keys()), format_func=lambda x: dic_clientes[x])
        tipo_ob = st.sidebar.selectbox("Entidad / Tipo:", ["SUNAT", "AFP"])
        situacion_ob = st.sidebar.selectbox("Estado de Declaración:", ["Declarado", "Por Declarar", "Fraccionado"])
        periodo_ob = st.sidebar.text_input("Periodo Tributario (AAAA-MM):", placeholder="Ej: 2026-07")

        ruc_cliente_sel = [c[1] for c in clientes_db if c[0] == cliente_sel][0]
        venc_sugerido = calcular_vencimiento_sunat_aproximado(ruc_cliente_sel, periodo_ob)

        venc_ob = st.sidebar.text_input("Fecha de Vencimiento (AAAA-MM-DD):", value=venc_sugerido)
        monto_ob = st.sidebar.number_input("Monto Determinado (S/):", min_value=0.0, step=10.0, format="%.2f")
        pago_ob = st.sidebar.date_input("Fecha de Pago Ejecutado (Opcional):", value=None, format="DD/MM/YYYY")
        tributos_ob = st.sidebar.text_input("Descripción / Código de Tributo:")

        if st.sidebar.button("Guardar Obligación", use_container_width=True):
            if not validar_periodo(periodo_ob):
                st.sidebar.error("El periodo debe mantener la estructura AAAA-MM.")
            elif monto_ob <= 0:
                st.sidebar.error("El monto ingresado debe ser mayor a 0.00.")
            else:
                estado_ob = "PAGADO" if pago_ob else "PENDIENTE"
                fecha_pago_str = pago_ob.strftime("%Y-%m-%d") if pago_ob else None
                conn = conectar_bd()
                c = conn.cursor()
                c.execute("""INSERT INTO obligaciones (cliente_id, tipo, periodo, monto, fecha_vencimiento, estado, fecha_pago, situacion, descripcion_tributos)
                             VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                          (cliente_sel, tipo_ob, periodo_ob, monto_ob, venc_ob, estado_ob, fecha_pago_str, situacion_ob, tributos_ob))
                conn.commit()
                conn.close()
                st.sidebar.success("Obligación financiera registrada en el sistema.")
                st.rerun()

# PANEL SUPERIOR DE FILTROS
st.markdown('<div class="subtitulo-corporativo">Filtros de Búsqueda y Consulta</div>', unsafe_allow_html=True)
col_filtro1, col_filtro2, col_filtro3 = st.columns([2, 2, 1])

with col_filtro1:
    txt_buscar = st.text_input("Búsqueda por Razón Social o RUC:", placeholder="Ingrese texto de búsqueda...")
with col_filtro2:
    opciones_filtro_cliente = {0: "TODOS LOS CONTRIBUYENTES"}
    opciones_filtro_cliente.update(dic_clientes)
    cliente_filtro_id = st.selectbox("Filtrar por Contribuyente:", options=list(opciones_filtro_cliente.keys()), format_func=lambda x: opciones_filtro_cliente[x])
with col_filtro3:
    periodo_filtro = st.text_input("Filtrar por Periodo:", placeholder="AAAA-MM")

# SECCIONES PRINCIPALES
tab_pendientes, tab_pagados = st.tabs(["Obligaciones Pendientes", "Historial de Obligaciones Pagadas"])

# ----------------- TABLA DE PENDIENTES -----------------
with tab_pendientes:
    conn = conectar_bd()
    c = conn.cursor()
    query_pend = """
    SELECT o.id, c.razon_social, c.ruc, c.regimen, o.tipo, o.situacion, o.periodo, o.monto, o.fecha_vencimiento, o.descripcion_tributos
    FROM obligaciones o JOIN clientes c ON o.cliente_id = c.id WHERE o.estado = 'PENDIENTE'
    """
    params_p = []
    if txt_buscar:
        query_pend += " AND (c.razon_social LIKE ? OR c.ruc LIKE ?)"
        params_p.extend([f"%{txt_buscar}%", f"%{txt_buscar}%"])
    if cliente_filtro_id != 0:
        query_pend += " AND o.cliente_id = ?"
        params_p.append(cliente_filtro_id)
    if periodo_filtro:
        query_pend += " AND o.periodo LIKE ?"
        params_p.append(f"%{periodo_filtro}%")
    query_pend += " ORDER BY o.fecha_vencimiento ASC"

    c.execute(query_pend, params_p)
    filas_pend = c.fetchall()
    conn.close()

    if filas_pend:
        data_pend = []
        hoy = datetime.now().date()
        limite = hoy + timedelta(days=3)

        for f in filas_pend:
            alerta = "En Plazo"
            if f[8]:
                try:
                    f_venc = datetime.strptime(f[8], "%Y-%m-%d").date()
                    if f_venc < hoy:
                        alerta = "Vencido"
                    elif hoy <= f_venc <= limite:
                        alerta = "Próximo a Vencer"
                except ValueError:
                    pass

            data_pend.append({
                "ID": f[0], "Estatus": alerta, "Razón Social": f[1], "RUC": f[2], "Régimen": f[3],
                "Tipo": f[4], "Situación": f[5], "Periodo": f[6], "Monto (S/)": f"{f[7]:.2f}",
                "Vencimiento": f[8] or "", "Detalle / Tributos": f[9] or ""
            })

        st.dataframe(data_pend, use_container_width=True)

        col_accion1, col_accion2 = st.columns([1, 1])
        with col_accion1:
            st.markdown('<div class="subtitulo-corporativo">Procesar Pago de Obligación</div>', unsafe_allow_html=True)
            id_pagar = st.number_input("ID de Obligación a Liquidar:", min_value=1, step=1)
            voucher_file = st.file_uploader("Adjuntar Comprobante de Pago (Imagen/PDF):", type=["png", "jpg", "jpeg", "pdf"])

            if st.button("Confirmar Pago", use_container_width=True):
                conn = conectar_bd()
                c = conn.cursor()
                c.execute("SELECT id FROM obligaciones WHERE id=? AND estado='PENDIENTE'", (id_pagar,))
                if c.fetchone():
                    ruta_v = guardar_voucher_local(voucher_file, id_pagar)
                    hoy_str = datetime.now().strftime("%Y-%m-%d")
                    c.execute("UPDATE obligaciones SET estado='PAGADO', fecha_pago=?, ruta_voucher=? WHERE id=?", (hoy_str, ruta_v, id_pagar))
                    conn.commit()
                    st.success(f"La obligación con ID {id_pagar} fue actualizada a estado PAGADO.")
                    conn.close()
                    st.rerun()
                else:
                    st.error("El ID especificado no corresponde a una obligación pendiente válida.")
                    conn.close()

        with col_accion2:
            st.markdown('<div class="subtitulo-corporativo">Eliminación de Registro</div>', unsafe_allow_html=True)
            id_elim = st.number_input("ID de Obligación a Remover:", min_value=1, step=1)
            if st.button("Eliminar Registro", use_container_width=True):
                conn = conectar_bd()
                c = conn.cursor()
                c.execute("DELETE FROM obligaciones WHERE id=?", (id_elim,))
                conn.commit()
                conn.close()
                st.warning(f"La obligación con ID {id_elim} ha sido removida del sistema.")
                st.rerun()

        if OPENPYXL_DISPONIBLE:
            st.markdown("---")
            bytes_excel = generar_excel_bytes(False, txt_buscar, cliente_filtro_id, periodo_filtro)
            st.download_button("Exportar Reporte de Pendientes (Excel)", data=bytes_excel, file_name=f"Reporte_Pendientes_{datetime.now().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.info("No se registraron obligaciones pendientes con los criterios seleccionados.")

# ----------------- TABLA DE PAGADOS -----------------
with tab_pagados:
    conn = conectar_bd()
    c = conn.cursor()
    query_pag = """
    SELECT o.id, c.razon_social, c.ruc, c.regimen, o.tipo, o.situacion, o.periodo, o.monto, o.fecha_vencimiento, o.fecha_pago, o.ruta_voucher, o.descripcion_tributos
    FROM obligaciones o JOIN clientes c ON o.cliente_id = c.id WHERE o.estado = 'PAGADO'
    """
    params_pag = []
    if txt_buscar:
        query_pag += " AND (c.razon_social LIKE ? OR c.ruc LIKE ?)"
        params_pag.extend([f"%{txt_buscar}%", f"%{txt_buscar}%"])
    if cliente_filtro_id != 0:
        query_pag += " AND o.cliente_id = ?"
        params_pag.append(cliente_filtro_id)
    if periodo_filtro:
        query_pag += " AND o.periodo LIKE ?"
        params_pag.append(f"%{periodo_filtro}%")
    query_pag += " ORDER BY o.fecha_pago DESC"

    c.execute(query_pag, params_pag)
    filas_pag = c.fetchall()
    conn.close()

    if filas_pag:
        data_pag = []
        for f in filas_pag:
            data_pag.append({
                "ID": f[0], "Razón Social": f[1], "RUC": f[2], "Régimen": f[3],
                "Tipo": f[4], "Situación": f[5], "Periodo": f[6], "Monto (S/)": f"{f[7]:.2f}",
                "Vencimiento": f[8] or "", "Fecha Pago": f[9] or "", "Comprobante": "Adjunto" if f[10] else "Sin Archivo",
                "Detalle / Tributos": f[11] or ""
            })

        st.dataframe(data_pag, use_container_width=True)

        col_v1, col_v2 = st.columns(2)
        with col_v1:
            st.markdown('<div class="subtitulo-corporativo">Visualización de Comprobante</div>', unsafe_allow_html=True)
            id_ver_v = st.number_input("ID de Registro a Consultar:", min_value=1, step=1, key="v_ver")
            if st.button("Consultar Comprobante", use_container_width=True):
                conn = conectar_bd()
                c = conn.cursor()
                c.execute("SELECT ruta_voucher FROM obligaciones WHERE id=?", (id_ver_v,))
                res = c.fetchone()
                conn.close()
                if res and res[0] and os.path.exists(res[0]):
                    if res[0].lower().endswith((".png", ".jpg", ".jpeg")):
                        st.image(res[0])
                    else:
                        st.info(f"Ruta de archivo digital: {res[0]}")
                else:
                    st.error("No se registra comprobante adjunto para el ID indicado.")

        with col_v2:
            st.markdown('<div class="subtitulo-corporativo">Actualización de Comprobante</div>', unsafe_allow_html=True)
            id_sub_v = st.number_input("ID de Registro Objetivo:", min_value=1, step=1, key="v_sub")
            v_tardio = st.file_uploader("Seleccionar Nuevo Archivo:", type=["png", "jpg", "jpeg", "pdf"], key="f_sub")
            if st.button("Actualizar Archivo", use_container_width=True):
                if v_tardio:
                    ruta_t = guardar_voucher_local(v_tardio, id_sub_v)
                    conn = conectar_bd()
                    c = conn.cursor()
                    c.execute("UPDATE obligaciones SET ruta_voucher=? WHERE id=?", (ruta_t, id_sub_v))
                    conn.commit()
                    conn.close()
                    st.success("Comprobante adjuntado correctamente.")
                    st.rerun()

        if OPENPYXL_DISPONIBLE:
            st.markdown("---")
            bytes_excel_p = generar_excel_bytes(True, txt_buscar, cliente_filtro_id, periodo_filtro)
            st.download_button("Exportar Historial de Pagados (Excel)", data=bytes_excel_p, file_name=f"Reporte_Pagados_{datetime.now().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.info("No se registran transacciones finalizadas en el historial.")
