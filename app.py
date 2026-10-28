import os
import io
import base64
import hashlib
import re
from datetime import datetime
from zoneinfo import ZoneInfo
from flask import Flask, render_template, request, redirect, url_for, Response, render_template_string

# Supabase y utilidades
from supabase import create_client, Client

# Lectura de Excel para trabajadores
import pandas as pd

# Generación de PDF (ReportLab)
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# Generación de Word (python-docx)
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

app = Flask(__name__)

# Configuración de Supabase
SUPABASE_URL = os.environ.get("SUPABASE_URL")
  

if not SUPABASE_URL or not SUPABASE_KEY:
    raise RuntimeError(
        "Faltan las variables de entorno SUPABASE_URL y/o SUPABASE_KEY. "
        "Configúralas antes de arrancar la app."
    )

supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

TZ_MADRID = ZoneInfo("Europe/Madrid")




# variale URL

# URL base oficial leída estrictamente desde Render
BASE_URL_PROD = os.environ.get("BASE_URL")
if not BASE_URL_PROD:
    raise RuntimeError(
        "Falta la variable de entorno BASE_URL. "
        "Configúrala en Render antes de arrancar la app."
    )

# URL base oficial leída desde las variables de entorno de Render
#BASE_URL_PROD = os.environ.get("BASE_URL", "https://app-firmas-cae.onrender.com")

SECRET_SALT = os.environ.get("SECRET_SALT")
if not SECRET_SALT:
    raise RuntimeError(
        "Falta la variable de entorno SECRET_SALT. Configúrala antes de arrancar la app."
    )

# URL base oficial de producción en Render
BASE_URL_PROD = "https://app-firmas-cae.onrender.com"

def get_current_month_name():
    months = {
        1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
        5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
        9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre",
    }
    now = datetime.now(TZ_MADRID)
    return f"{months[now.month]} {now.year}"

def generar_token(dni, mes):
    """Genera un token seguro y único para el DNI y mes sin exponer el DNI en la URL."""
    raw = f"{dni}_{mes}_{SECRET_SALT}"
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()[:24]

def enviar_otp_placeholder(destino, codigo):
    print(f"[OTP pendiente de implementar] Se enviaría el código {codigo} a {destino}")

def enmascarar_dni(dni):
    """Oculta los caracteres centrales del DNI con asteriscos (ej: 288****2q)."""
    dni = str(dni).strip()
    if len(dni) > 5:
        return dni[:3] + "****" + dni[-2:]
    return dni

# Ruta raíz redirige al panel de control de firmas
@app.route("/")
def home():
    return redirect(url_for('ver_firmas'))

# Ruta de privacidad
@app.route("/privacidad")
def privacidad():
    try:
        with open("privacidad.txt", "r", encoding="utf-8") as f:
            contenido = f.read().split("\n---\n")[0]
    except Exception as e:
        contenido = "No se ha podido cargar la información de privacidad en este momento."
        print(f"Error leyendo privacidad.txt: {e}")

    return render_template_string("""
        <!DOCTYPE html>
        <html lang="es">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Protección de Datos</title>
            <style>
                body { font-family: Arial, sans-serif; background-color: #f8f9fa; color: #333; padding: 20px; }
                .card { background: white; padding: 25px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); max-width: 650px; margin: 0 auto; white-space: pre-wrap; line-height: 1.5; font-size: 14px; }
            </style>
        </head>
        <body>
            <div class="card">{{ contenido }}</div>
        </body>
        </html>
    """, contenido=contenido)

# Ruta para que los trabajadores firmen usando un token seguro
@app.route("/firmar/<token>", methods=["GET", "POST"])
def firmar(token):
    selected_month = request.args.get("mes", get_current_month_name())
    print(f"Token recibido en URL: {token} | Mes: {selected_month}")

    excel_path = "trabajadores_con_telefono.xlsx"
    trabajador_encontrado = None
    dni_encontrado = None

    if os.path.exists(excel_path):
        try:
            df_trab = pd.read_excel(excel_path)
            for _, row in df_trab.iterrows():
                t_dni = str(row.get("dni", "")).strip()
                token_calculado = generar_token(t_dni, selected_month)
                
                if token_calculado == token:
                    trabajador_encontrado = str(row.get("nombre", "")).strip()
                    dni_encontrado = t_dni
                    break
        except Exception as e:
            print(f"Error leyendo Excel: {e}")

    if not trabajador_encontrado or not dni_encontrado:
        return "Enlace de firma inválido o caducado.", 404

    if request.method == "POST":
        firma_data = request.form.get("firma")
        if not firma_data:
            return "No se ha capturado ninguna firma.", 400

        ip = request.headers.get("X-Forwarded-For", request.remote_addr)
        user_agent = request.headers.get("User-Agent", "")
        fecha_str = datetime.now(TZ_MADRID).strftime("%Y-%m-%d %H:%M:%S")

        campos_hash = [
            trabajador_encontrado,
            dni_encontrado,
            selected_month,
            fecha_str,
            ip,
            user_agent,
            firma_data,
        ]
        hash_input = "|".join(campos_hash).encode("utf-8")
        hash_sha256 = hashlib.sha256(hash_input).hexdigest()

        try:
            supabase.table("registro_firmas").insert({
                "fecha_utc": fecha_str,
                "mes": selected_month,
                "nombre": trabajador_encontrado,
                "dni": dni_encontrado,
                "firma_base64": firma_data,
                "ip": ip,
                "user_agent": user_agent,
                "hash_sha256": hash_sha256,
            }).execute()
        except Exception as e:
            print(f"Error guardando en Supabase: {e}")
            return f"Error al guardar la firma en la base de datos: {e}", 500

        return render_template_string("""
            <!DOCTYPE html>
            <html lang="es">
            <head>
                <meta charset="UTF-8">
                <meta name="viewport" content="width=device-width, initial-scale=1.0">
                <title>Firma Registrada</title>
                <style>
                    body { font-family: Arial, sans-serif; text-align: center; padding: 20px; background-color: #f8f9fa; color: #333; }
                    .card { background: white; padding: 30px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); display: inline-block; max-width: 500px; width: 100%; box-sizing: border-box; }
                    h2 { color: #27ae60; }
                </style>
            </head>
            <body>
                <div class="card">
                    <h2>¡Firma registrada con éxito!</h2>
                    <p>Gracias, <strong>{{ nombre }}</strong>. Tu conformidad de obligaciones salariales y de seguridad social (CAE) para el mes de <strong>{{ mes }}</strong> ha sido registrada correctamente.</p>
                </div>
            </body>
            </html>
        """, nombre=trabajador_encontrado, mes=selected_month)

    firmar_template = """
    <!DOCTYPE html>
    <html lang="es">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
        <title>Firma CAE Obligaciones Salariales - {{ nombre }}</title>
        <script src="https://cdn.jsdelivr.net/npm/signature_pad@4.1.7/dist/signature_pad.umd.min.js"></script>
        <style>
            * { box-sizing: border-box; }
            body { font-family: Arial, sans-serif; margin: 0; padding: 12px; background-color: #f8f9fa; color: #333; display: flex; justify-content: center; align-items: center; min-height: 100vh; }
            .container { background: white; padding: 18px; border-radius: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.1); width: 100%; max-width: 500px; text-align: center; }
            h2 { color: #1a365d; margin-top: 0; margin-bottom: 5px; font-size: 20px; }
            p { font-size: 13px; color: #555; margin-bottom: 12px; line-height: 1.4; text-align: left; }
            .legal-box { background: #f1f4f6; border-left: 4px solid #1a365d; padding: 10px 12px; margin-bottom: 15px; text-align: left; border-radius: 4px; font-size: 12px; }
            .pad-wrapper { position: relative; width: 100%; height: 280px; margin-bottom: 15px; touch-action: none; }
            .signature-pad { border: 2px dashed #bdc3c7; border-radius: 8px; background-color: #fff; width: 100%; height: 100%; cursor: crosshair; touch-action: none; }
            .buttons { display: flex; justify-content: space-between; gap: 10px; }
            button, .btn-clear { padding: 14px 15px; font-size: 15px; font-weight: bold; border-radius: 6px; border: none; cursor: pointer; width: 50%; }
            .btn-clear { background-color: #e74c3c; color: white; }
            .btn-submit { background-color: #27ae60; color: white; }
            button:active { opacity: 0.8; }
            
            /* Estilos para la barra continua de grosor */
            .slider-container { display: flex; align-items: center; justify-content: center; gap: 10px; margin-bottom: 12px; }
            .slider-container label { font-size: 12px; color: #1a365d; font-weight: bold; }
            .slider-container input[type="range"] { width: 65%; cursor: pointer; }
        </style>
    </head>
    <body>
        <div class="container">
            <h2>Firma CAE: {{ nombre }}</h2>
            <p style="text-align: center; color: #666; margin-bottom: 10px;">Periodo: <strong>{{ mes }}</strong></p>
            
            <div class="legal-box">
                <p style="margin:0;"><strong>Aviso CAE:</strong> Declaro estar informado y conforme con la puesta a disposición de los recibos de salarios y boletines de cotización a la Seguridad Social correspondientes al periodo indicado.</p>
                <p style="margin:8px 0 0 0;"><a href="/privacidad" target="_blank" style="color:#1a365d;">📄 Ver información sobre el tratamiento de tus datos</a></p>
            </div>
            
            <!-- Control deslizante continuo para ajustar el grosor del trazo -->
            <div class="slider-container">
                <label for="thickness">Grosor de trazo:</label>
                <input type="range" id="thickness" min="0.3" max="3.0" step="0.1" value="0.8" oninput="cambiarGrosor(this.value)">
            </div>

            <form method="POST" id="firma-form">
                <div class="pad-wrapper">
                    <canvas id="signature-pad" class="signature-pad"></canvas>
                </div>
                <input type="hidden" name="firma" id="firma-input">
                <div class="buttons">
                    <button type="button" class="btn-clear" id="clear">Borrar</button>
                    <button type="submit" class="btn-submit">Guardar Firma</button>
                </div>
            </form>
        </div>

        <script>
            const canvas = document.getElementById('signature-pad');
            
            // Inicialización de la firma con trazo por defecto (0.8)
            const signaturePad = new SignaturePad(canvas, {
                minWidth: 0.8,
                maxWidth: 1.4,
                penColor: "rgb(10, 25, 80)",
                throttle: 0
            });

            // Ajusta el grosor en tiempo real según el deslizador
            function cambiarGrosor(val) {
                const valor = parseFloat(val);
                signaturePad.minWidth = valor;
                signaturePad.maxWidth = valor + 0.6;
            }

            function resizeCanvas() {
                const ratio = Math.max(window.devicePixelRatio || 1, 1);
                canvas.width = canvas.offsetWidth * ratio;
                canvas.height = canvas.offsetHeight * ratio;
                canvas.getContext("2d").scale(ratio, ratio);
                signaturePad.clear();
            }

            window.addEventListener("resize", resizeCanvas);
            resizeCanvas();

            document.getElementById('clear').addEventListener('click', function () {
                signaturePad.clear();
            });

            document.getElementById('firma-form').addEventListener('submit', function (e) {
                if (signaturePad.isEmpty()) {
                    alert("Por favor, dibuje su firma antes de enviar.");
                    e.preventDefault();
                } else {
                    document.getElementById('firma-input').value = signaturePad.toDataURL();
                }
            });
        </script>
    </body>
    </html>
    """
    return render_template_string(
        firmar_template, nombre=trabajador_encontrado, mes=selected_month
    )

# Panel de Control Principal
@app.route("/ver_firmas", methods=["GET"])
def ver_firmas():
    default_month = get_current_month_name()
    selected_month = request.args.get("mes", default_month)

    excel_path = "trabajadores_con_telefono.xlsx"
    trabajadores = []
    if os.path.exists(excel_path):
        try:
            df_trab = pd.read_excel(excel_path)
            for _, row in df_trab.iterrows():
                nombre = str(row.get("nombre", "")).strip()
                dni = str(row.get("dni", "")).strip()
                telefono = str(row.get("telefono", "")).strip()
                tel_limpio = re.sub(r"\D", "", telefono)
                
                token = generar_token(dni, selected_month)
                dni_mascarado = enmascarar_dni(dni)

                trabajadores.append({
                    "nombre": nombre,
                    "dni": dni,
                    "dni_mascarado": dni_mascarado,
                    "telefono": telefono,
                    "tel_limpio": tel_limpio,
                    "token": token
                })
        except Exception as e:
            print(f"Error leyendo Excel de trabajadores: {e}")

    try:
        response = (
            supabase.table("registro_firmas")
            .select("id, fecha_utc, mes, nombre, dni, firma_base64, ip, hash_sha256")
            .eq("mes", selected_month)
            .order("id", desc=True)
            .execute()
        )
        registros = response.data if response.data else []
    except Exception as e:
        print(f"Error consultando Supabase: {e}")
        registros = []

    try:
        all_res = supabase.table("registro_firmas").select("mes").execute()
        meses_set = set(r["mes"] for r in all_res.data if r.get("mes"))
        meses_set.add(default_month)
        meses_disponibles = sorted(list(meses_set), reverse=True)
    except:
        meses_disponibles = [default_month]

    html_template = """
    <!DOCTYPE html>
    <html lang="es">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Obligaciones Salariales y Seguros Sociales (CAE) - {{ selected_month }}</title>
        <style>
            body { font-family: Arial, sans-serif; margin: 20px; background-color: #f8f9fa; color: #333; }
            h2, h3 { color: #1a365d; margin-bottom: 10px; }
            p.intro { font-size: 14px; color: #555; margin-bottom: 20px; line-height: 1.5; }
            .filter-container { background: #fff; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); margin-bottom: 25px; display: flex; flex-wrap: wrap; justify-content: space-between; align-items: center; gap: 15px; }
            .filter-left { display: flex; align-items: center; gap: 10px; }
            .filter-right { display: flex; gap: 10px; }
            select, button { padding: 8px 12px; font-size: 14px; border-radius: 4px; border: 1px solid #ccc; }
            .btn-pdf { background-color: #e74c3c; color: white; text-decoration: none; padding: 9px 15px; border-radius: 4px; font-weight: bold; font-size: 13px; display: inline-block; }
            .btn-word { background-color: #2980b9; color: white; text-decoration: none; padding: 9px 15px; border-radius: 4px; font-weight: bold; font-size: 13px; display: inline-block; }
            .btn-whatsapp { background-color: #27ae60; color: white; text-decoration: none; padding: 6px 12px; border-radius: 4px; font-size: 12px; font-weight: bold; display: inline-block; }
            table { width: 100%; border-collapse: collapse; background: white; border-radius: 8px; overflow: hidden; box-shadow: 0 2px 4px rgba(0,0,0,0.1); margin-bottom: 30px; }
            th, td { padding: 10px 12px; text-align: left; border-bottom: 1px solid #eee; font-size: 13px; }
            th { background-color: #1a365d; color: white; }
            tr:hover { background-color: #f1f1f1; }
            img.firma-img { max-height: 45px; background: #fff; padding: 2px; border: 1px solid #ddd; border-radius: 3px; }
            .section-title { margin-top: 40px; border-bottom: 2px solid #1a365d; padding-bottom: 5px; }
        </style>
    </head>
    <body>
        <h2>Control de Cumplimiento Salarial y Seguridad Social ({{ selected_month }})</h2>
        <p class="intro">
            Panel oficial de verificación para la <strong>Coordinación de Actividades Empresariales (CAE)</strong>.
        </p>
        
        <div class="filter-container">
            <div class="filter-left">
                <form method="GET" action="/ver_firmas" style="margin: 0; display: flex; align-items: center; gap: 8px;">
                    <label for="mes"><strong>Filtrar por Mes:</strong></label>
                    <select name="mes" id="mes" onchange="this.form.submit()">
                        {% for m in meses_disponibles %}
                            <option value="{{ m }}" {% if m == selected_month %}selected{% endif %}>{{ m }}</option>
                        {% endfor %}
                    </select>
                </form>
            </div>
            <div class="filter-right">
                <a href="/descargar_pdf?mes={{ selected_month }}" class="btn-pdf">📄 Descargar Certificado PDF</a>
                <a href="/descargar_word?mes={{ selected_month }}" class="btn-word">📝 Descargar Certificado Word</a>
            </div>
        </div>

        <table>
            <thead>
                <tr>
                    <th>Fecha (Madrid)</th>
                    <th>Mes</th>
                    <th>Nombre</th>
                    <th>DNI</th>
                    <th>Firma</th>
                    <th>IP</th>
                    <th>Hash SHA-256 (Integridad)</th>
                </tr>
            </thead>
            <tbody>
                {% if registros %}
                    {% for reg in registros %}
                    <tr>
                        <td>{{ reg.fecha_utc }}</td>
                        <td>{{ reg.mes }}</td>
                        <td>{{ reg.nombre }}</td>
                        <td>{{ reg.dni }}</td>
                        <td>
                            {% if reg.firma_base64 %}
                                {% set f_val = reg.firma_base64 | string %}
                                {% if f_val.startswith('data:image') %}
                                    <img src="{{ f_val }}" class="firma-img" alt="Firma digitalizada">
                                {% else %}
                                    <span style="font-size: 11px; color: #27ae60; font-weight: bold;">[Firma OK]</span>
                                {% endif %}
                            {% else %}
                                <span style="color: #e74c3c;">Sin firma</span>
                            {% endif %}
                        </td>
                        <td>{{ reg.ip }}</td>
                        <td style="font-family: monospace; font-size: 11px; color: #555;">{{ reg.hash_sha256 }}</td>
                    </tr>
                    {% endfor %}
                {% else %}
                    <tr>
                        <td colspan="7" style="text-align: center; color: #777; padding: 20px;">No hay registros de firmas para este mes.</td>
                    </tr>
                {% endif %}
            </tbody>
        </table>

        <h3 class="section-title">📤 Enviar Enlaces de Firma por WhatsApp (Periodo: {{ selected_month }})</h3>
        
        <table>
            <thead>
                <tr>
                    <th>Trabajador</th>
                    <th>DNI (Enmascarado)</th>
                    <th>Teléfono</th>
                    <th>Acción WhatsApp</th>
                </tr>
            </thead>
            <tbody>
                {% for t in trabajadores %}
                <tr>
                    <td>{{ t.nombre }}</td>
                    <td style="font-family: monospace; font-weight: bold;">{{ t.dni_mascarado }}</td>
                    <td>{{ t.telefono }}</td>
                    <td>
                        {% if t.tel_limpio and t.dni %}
                            {% set link_firma = base_url_prod ~ "/firmar/" ~ t.token ~ "?mes=" ~ (selected_month | urlencode) %}
                            {% set texto_wa = 'Hola ' ~ t.nombre ~ ', por favor firma el certificado de obligaciones salariales y seguridad social CAE correspondiente al mes de ' ~ selected_month ~ ' (DNI: ' ~ t.dni_mascarado ~ ') accediendo al siguiente enlace: ' ~ link_firma %}
                            
                            <a href="https://api.whatsapp.com/send?phone={{ t.tel_limpio }}&text={{ texto_wa | urlencode }}" 
                               target="_blank" class="btn-whatsapp">
                                📤 Enviar WhatsApp (DNI: {{ t.dni_mascarado }})
                            </a>
                        {% else %}
                            <span style="color: #e74c3c; font-size: 12px;">Sin teléfono o DNI válido</span>
                        {% endif %}
                    </td>
                </tr>
                {% endfor %}
            </tbody>
        </table>
    </body>
    </html>
    """

    return render_template_string(
        html_template,
        registros=registros,
        selected_month=selected_month,
        meses_disponibles=meses_disponibles,
        trabajadores=trabajadores,
        base_url_prod=BASE_URL_PROD,
    )

@app.route("/descargar_pdf", methods=["GET"])
def descargar_pdf():
    selected_month = request.args.get("mes", get_current_month_name())
    try:
        response = (
            supabase.table("registro_firmas")
            .select("*")
            .eq("mes", selected_month)
            .order("id", desc=True)
            .execute()
        )
        registros = response.data if response.data else []
    except:
        registros = []

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=A4, rightMargin=20, leftMargin=20, topMargin=45, bottomMargin=30
    )
    elements = []
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "TitleStyle",
        parent=styles["Heading1"],
        fontSize=11,
        textColor=colors.HexColor("#1a365d"),
        spaceAfter=2,
        alignment=1
    )
    subtitle_style = ParagraphStyle(
        "SubTitleStyle",
        parent=styles["Normal"],
        fontSize=8,
        textColor=colors.HexColor("#4a5568"),
        spaceAfter=10,
        alignment=1,
        leading=10
    )

    elements.append(Paragraph("CERTIFICADO DE CUMPLIMIENTO DE OBLIGACIONES SALARIALES Y DE SEGURIDAD SOCIAL", title_style))
    elements.append(Paragraph(
        f"<b>Coordinación de Actividades Empresariales (CAE)</b> — Periodo: {selected_month}<br/>"
        "Documento acreditativo de puesta a disposición, verificación y conformidad de la documentación laboral, "
        "recibos de salarios y boletines de cotización.",
        subtitle_style
    ))
    elements.append(Spacer(1, 5))

    table_data = [["Fecha (Madrid)", "Nombre", "DNI", "IP", "Hash SHA-256 (Completo)", "Firma"]]
    
    for reg in registros:
        firma_b64 = reg.get("firma_base64", "")
        firma_element = Paragraph("Sin firma", styles["Normal"])
        
        if firma_b64:
            try:
                if ',' in firma_b64:
                    _, encoded_data = firma_b64.split(',', 1)
                else:
                    encoded_data = firma_b64
                img_bytes = base64.b64decode(encoded_data)
                img_io = io.BytesIO(img_bytes)
                firma_element = RLImage(img_io, width=70, height=30)
            except Exception:
                firma_element = Paragraph("[Error Firma]", styles["Normal"])

        table_data.append([
            Paragraph(str(reg.get("fecha_utc", "")), styles["Normal"]),
            Paragraph(str(reg.get("nombre", "")), styles["Normal"]),
            Paragraph(str(reg.get("dni", "")), styles["Normal"]),
            Paragraph(str(reg.get("ip", "")), styles["Normal"]),
            Paragraph(f"<font size=5.5>{str(reg.get('hash_sha256', ''))}</font>", styles["Normal"]),
            firma_element,
        ])

    t = Table(table_data, colWidths=[75, 95, 60, 80, 160, 85])
    t.setStyle(
        TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1a365d")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
            ("ALIGN", (0, 0), (-1, -1), "LEFT"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("BOTTOMPADDING", (0, 0), (-1, 0), 6),
            ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#f8f9fa")),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#dddddd")),
        ])
    )
    elements.append(t)

    def add_header_logo(canvas, document):
        canvas.saveState()
        if os.path.exists("logo.png"):
            try:
                canvas.drawImage("logo.png", 20, 812, width=50, height=22, preserveAspectRatio=True, mask='auto')
            except Exception as e:
                print(f"No se pudo cargar el logo en el PDF: {e}")
        canvas.restoreState()

    doc.build(elements, onFirstPage=add_header_logo, onLaterPages=add_header_logo)

    buffer.seek(0)
    return Response(
        buffer,
        mimetype="application/pdf",
        headers={
            "Content-Disposition": f"attachment; filename=certificado_obligaciones_salariales_{selected_month.replace(' ', '_')}.pdf"
        },
    )

@app.route("/descargar_word", methods=["GET"])
def descargar_word():
    selected_month = request.args.get("mes", get_current_month_name())
    try:
        response = (
            supabase.table("registro_firmas")
            .select("*")
            .eq("mes", selected_month)
            .order("id", desc=True)
            .execute()
        )
        registros = response.data if response.data else []
    except:
        registros = []

    doc = Document()
    
    if os.path.exists("logo.png"):
        try:
            section = doc.sections[0]
            header = section.header
            hp = header.paragraphs[0]
            hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            run_logo = hp.add_run()
            run_logo.add_picture("logo.png", width=Inches(1.2))
        except Exception as e:
            print(f"No se pudo añadir el logo en el Word: {e}")

    h = doc.add_heading("CERTIFICADO DE CUMPLIMIENTO DE OBLIGACIONES SALARIALES Y DE SEGURIDAD SOCIAL", level=1)
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    p_intro = doc.add_paragraph()
    p_intro.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_intro = p_intro.add_run(
        f"Coordinación de Actividades Empresariales (CAE) — Periodo: {selected_month}\n"
        "Acreditación formal de puesta a disposición y conformidad de la documentación laboral, salarial y de cotización."
    )
    run_intro.italic = True
    run_intro.font.size = Pt(10)
    
    doc.add_paragraph("―" * 50)

    if not registros:
        doc.add_paragraph("No hay registros de firmas para este periodo.")
    
    for reg in registros:
        doc.add_heading(f"Trabajador: {reg.get('nombre', '')} (DNI: {reg.get('dni', '')})", level=2)
        
        p_detalles = doc.add_paragraph()
        p_detalles.add_run("• Fecha y Hora de Validación (Madrid): ").bold = True
        p_detalles.add_run(f"{reg.get('fecha_utc', '')}\n")
        p_detalles.add_run("• Dirección IP de Origen: ").bold = True
        p_detalles.add_run(f"{reg.get('ip', '')}\n")
        p_detalles.add_run("• Hash SHA-256 de Integridad (Completo): ").bold = True
        
        run_hash = p_detalles.add_run(f"{reg.get('hash_sha256', '')}\n")
        run_hash.font.size = Pt(8.5)
        
        firma_b64 = reg.get("firma_base64", "")
        if firma_b64:
            try:
                if ',' in firma_b64:
                    _, encoded_data = firma_b64.split(',', 1)
                else:
                    encoded_data = firma_b64
                
                img_bytes = base64.b64decode(encoded_data)
                img_stream = io.BytesIO(img_bytes)
                
                p_firma_label = doc.add_paragraph()
                p_firma_label.add_run("Firma digitalizada de conformidad:").bold = True
                
                doc.add_picture(img_stream, width=Inches(2.2))
            except Exception as e:
                doc.add_paragraph(f"[No se pudo incrustar la imagen de la firma: {e}]")
        else:
            doc.add_paragraph("Firma: [Sin firma registrada]")
            
        doc.add_paragraph("―" * 40)

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)

    return Response(
        output,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={
            "Content-Disposition": f"attachment; filename=certificado_obligaciones_salariales_{selected_month.replace(' ', '_')}.docx"
        },
    )

# Ruta para el Cron Job
@app.route("/cron_ping")
def cron_ping():
    try:
        supabase.table("registro_firmas").select("id").limit(1).execute()
        return "CRON_OK", 200
    except Exception as e:
        print(f"Error en el cron ping: {e}")
        return f"Error: {e}", 500

if __name__ == "__main__":
    debug_mode = os.environ.get("FLASK_DEBUG", "false").lower() == "true"
    app.run(host="0.0.0.0", port=5000, debug=debug_mode)
