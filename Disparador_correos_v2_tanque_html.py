import pandas as pd
import os
import win32com.client
from datetime import datetime, date, timedelta

# -----------------------------
# Archivos y rutas
# -----------------------------
EXCEL_FILE = os.path.join(os.getcwd(), "Reporte_Casos_Detallados.xlsx")
HTML_FILE_TEMPLATE = os.path.join(os.getcwd(), "avance_{local}.html")
EXCEL_LOCAL_TEMPLATE = os.path.join(os.getcwd(), "Reporte_{local}.xlsx")

# -----------------------------
# Activar/desactivar envío de correo
# -----------------------------
ENVIAR_CORREO = True  # True = enviar correo, False = solo generar archivos y gráficos

# -----------------------------
# Activar/desactivar reporte extra de Maura Román (Cynthia Li)
# -----------------------------
ENVIAR_REPORTE_CYNTHIA = False

# -----------------------------
# Parámetro de fecha (constante editable)
# Opciones: "hoy", "ayer" o una fecha exacta "YYYY-MM-DD"
# -----------------------------
FECHA = "hoy"

def resolver_fecha(valor):
    if valor is None or valor.lower() == "hoy":
        return date.today()
    if valor.lower() == "ayer":
        return date.today() - timedelta(days=1)
    return datetime.strptime(valor, "%Y-%m-%d").date()

fecha_reporte = resolver_fecha(FECHA)
fecha_texto = fecha_reporte.strftime("%d/%m/%Y")
es_hoy = fecha_reporte == date.today()

# -----------------------------
# Gráfico HTML: tanques de agua con escala de medición (25/50/75/100%).
# Se construye solo con tablas y colores de fondo (compatible con Outlook).
# -----------------------------
def generar_html_tanques(nombres, prod, meta, titulo, subtitulo=""):
    TANQUE_ALTO = 200   # alto del tanque en píxeles
    TANQUE_ANCHO = 74   # ancho del tanque en píxeles
    ALTO_ETIQUETAS = 64 # alto aprox. de las etiquetas bajo el tanque (para alinear la escala)

    # Columna de escala (100% arriba, 0% abajo)
    filas_escala = ""
    for etiqueta in ["100%", "75%", "50%", "25%"]:
        filas_escala += (f'<tr><td height="{TANQUE_ALTO // 4}" valign="top" align="right" '
                         f'style="font-family:Calibri;font-size:8pt;color:#8895A7;padding-right:5px;">'
                         f'{etiqueta} &#8212;</td></tr>')
    escala = (f'<td valign="bottom" style="padding:0 2px;">'
              f'<table cellpadding="0" cellspacing="0" border="0">{filas_escala}'
              f'<tr><td height="{ALTO_ETIQUETAS}"></td></tr></table></td>')

    columnas = ""
    for nombre, p, m in zip(nombres, prod, meta):
        p = int(p)
        m = int(m)
        pct = (p / m) if m else 0
        pct_txt = f"{round(pct * 100)}%"
        lleno = int(round(min(pct, 1.0) * TANQUE_ALTO))
        vacio = TANQUE_ALTO - lleno

        if pct >= 1:
            c_superficie, c_medio, c_fondo = "#C8F2D0", "#3CB35F", "#1E7E3C"  # verde: meta cumplida
            c_pct = "#1E7E3C"
        else:
            c_superficie, c_medio, c_fondo = "#D8F7F2", "#2BBBAD", "#0E8C80"  # turquesa: agua
            c_pct = "#0E8C80"

        sup = 4 if lleno > 8 else 0          # línea clara de superficie del agua
        fondo = int((lleno - sup) * 0.45)
        medio = lleno - sup - fondo

        filas = ""
        if vacio > 0:
            num_fuera = p if medio < 20 else ""
            filas += (f'<tr><td height="{vacio}" align="center" valign="bottom" '
                      f'style="font-family:Calibri;font-size:10pt;font-weight:bold;color:#23303E;">'
                      f'{num_fuera}</td></tr>')
        if sup > 0:
            filas += f'<tr><td height="{sup}" bgcolor="{c_superficie}"></td></tr>'
        if medio > 0:
            num_dentro = p if medio >= 20 else ""
            filas += (f'<tr><td height="{medio}" bgcolor="{c_medio}" align="center" '
                      f'style="font-family:Calibri;font-size:12pt;font-weight:bold;color:#FFFFFF;">'
                      f'{num_dentro}</td></tr>')
        if fondo > 0:
            filas += f'<tr><td height="{fondo}" bgcolor="{c_fondo}"></td></tr>'

        columnas += f'''
        <td align="center" valign="bottom" style="padding:0 9px;">
          <table cellpadding="0" cellspacing="0" border="0">
            <tr><td>
              <table width="{TANQUE_ANCHO}" cellpadding="0" cellspacing="0" border="0"
                     style="border:2px solid #5B7083;background-color:#ECF4F7;">
                {filas}
              </table>
            </td></tr>
            <tr><td align="center" style="font-family:Calibri;font-size:12pt;font-weight:bold;color:{c_pct};padding-top:5px;">{pct_txt}</td></tr>
            <tr><td align="center" style="font-family:Calibri;font-size:9pt;color:#23303E;">{p} de {m}</td></tr>
            <tr><td align="center" width="100" style="font-family:Calibri;font-size:8pt;color:#5B6B7C;padding-top:2px;">{nombre}</td></tr>
          </table>
        </td>'''

    return f'''
    <table cellpadding="0" cellspacing="0" border="0" style="border:1px solid #C9D4DE;background-color:#FFFFFF;">
      <tr>
        <td style="background-color:#23303E;padding:12px 18px;">
          <span style="font-family:Calibri;font-size:14pt;font-weight:bold;color:#FFFFFF;">{titulo}</span><br>
          <span style="font-family:Calibri;font-size:9pt;color:#9FB3C8;">{subtitulo}</span>
        </td>
      </tr>
      <tr>
        <td style="padding:18px 12px 12px 6px;">
          <table cellpadding="0" cellspacing="0" border="0"><tr>{escala}{columnas}</tr></table>
        </td>
      </tr>
      <tr>
        <td style="background-color:#F4F7FA;font-family:Calibri;font-size:8pt;color:#7F8C9B;padding:8px 18px;">
          Turquesa = avance &nbsp;&bull;&nbsp; Verde = meta cumplida &nbsp;&bull;&nbsp; El nivel del tanque indica el % de cumplimiento de la meta.
        </td>
      </tr>
    </table>'''

# -----------------------------
# Gerentes siempre en CC
# -----------------------------
GERENTES_CC = [
    "soporte.pyme@buro.com.pe; dherrera@buro.com.pe; marceo@buro.com.pe; gperez@buro.com.pe"
]

# -----------------------------
# Correos por LOCAL y SDV
# -----------------------------
CORREOS_LOCAL = {
    "LIMA 1": {
        "JZ": ["jbarahona.stbk@buro.com.pe"],
        "SDV": {
            "PEREZ MONTERO HELGAR GIOVANNI": "hperezm@buro.com.pe",
            "SORIA CARAZAS BRIZETH STEFANNY": "bsoriac@buro.com.pe",
            "YOMONA SOLIER MALENA": "myomonas.proveedorexterno@buro.com.pe",
            "DORADOR QUINTANA FREDDY RONALD": "fdoradorq@buro.com.pe"
        }
    },
    "LIMA 2": {
        "JZ": ["sbeteta@buro.com.pe"],
        "SDV": {
            "GARCIA MENDOZA ANTONIO ERNESTO": "agarciam@buro.com.pe",
            "JOHAN SAAVEDRA": "",
            "CENTENO GARCIA MIRIAM MARLENE": "mcentenog@buro.com.pe",
            "VIGO REYES BERTHA VALENTINA": "bvigor.proveedorexterno@buro.com.pe",
            "ACOSTA SEGOVIA ANA RITA": "aacostas@buro.com.pe",
            "ROSAS MENDOZA AURORA MARIA DEL ROSARIO": "mrosasm@buro.com.pe"
        }
    },
    "PB": {
        "JZ": ["jcolomad@buro.com.pe"],
        "SDV": {
            "MALPARTIDA MIRANDA ALFONSO": "amalpartidam@buro.com.pe",
            "SANTOS HEREDIA GIRALDO ERNESTO": "esantosh@buro.com.pe",
            "VACANTE": "",
            "MONTENEGRO REQUEJO ALBERTO ALEJANDRO": "amontenegror@buro.com.pe",
            "TAFUR RAMIREZ NOEMI MARITZA": "ntafurr.proveedorexterno@buro.com.pe",
            "VACANTE": "",
            "VACANTE": "",
            "HUAMAN ROMAN MAURA": "",
        }
    },
    "ZONA AREQUIPA": {
        "JZ": ["svillenas@buro.com.pe"],
        "SDV": {
            "DIAZ REVILLA CARMEN ROSA": "cdiazr@buro.com.pe",
            "FLORES SOSA YAMALY FABIOLA": "yfloress@buro.com.pe",
            "LOPEZ AMADO PATRICIA ELENA": "plopeza@buro.com.pe",
            "BUENO VALENCIA ROCIO": "rbuenov@buro.com.pe",
            "BAEZ RIVERA NELLY": "",
            "YANQUI APAZA YENI": "",
        }
    },
    "ZONA CENTRO": {
        "JZ": ["jvargas.stbk@buro.com.pe"],
        "SDV": {
            "TORRES CARDENAS ELIZABETH CLARISA": "etorresc@buro.com.pe",
            "MAMANI RICHARD ALFREDO HERNANDEZ": "rhernandezm@buro.com.pe"
        }
    },
    "ZONA CHICLAYO": {
        "JZ": ["gdavilas@buro.com.pe"],
        "SDV": {
            "OBLITAS ALVAREZ ALONSO LEANDRO": "",
        }
    },
    "ZONA CUSCO": {
        "JZ": ["asalazar.stbk@buro.com.pe"],
        "SDV": {
            "SALAZAR GAMARRA HAROLD YAMIL": "",
            "GARCIA MORENO IVAN TELMO": ""
        }
    },
    "ZONA ORIENTE": {
        "JZ": [""],
        "SDV": {
            "VING RAMIREZ REISER": "rvingr@buro.com.pe",
            "FACHIN RUIZ DAVID": "dfachinr@buro.com.pe",
            "CANCHANYA PAES JOSE LUIS": "jcanchanyap.proveedorexterno@buro.com.pe",
            "CRISPIN ALCANTARA JIMMY ERIK": "jcrispina@buro.com.pe",
            "MEDINA RIOS KARLA CECILIA": "kmedinar@buro.com.pe"

        }
    },
    "ZONA TACNA": {
        "JZ": ["lliendo.stbk@buro.com.pe"],
        "SDV": {
            "LIENDO CARPIO LUZ BERTA": "",
        }
    },
    "ZONA TRUJILLO": {
        "JZ": ["lrodriguezr@buro.com.pe"],
        "SDV": {
            "VARGAS PEREZ TEODORO LEONIDAS": "tvargasp.proveedorexterno@buro.com.pe",
            "MARIN LLAMOCTANTA CARLOS MALAQUIAS": "cmarinll@buro.com.pe",
            "VILLEGAS ARIAS ROBERTO": "rvillegasa@buro.com.pe",
            "VALERA VILLENA DIEGO ALONSO": "dvalerav@buro.com.pe"
        }
    }
}

# -----------------------------
# Leer Excel
# -----------------------------
df = pd.read_excel(EXCEL_FILE, dtype={"USUARIO VENDOR": str})

# -----------------------------
# Filtrar por fecha del reporte
# -----------------------------
df["Fecha Envío"] = pd.to_datetime(df["Fecha Envío"])
df = df[df["Fecha Envío"].dt.date == fecha_reporte]

# -----------------------------
# Eliminar duplicados por Nro Doc
# -----------------------------
df = df.drop_duplicates(subset=["Nro Doc"])

# -----------------------------
# Outlook (solo si se envían correos)
# -----------------------------
if ENVIAR_CORREO:
    outlook = win32com.client.Dispatch("Outlook.Application")

# -----------------------------
# Por cada LOCAL definido en CORREOS_LOCAL
# -----------------------------
for local, info_local in CORREOS_LOCAL.items():

    # Filtrar datos del local (puede estar vacío)
    df_local = df[df["LOCAL"].str.upper() == local.upper()]

    # Obtener destinatarios
    destinatarios = list(info_local["SDV"].values()) + info_local["JZ"]

    # Si no hay correos, avisar y continuar
    if not destinatarios:
        print(f"⚠ LOCAL {local} no tiene destinatarios definidos. Saltando...")
        continue

    # -----------------------------
    # Preparar datos para la gráfica
    # -----------------------------
    sdv_list = list(info_local["SDV"].keys())
    prod = []
    meta = []

    for sdv in sdv_list:
        # Contar consultas de este SDV (0 si no hay registros)
        asesores_sdv = df_local[df_local["SupervisorVenta"] == sdv]["USUARIO VENDOR"].unique()
        consultas_sdv = df_local[df_local["USUARIO VENDOR"].isin(asesores_sdv)]["Nro Doc"].count()
        prod.append(consultas_sdv)

        if sdv.upper() == "FACHIN RUIZ DAVID":
            meta.append(6)  # Meta por SDV para david
        else:
            meta.append(20)  # Meta por SDV para los demás

    # -----------------------------
    # Textos según la fecha (hoy = avance, otra fecha = cierre)
    # -----------------------------
    if es_hoy:
        titulo_grafico = f"Producción diaria por SDV - LOCAL {local}"
        asunto = f"Avance diario consultas Checkalo - LOCAL {local}"
        parrafo = f"Se envía el avance diario de consultas <b>Checkalo</b> para <b>{local}</b>:"
    else:
        titulo_grafico = f"Cierre diario por SDV - LOCAL {local} ({fecha_texto})"
        asunto = f"Cierre diario consultas Checkalo - LOCAL {local} ({fecha_texto})"
        parrafo = f"Se envía el <b>cierre diario</b> de consultas <b>Checkalo</b> correspondiente al día <b>{fecha_texto}</b> para <b>{local}</b>."

    # -----------------------------
    # Crear gráfico HTML (tanques con escala)
    # -----------------------------
    nombres_cortos = [" ".join(s.split()[:2]).title() for s in sdv_list]
    html_grafico = generar_html_tanques(nombres_cortos, prod, meta, titulo_grafico,
                                        f"Corte al {fecha_texto}")

    # Guardar vista previa HTML (para revisar el diseño en el navegador)
    HTML_FILE = HTML_FILE_TEMPLATE.format(local=local.replace(" ", "_"))
    with open(HTML_FILE, "w", encoding="utf-8") as f:
        f.write(html_grafico)

    # Guardar Excel (aunque df_local esté vacío)
    EXCEL_LOCAL = EXCEL_LOCAL_TEMPLATE.format(local=local.replace(" ", "_"))
    df_local.to_excel(EXCEL_LOCAL, index=False)

    # -----------------------------
    # Enviar correo (solo si ENVIAR_CORREO = True)
    # -----------------------------
    if ENVIAR_CORREO:
        mail_out = outlook.CreateItem(0)
        mail_out.To = ";".join(destinatarios)
        mail_out.CC = ";".join(GERENTES_CC)
        mail_out.Subject = asunto

        mail_out.HTMLBody = f"""
        <html>
        <body style="font-family:Calibri, sans-serif; font-size:11pt;">
        <p>Buenas tardes,</p>
        <p>{parrafo}</p>
        {html_grafico}
        <p>Saludos,<br>Robot Checkalo</p>
        </body>
        </html>
        """

        # Adjuntar Excel
        attachment_excel = mail_out.Attachments.Add(
            Source=EXCEL_LOCAL
        )

        mail_out.Send()
        print(f"📧 Correo enviado correctamente para LOCAL {local}")
    else:
        print(f"✅ Generado gráfico y Excel para LOCAL {local} (correo desactivado)")

    # -----------------------------------------------------------
    # ADICIONAL: REPORTE EXCLUSIVO PARA CYNTHIA LI (MAURA ROMÁN)
    # -----------------------------------------------------------
    if ENVIAR_REPORTE_CYNTHIA and "HUAMAN ROMAN MAURA" in info_local["SDV"]:
        # Filtrar solo los datos de Maura
        df_maura = df_local[df_local["SupervisorVenta"] == "HUAMAN ROMAN MAURA"]
        conteo_maura = df_maura.shape[0]

        # Generar Excel exclusivo de Maura
        EXCEL_MAURA = os.path.join(os.getcwd(), "Reporte_MAURA_ROMAN.xlsx")
        df_maura.to_excel(EXCEL_MAURA, index=False)

        # Generar mini-gráfico HTML para este correo extra
        html_maura = generar_html_tanques(["MAURA ROMÁN"], [conteo_maura], [20],
                                          "Checkalo Maura Román", f"Corte al {fecha_texto}")

        if ENVIAR_CORREO:
            mail_m = outlook.CreateItem(0)
            mail_m.To = "jcolomad@buro.com.pe; mhuamanr@buro.com.pe"
            mail_m.CC = "Cinthia.Li@scotiabank.com.pe; " + ";".join(GERENTES_CC)
            mail_m.Subject = f"REPORTE : MAURA HUAMAN - {fecha_texto}"
            mail_m.HTMLBody = f"""
            <html><body style="font-family:Calibri, sans-serif; font-size:11pt;">
            <p>Estimados,</p>
            <p>Se envía el conteo de consultas <b>Checkalo</b> de la supervisora <b>Maura Román</b> correspondiente al día <b>{fecha_texto}</b>.</p>
            <p><b>Total procesado: {conteo_maura}</b></p>
            {html_maura}
            <p>Saludos,<br>Robot Checkalo</p>
            </body></html>
            """
            # Adjuntar el Excel exclusivo de Maura
            mail_m.Attachments.Add(Source=EXCEL_MAURA)

            mail_m.Send()
            print(f"📧 Reporte extra de Maura Román enviado a Cynthia Li ({conteo_maura} casos)")
        else:
            print(f"✅ Generado reporte extra de Maura Román (Cynthia Li) - Correo desactivado")

print("--- Proceso completado ---")
