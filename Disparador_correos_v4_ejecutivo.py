import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import os
import win32com.client
from datetime import datetime, date, timedelta

# -----------------------------
# Archivos y rutas
# -----------------------------
EXCEL_FILE = os.path.join(os.getcwd(), "Reporte_Casos_Detallados.xlsx")
IMG_FILE_TEMPLATE = os.path.join(os.getcwd(), "avance_{local}.png")
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
# Gráfico ejecutivo: barras horizontales ordenadas por cumplimiento,
# con semáforo de colores (verde >=100%, ámbar >=70%, rojo <70%)
# -----------------------------
def dibujar_barras_ejecutivo(nombres, prod, meta, titulo, img_file):
    datos = sorted(zip(nombres, prod, meta),
                   key=lambda t: (t[1] / t[2]) if t[2] else 0, reverse=True)
    n = len(datos)
    fig, ax = plt.subplots(figsize=(11, max(4, 0.65 * n + 1.8)), constrained_layout=True)
    fig.patch.set_facecolor("white")
    posiciones = list(range(n - 1, -1, -1))

    for yi, (nom, p, m) in zip(posiciones, datos):
        p, m = int(p), int(m)
        pct = p / m if m else 0
        if pct >= 1:
            color = "#1E8E44"
        elif pct >= 0.7:
            color = "#E8A013"
        else:
            color = "#C0392B"

        # Pista de fondo (la meta completa) y barra de avance
        ax.barh(yi, 1.0, height=0.62, color="#EDF1F5", edgecolor="#D5DDE5",
                linewidth=1, zorder=1)
        if pct > 0:
            ax.barh(yi, min(pct, 1.0), height=0.62, color=color, zorder=2)

        ax.text(1.03, yi, f"{p} / {m}   ({round(pct * 100)}%)",
                ha="left", va="center", fontsize=10, fontweight="bold", color="#37474F")
        if pct >= 0.15:
            ax.text(min(pct, 1.0) - 0.02, yi, f"{round(pct * 100)}%",
                    ha="right", va="center", fontsize=9, fontweight="bold", color="white")

    # Línea de meta (100%)
    ax.axvline(1.0, color="#37474F", linewidth=1.4, linestyle=(0, (4, 3)), zorder=3)
    ax.text(1.0, n - 0.25, "Meta", ha="center", va="bottom",
            fontsize=9, fontweight="bold", color="#37474F")

    ax.set_yticks(posiciones)
    ax.set_yticklabels([d[0] for d in datos], fontsize=10)
    ax.set_xlim(0, 1.42)
    ax.set_ylim(-0.6, n - 0.1)
    ax.set_xticks([0, 0.25, 0.50, 0.75, 1.0])
    ax.set_xticklabels(["0%", "25%", "50%", "75%", "100%"], fontsize=9, color="#7F8C9B")
    for sp in ax.spines.values():
        sp.set_visible(False)
    ax.tick_params(length=0)
    ax.set_title(titulo, fontsize=14, fontweight="bold", color="#1F3864", pad=14, loc="left")
    fig.text(0.01, 0.005, "Verde = meta cumplida   ·   Ámbar = avance ≥ 70%   ·   Rojo = avance < 70%",
             fontsize=8, color="#7F8C9B")

    plt.savefig(img_file, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)

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
    # Crear gráfico (barras ejecutivas horizontales)
    # -----------------------------
    nombres_cortos = [" ".join(s.split()[:2]).title() for s in sdv_list]
    IMG_FILE = IMG_FILE_TEMPLATE.format(local=local.replace(" ", "_"))
    dibujar_barras_ejecutivo(nombres_cortos, prod, meta, titulo_grafico, IMG_FILE)

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
        <p><img src="cid:avance_{local.replace(' ','_')}"></p>
        <p>Saludos,<br>Robot Checkalo</p>
        </body>
        </html>
        """

        # Adjuntar imagen
        attachment_img = mail_out.Attachments.Add(
            Source=IMG_FILE,
            Type=1,
            Position=0
        )
        attachment_img.PropertyAccessor.SetProperty(
            "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
            f"avance_{local.replace(' ','_')}"
        )

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

        # Generar mini-gráfico para este correo extra
        fig_m, ax_m = plt.subplots(figsize=(4, 4))
        ax_m.bar(["MAURA ROMÁN"], [conteo_maura], color="#4F81BD")
        ax_m.set_title(f"Checkalo Maura Román - {fecha_texto}")
        ax_m.set_ylim(0, max(conteo_maura + 5, 25))
        for i, v in enumerate([conteo_maura]):
            ax_m.text(i, v + 0.5, str(v), ha='center', fontweight='bold')

        IMG_MAURA = os.path.join(os.getcwd(), "avance_MAURA_INDIVIDUAL.png")
        plt.savefig(IMG_MAURA, dpi=100, bbox_inches='tight')
        plt.close(fig_m)

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
            <img src="cid:img_m_indiv">
            <p>Saludos,<br>Robot Checkalo</p>
            </body></html>
            """
            att_m = mail_m.Attachments.Add(Source=IMG_MAURA, Type=1, Position=0)
            att_m.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "img_m_indiv")

            # Adjuntar el Excel exclusivo de Maura
            mail_m.Attachments.Add(Source=EXCEL_MAURA)

            mail_m.Send()
            print(f"📧 Reporte extra de Maura Román enviado a Cynthia Li ({conteo_maura} casos)")
        else:
            print(f"✅ Generado reporte extra de Maura Román (Cynthia Li) - Correo desactivado")

print("--- Proceso completado ---")
