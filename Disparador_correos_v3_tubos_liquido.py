import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch
from matplotlib.colors import LinearSegmentedColormap
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
# Gráfico: tubos verticales con líquido degradado (azul = avance, verde = meta cumplida)
# -----------------------------
def dibujar_tubos_liquido(nombres, prod, meta, titulo, img_file):
    n = len(nombres)
    fig, ax = plt.subplots(figsize=(max(8, 1.5 * n), 6.5), constrained_layout=True)
    fig.patch.set_facecolor("white")
    ancho = 0.55

    for y in (0.25, 0.50, 0.75):
        ax.axhline(y, color="#E1E8F0", linewidth=1, linestyle="--", zorder=0)

    for i, (p, m) in enumerate(zip(prod, meta)):
        p, m = int(p), int(m)
        pct = p / m if m else 0
        h = min(pct, 1.0)

        # Tubo vacío de fondo
        ax.add_patch(FancyBboxPatch((i - ancho / 2, 0), ancho, 1.0,
                                    boxstyle="round,pad=0,rounding_size=0.08",
                                    facecolor="#F2F6FA", edgecolor="#ADBDCC",
                                    linewidth=1.6, zorder=1))

        # Líquido con degradado vertical
        if h > 0.01:
            colores = ["#1E8E44", "#7DDB8A"] if pct >= 1 else ["#1F6FC4", "#9BD4F5"]
            cmap = LinearSegmentedColormap.from_list("liquido", colores)
            grad = np.linspace(0, 1, 256).reshape(-1, 1)
            img = ax.imshow(grad, extent=(i - ancho / 2, i + ancho / 2, 0, h),
                            origin="lower", aspect="auto", cmap=cmap, zorder=2)
            recorte = FancyBboxPatch((i - ancho / 2, 0), ancho, h,
                                     boxstyle=f"round,pad=0,rounding_size={min(0.08, h / 2):.3f}",
                                     transform=ax.transData)
            img.set_clip_path(recorte)
            # Brillo en la superficie del líquido
            ax.plot([i - ancho / 2 + 0.05, i + ancho / 2 - 0.05], [h, h],
                    color="white", alpha=0.7, linewidth=2, zorder=3)

        color_num = "white" if h >= 0.12 else "#1F4E79"
        y_num = h / 2 if h >= 0.12 else h + 0.05
        ax.text(i, y_num, str(p), ha="center", va="center",
                fontsize=12, fontweight="bold", color=color_num, zorder=4)
        ax.text(i, 1.04, f"Meta {m}", ha="center", va="bottom",
                fontsize=9, fontweight="bold", color="#C00000")
        ax.text(i, 1.12, f"{round(pct * 100)}%", ha="center", va="bottom",
                fontsize=10, fontweight="bold",
                color="#1E8E44" if pct >= 1 else "#1F6FC4")

    ax.set_xticks(range(n))
    ax.set_xticklabels(nombres, rotation=45, ha="right", fontsize=9)
    ax.set_xlim(-0.7, n - 0.3)
    ax.set_ylim(0, 1.22)
    ax.set_yticks([])
    for sp in ax.spines.values():
        sp.set_visible(False)
    ax.set_title(titulo, fontsize=14, fontweight="bold", color="#1F3864", pad=16)

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
    # Crear gráfico (tubos de líquido)
    # -----------------------------
    nombres_cortos = [" ".join(s.split()[:2]).title() for s in sdv_list]
    IMG_FILE = IMG_FILE_TEMPLATE.format(local=local.replace(" ", "_"))
    dibujar_tubos_liquido(nombres_cortos, prod, meta, titulo_grafico, IMG_FILE)

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
