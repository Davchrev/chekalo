import win32com.client
import pandas as pd
from bs4 import BeautifulSoup
import re
from datetime import datetime, date, timedelta
import matplotlib.pyplot as plt
import numpy as np
import os
import locale

# Configurar el idioma a español para el nombre del día
try:
    locale.setlocale(locale.LC_TIME, "es_ES.UTF-8")
except:
    locale.setlocale(locale.LC_TIME, "es_ES") # Alternativa para Windows

# -----------------------------
# Parámetro de fecha (constante editable)
# Opciones: "hoy", "ayer" o una fecha exacta "YYYY-MM-DD"
# -----------------------------
FECHA = "ayer"

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
# Configuración de Archivos
# -----------------------------
EXCEL_FILE = os.path.join(os.getcwd(), "Reporte_Casos_Detallados.xlsx")
DOTACION_FILE = os.path.join(os.getcwd(), "Dotacion1.xlsx")
IMG_FILE = os.path.join(os.getcwd(), "avance_checkalo.png")

# -----------------------------
# Conexión a Outlook
# -----------------------------
outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
inbox = namespace.GetDefaultFolder(6)  # Bandeja de entrada
robot_folder = inbox.Folders["SEGUIMIENTO"].Folders["ROBOT-CHECKALO"]

campos_interes = ["Tipo de Integrante", "Nro Doc", "Razon Social", "Apellido Paterno", "Apellido Materno", "Nombre1", "Nombre2"]
data = []

# -----------------------------
# Procesamiento de Correos
# -----------------------------
for mail in robot_folder.Items:
    try:
        if mail.Class != 43: continue

        sender_email = mail.SenderEmailAddress
        sender_name = mail.SenderName
        subject = mail.Subject
        fecha_outlook = mail.SentOn
        fecha = datetime(fecha_outlook.year, fecha_outlook.month, fecha_outlook.day,
                         fecha_outlook.hour, fecha_outlook.minute, fecha_outlook.second) if fecha_outlook else None

        soup = BeautifulSoup(mail.HTMLBody, "html.parser")
        cuerpo = soup.get_text(separator="\n").replace("\xa0", " ")

        # EXTRAER USUARIO VENDOR
        match_v = re.search(r"USUARIO\s*VENDOR\s*:\s*(\d+)", cuerpo, re.I)
        usuario_vendor = match_v.group(1) if match_v else "NO_DETECTADO"

        lineas = [l.strip() for l in cuerpo.splitlines() if l.strip()]

        bloque = {c: "" for c in campos_interes}
        num_actual = None

        for i, l in enumerate(lineas):
            # 1. Detectar inicio de bloque (ej: "01", "02")
            m_bloque = re.match(r"^(\d{2})(?:\s|$)", l)

            # Si detectamos cambio de bloque, guardamos el anterior
            if m_bloque:
                num_det = m_bloque.group(1)
                if num_det != num_actual:
                    # GUARDAR BLOQUE ANTERIOR (Solo si tiene Nro Doc)
                    if num_actual is not None and bloque.get("Nro Doc"):
                        res = bloque.copy()
                        res.update({"USUARIO VENDOR": usuario_vendor, "Correo Remitente": sender_email,
                                    "Nombre Remitente": sender_name, "Fecha Envío": fecha, "Asunto": subject})
                        data.append(res)

                    # Reset para nuevo bloque
                    bloque = {c: "" for c in campos_interes}
                    num_actual = num_det

            # 2. Detectar Campo: Valor (Flexible con o sin número adelante)
            m_campo_val = re.match(r"^(?:\d{2}\s+)?(.+?):\s*(.*)$", l)

            if m_campo_val:
                campo_leido = m_campo_val.group(1).strip().upper()
                valor_leido = m_campo_val.group(2).strip()

                # Si el valor está vacío, buscar en la siguiente línea
                if not valor_leido and i + 1 < len(lineas):
                    siguiente_linea = lineas[i + 1].strip()
                    # Verificar que la siguiente línea no sea inicio de nuevo campo o bloque
                    if not re.match(r"^(?:\d{2}\s+)?(.+?):\s*(.*)$", siguiente_linea) and not re.match(r"^(\d{2})(?:\s|$)", siguiente_linea):
                        valor_leido = siguiente_linea

                # Comparación insensible a mayúsculas para mapear al campo correcto
                for c_real in campos_interes:
                    if c_real.upper() == campo_leido:
                        bloque[c_real] = valor_leido
                        break

            # 3. Parche para DNI que aparece solo en una línea (8 u 11 dígitos)
            # SOLO si estamos dentro de un bloque y no es un campo con formato "Campo:"
            elif num_actual is not None and not m_campo_val:
                m_dni_suelto = re.match(r"^(\d{8}|\d{11})$", l)
                if m_dni_suelto and not bloque["Nro Doc"]:
                    bloque["Nro Doc"] = m_dni_suelto.group(1)

        # Guardar último bloque del correo (Si cumple la regla del DNI)
        if num_actual is not None and bloque.get("Nro Doc"):
            res = bloque.copy()
            res.update({"USUARIO VENDOR": usuario_vendor, "Correo Remitente": sender_email,
                        "Nombre Remitente": sender_name, "Fecha Envío": fecha, "Asunto": subject})
            data.append(res)

    except Exception as e:
        print(f"❌ Error en correo {subject}: {e}")

# -----------------------------
# Consolidación de Datos
# -----------------------------
if not data:
    print("⚠️ No se encontraron registros válidos con Nro Doc."); exit()

df = pd.DataFrame(data)

# Cruce con Dotación para obtener Zonas/Locales
if os.path.exists(DOTACION_FILE):
    dot_df = pd.read_excel(DOTACION_FILE, dtype={"USUARIO VENDOR": str})
    df["USUARIO VENDOR"] = df["USUARIO VENDOR"].astype(str)
    df = df.merge(dot_df[["USUARIO VENDOR", "Territorio", "LOCAL", "SEDE", "SupervisorVenta", "JZ"]], on="USUARIO VENDOR", how="left")

# -----------------------------
# Hardcodear dotación para usuarios que no cruzan
# -----------------------------
hardcode_dotacion = {
    "3377951": {"Territorio": "PROYECTO PB", "LOCAL": "PB", "SEDE": "PUCALLPA", "SupervisorVenta": "MEDINA RIOS KARLA CECILIA", "JZ": "JOSE LUIS COLOMA"},
    "3377961": {"Territorio": "PROYECTO PB2", "LOCAL": "PB", "SEDE": "PB2", "SupervisorVenta": "MONTENEGRO REQUEJO ALBERTO ALEJANDRO", "JZ": "JOSE LUIS COLOMA"},
    "3377953": {"Territorio": "PROYECTO PB2", "LOCAL": "PB", "SEDE": "PB2", "SupervisorVenta": "MONTENEGRO REQUEJO ALBERTO ALEJANDRO", "JZ": "JOSE LUIS COLOMA"},
    "3374051": {"Territorio": "PROYECTO PB2", "LOCAL": "PB", "SEDE": "PB2", "SupervisorVenta": "RIVERA CAVERO JOSE ANTONIO", "JZ": "JOSE LUIS COLOMA"},
    "3377956": {"Territorio": "PROYECTO PB2", "LOCAL": "PB", "SEDE": "PB2", "SupervisorVenta": "RIVERA CAVERO JOSE ANTONIO", "JZ": "JOSE LUIS COLOMA"},
    "3377944": {"Territorio": "PROYECTO PB2", "LOCAL": "PB", "SEDE": "PB2", "SupervisorVenta": "MALPARTIDA MIRANDA ALFONSO", "JZ": "JOSE LUIS COLOMA"},
    "33877740": {"Territorio": "TERRITORIO LIMA", "LOCAL": "LIMA 2", "SEDE": "LIMA CERCADO", "SupervisorVenta": "ROSAS MENDOZA AURORA MARIA DEL ROSARIO", "JZ": "BETETA CHOTA NELSON SANTIAGO"},
    "337": {"Territorio": "TERRITORIO LIMA", "LOCAL": "LIMA 2", "SEDE": "LIMA CERCADO", "SupervisorVenta": "BETETA CHOTA NELSON SANTIAGO", "JZ": "BETETA CHOTA NELSON SANTIAGO"},
    "33": {"Territorio": "TERRITORIO LIMA", "LOCAL": "LIMA 2", "SEDE": "LIMA CERCADO", "SupervisorVenta": "BETETA CHOTA NELSON SANTIAGO", "JZ": "BETETA CHOTA NELSON SANTIAGO"},
    "3377880": {"Territorio": "PROYECTO PB2", "LOCAL": "PB", "SEDE": "PB2", "SupervisorVenta": "MALPARTIDA MIRANDA ALFONSO", "JZ": "JOSE LUIS COLOMA"},
    "3373473": {"Territorio": "PROYECTO PB2", "LOCAL": "PB", "SEDE": "PB2", "SupervisorVenta": "MALPARTIDA MIRANDA ALFONSO", "JZ": "JOSE LUIS COLOMA"},
    "3377946": {"Territorio": "PROYECTO PB2", "LOCAL": "PB", "SEDE": "PB2", "SupervisorVenta": "MALPARTIDA MIRANDA ALFONSO", "JZ": "JOSE LUIS COLOMA"},
    "3377949": {"Territorio": "PROYECTO PB2", "LOCAL": "PB", "SEDE": "PB2", "SupervisorVenta": "MALPARTIDA MIRANDA ALFONSO", "JZ": "JOSE LUIS COLOMA"},
    "3377955": {"Territorio": "PROYECTO PB2", "LOCAL": "PB", "SEDE": "PB2", "SupervisorVenta": "RIVERA CAVERO JOSE ANTONIO", "JZ": "JOSE LUIS COLOMA"},
    "3377960": {"Territorio": "PROYECTO PB2", "LOCAL": "PB", "SEDE": "PB2", "SupervisorVenta": "MONTENEGRO REQUEJO ALBERTO ALEJANDRO", "JZ": "JOSE LUIS COLOMA"},
    "3377947": {"Territorio": "PROYECTO PB2", "LOCAL": "PB", "SEDE": "PB2", "SupervisorVenta": "MALPARTIDA MIRANDA ALFONSO", "JZ": "JOSE LUIS COLOMA"},
    "3374070": {"Territorio": "TERRITORIO NORTE ORIENTE", "LOCAL": "ZONA TRUJILLO", "SEDE": "TRUJILLO 2", "SupervisorVenta": "VILLEGAS ARIAS ROBERTO", "JZ": "VACANTE TRUJILLO SRDV"},
    "3377749": {"Territorio": "TERRITORIO CENTRO SUR", "LOCAL": "ZONA CUSCO", "SEDE": "CUSCO 1", "SupervisorVenta": "SALAZAR GAMARRA HAROLD YAMIL", "JZ": "AMERICO SALAZAR PEZET"},
    "3377964": {"Territorio": "PROYECTO PB2", "LOCAL": "PB", "SEDE": "PB2", "SupervisorVenta": "HUAMAN ROMAN MAURA", "JZ": "JOSE LUIS COLOMA"},
    "3377963": {"Territorio": "PROYECTO PB2", "LOCAL": "PB", "SEDE": "PB2", "SupervisorVenta": "HUAMAN ROMAN MAURA", "JZ": "JOSE LUIS COLOMA"},
    "3373580": {"Territorio": "TERRITORIO NORTE ORIENTE", "LOCAL": "ZONA TRUJILLO", "SEDE": "TRUJILLO 2", "SupervisorVenta": "VARGAS PEREZ TEODORO LEONIDAS", "JZ": "VACANTE TRUJILLO SRDV"},
    "3377976": {"Territorio": "PROYECTO PB2", "LOCAL": "PB", "SEDE": "PB2", "SupervisorVenta": "HUAMAN ROMAN MAURA", "JZ": "JOSE LUIS COLOMA"},
    "3377861": {"Territorio": "TERRITORIO NORTE ORIENTE", "LOCAL": "ZONA TRUJILLO", "SEDE": "TRUJILLO 2", "SupervisorVenta": "VALERA VILLENA DIEGO ALONSO", "JZ": "VACANTE TRUJILLO SRDV"},
    "3377932": {"Territorio": "TERRITORIO LIMA", "LOCAL": "LIMA 1", "SEDE": "LIMA ", "SupervisorVenta": "DORADOR QUINTANA FREDDY RONALD", "JZ": "BARAHONA ZAVALA JOSE GABRIEL"},
    "3377708 ": {"Territorio": "TERRITORIO CENTRO SUR", "LOCAL": "ZONA CUSCO", "SEDE": "CUSCO 1", "SupervisorVenta": "SALAZAR GAMARRA HAROLD YAMIL", "JZ": "AMERICO SALAZAR PEZET"},
    "3377945": {"Territorio": "PROYECTO PB2", "LOCAL": "PB", "SEDE": "PB2", "SupervisorVenta": "MALPARTIDA MIRANDA ALFONSO", "JZ": "JOSE LUIS COLOMA"},
    "3377975": {"Territorio": "PROYECTO PB2", "LOCAL": "PB", "SEDE": "PB2", "SupervisorVenta": "HUAMAN ROMAN MAURA", "JZ": "JOSE LUIS COLOMA"},
    "3377599": {"Territorio": "PROYECTO PB2", "LOCAL": "PB", "SEDE": "PB2", "SupervisorVenta": "RIVERA CAVERO JOSE ANTONIO", "JZ": "JOSE LUIS COLOMA"},
    "3377772": {"Territorio": "PROYECTO PB2", "LOCAL": "PB", "SEDE": "PB2", "SupervisorVenta": "MALPARTIDA MIRANDA ALFONSO", "JZ": "JOSE LUIS COLOMA"}
}

for usuario, info in hardcode_dotacion.items():
    mask = df["USUARIO VENDOR"] == usuario
    for col, val in info.items():
        df.loc[mask, col] = val

# Exportar Excel Histórico
df.to_excel(EXCEL_FILE, index=False)

# -----------------------------
# Lógica para el Gráfico (KPI)
# -----------------------------
df["Fecha Envío"] = pd.to_datetime(df["Fecha Envío"])

# Filtrar solo lo enviado en la fecha del reporte
df_dia = df[df["Fecha Envío"].dt.date == fecha_reporte].copy()
df_dia_unique = df_dia.drop_duplicates(subset=["Nro Doc"])

# Configuración de Zonas y Metas
zonas = ["LIMA 1","LIMA 2","PB","ZONA AREQUIPA","ZONA CENTRO","ZONA CHICLAYO","ZONA CUSCO","ZONA ORIENTE","ZONA TACNA","ZONA TRUJILLO"]
sdv_map = {"LIMA 1": 4, "LIMA 2": 5, "PB": 9, "ZONA AREQUIPA": 6, "ZONA CENTRO": 2, "ZONA CHICLAYO": 1, "ZONA CUSCO": 2, "ZONA ORIENTE": 3, "ZONA TACNA": 1, "ZONA TRUJILLO": 5}

prod_dia = df_dia_unique["LOCAL"].value_counts().reindex(zonas, fill_value=0)
meta_dia = pd.Series({z: 20 * sdv_map.get(z, 1) for z in zonas})
meta_dia["ZONA ORIENTE"] = 46
cumplimiento = (prod_dia / meta_dia).fillna(0)

# Generar Gráfico (barras ejecutivas horizontales)
dibujar_barras_ejecutivo(zonas, prod_dia.tolist(), meta_dia.tolist(),
                         f"Producción Checkalo - {fecha_texto}", IMG_FILE)

# -----------------------------
# Envío de Reporte por Outlook
# -----------------------------
fecha_hora_texto = datetime.now().strftime("%A %d/%m/%y %H:%M").lower()
hora_asunto = datetime.now().strftime("%H:%M")

# Etiqueta en el asunto y texto de corte según la fecha del reporte
if es_hoy:
    etiqueta = ""
    corte_texto = fecha_hora_texto
elif fecha_reporte == date.today() - timedelta(days=1):
    etiqueta = " (AYER)"
    corte_texto = fecha_texto
else:
    etiqueta = f" ({fecha_texto})"
    corte_texto = fecha_texto

mail_out = outlook.CreateItem(0)
mail_out.To = "marceo@buro.com.pe; dherrera@buro.com.pe; gperez@buro.com.pe; jbarahona.stbk@buro.com.pe; sbeteta@buro.com.pe; jcolomad@buro.com.pe; svillenas@buro.com.pe; jvargas.stbk@buro.com.pe; gdavilas@buro.com.pe; asalazar.stbk@buro.com.pe; lliendo.stbk@buro.com.pe; lrodriguezr@buro.com.pe"
mail_out.Subject = f"Avance diario consultas Checkalo{etiqueta} - {hora_asunto}"

mail_out.HTMLBody = f"""
<html>
<body style="font-family: Calibri, sans-serif;">
    <p>Buen día,</p>
    <p>Se adjunta el reporte de consultas Checkalo procesadas con corte al día <b>{corte_texto}</b>.</p>
    <img src="cid:grafico_avance">
    <br>
    <p>Saludos,<br><b>Robot Checkalo</b></p>
</body>
</html>
"""

# Adjuntar Imagen embebida
att_img = mail_out.Attachments.Add(IMG_FILE)
att_img.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "grafico_avance")

# Adjuntar Excel
mail_out.Attachments.Add(EXCEL_FILE)

mail_out.Display() # Usa .Send() si quieres que se envíe automáticamente
print(f"✔ Proceso finalizado. {len(df_dia_unique)} registros únicos detectados el {fecha_texto}.")
