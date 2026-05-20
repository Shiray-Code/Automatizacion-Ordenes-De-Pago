import os
import re
import win32com.client

from datetime import datetime

# =====================================================
# OUTLOOK
# =====================================================

outlook = win32com.client.Dispatch(
    "Outlook.Application"
).GetNamespace("MAPI")

# =====================================================
# BANDEJA ENTRADA
# =====================================================

bandeja = outlook.GetDefaultFolder(6)

mensajes = bandeja.Items

mensajes.Sort(
    "[ReceivedTime]",
    True
)

# =====================================================
# RUTA BASE
# =====================================================

PASTA_BASE = r"(Datos Personales)"

os.makedirs(
    PASTA_BASE,
    exist_ok=True
)

# =====================================================
# BUSCAR CORREO
# =====================================================

correo_encontrado = None

for mensaje in mensajes:

    try:

        remitente = str(
            mensaje.SenderName
        ).strip()

        asunto = str(
            mensaje.Subject
        ).strip()

        # =====================================================
        # FILTROS
        # =====================================================

        if (

            "(Datos Personales)".lower()
            in remitente.lower()

            and

            "NOTIFICACIÓN PARA REALIZAR OP DEL"
            in asunto.upper()

        ):

            correo_encontrado = mensaje

            print("\nCorreo encontrado:")
            print(f"Remitente: {remitente}")
            print(f"Asunto: {asunto}")

            break

    except:

        pass

# =====================================================
# VALIDAR
# =====================================================

if not correo_encontrado:

    raise Exception(
        "No se encontró correo válido."
    )

# =====================================================
# CREAR CARPETA SEMANA
# =====================================================

match_semana = re.search(
    r"SEMANA\s*(\d+)",
    correo_encontrado.Subject,
    re.IGNORECASE
)

if match_semana:

    numero_semana = match_semana.group(1)

else:

    numero_semana = datetime.now().strftime("%Y%m%d")

nombre_carpeta = f"Semana_{numero_semana}"

PASTA_SEMANA = os.path.join(
    PASTA_BASE,
    nombre_carpeta
)

os.makedirs(
    PASTA_SEMANA,
    exist_ok=True
)

print(f"\nCarpeta creada:")
print(PASTA_SEMANA)

# =====================================================
# DESCARGAR ADJUNTOS
# =====================================================

descargados = 0

for adjunto in correo_encontrado.Attachments:

    nombre_archivo = adjunto.FileName

    # =====================================================
    # SOLO PDF
    # =====================================================

    if not nombre_archivo.lower().endswith(
        ".pdf"
    ):

        continue

    ruta_guardado = os.path.join(
        PASTA_SEMANA,
        nombre_archivo
    )

    adjunto.SaveAsFile(
        ruta_guardado
    )

    descargados += 1

    print(f"Descargado: {nombre_archivo}")

# =====================================================
# RESULTADO
# =====================================================

print("\n" + "=" * 60)
print(f"PDFs descargados: {descargados}")
print("PROCESO TERMINADO")
print("=" * 60)
