import os
import re
import json
import pdfplumber

from collections import defaultdict

# =====================================================
# RUTAS
# =====================================================

PASTA_FACTURAS_BASE = r"C:\Users\KevinDanielShirayVer\OneDrive - INPROLEC S.A\Escritorio\Gestion de Compras\Facturas Por semana"

SALIDA_JSON = "facturas.json"

# =====================================================
# DETECTAR CARPETA MÁS RECIENTE
# =====================================================

carpetas = []

for carpeta in os.listdir(
    PASTA_FACTURAS_BASE
):

    ruta_carpeta = os.path.join(
        PASTA_FACTURAS_BASE,
        carpeta
    )

    if os.path.isdir(
        ruta_carpeta
    ):

        carpetas.append(
            ruta_carpeta
        )

if not carpetas:

    raise Exception(
        "No se encontraron carpetas semanales."
    )

PASTA_FACTURAS = max(
    carpetas,
    key=os.path.getmtime
)

print("=" * 60)
print("CARPETA SEMANAL:")
print(PASTA_FACTURAS)
print("=" * 60)

# =====================================================
# FACTURAS
# =====================================================

facturas_por_proveedor = defaultdict(list)

for archivo in os.listdir(PASTA_FACTURAS):

    if not archivo.lower().endswith(".pdf"):
        continue

    ruta_pdf = os.path.join(
        PASTA_FACTURAS,
        archivo
    )

    texto_completo = ""

    print(f"\nProcesando: {archivo}")

    with pdfplumber.open(ruta_pdf) as pdf:

        for pagina in pdf.pages:

            texto = pagina.extract_text()

            if texto:

                texto_completo += texto + "\n"

    # =====================================================
    # OC
    # =====================================================

    match_oc = re.search(
        r"\b46\d{8}\b",
        texto_completo
    )

    oc = match_oc.group(0) if match_oc else ""

    # =====================================================
    # HES
    # =====================================================

    match_hes = re.search(
        r"HES[:\s]*(\d+)",
        texto_completo,
        re.IGNORECASE
    )

    hes = match_hes.group(1) if match_hes else ""

    # =====================================================
    # FACTURA
    # =====================================================

    match_factura = re.search(
        r"(?:N°|Folio)[^\d]*(\d+)",
        texto_completo,
        re.IGNORECASE
    )

    factura_numero = (
        match_factura.group(1)
        if match_factura
        else ""
    )

    # =====================================================
    # MONTO
    # =====================================================

    montos = re.findall(
        r"\$[\s]*([\d\.]+)",
        texto_completo
    )

    monto = 0

    if montos:

        monto = max([

            int(
                m.replace(".", "")
            )

            for m in montos
        ])

    # =====================================================
    # PROVEEDOR
    # =====================================================

    proveedor = "SIN_PROVEEDOR"

    lineas = texto_completo.split("\n")

    for linea in lineas:

        linea_upper = linea.upper().strip()

        # =====================================================
        # BUSCAR EMPRESA
        # =====================================================

        if any([

            "SPA" in linea_upper,
            "S.A" in linea_upper,
            "LTDA" in linea_upper,
            "LIMITADA" in linea_upper,
            "EIRL" in linea_upper

        ]):

            # =====================================================
            # IGNORAR BASURA
            # =====================================================

            if any([

                "SEÑOR" in linea_upper,
                "FACTURA" in linea_upper,
                "CLIENTE" in linea_upper,
                "CEDIBLE" in linea_upper,
                "INPROLEC" in linea_upper

            ]):

                continue

            proveedor = linea.strip()

            # =====================================================
            # ELIMINAR RUT
            # =====================================================

            proveedor = re.sub(
                r'R\.?U\.?T\.?.*',
                '',
                proveedor,
                flags=re.IGNORECASE
            )

            # =====================================================
            # LIMPIAR CARACTERES
            # =====================================================

            proveedor = re.sub(
                r'[\\/:*?"<>|]',
                "-",
                proveedor
            )

            proveedor = proveedor.strip()

            break

    # =====================================================
    # DESCRIPCIÓN
    # =====================================================

    descripcion = ""

    match_desc = re.search(
        r"1,00\s+(.*?)\s+\$",
        texto_completo,
        re.DOTALL
    )

    if match_desc:

        descripcion = (
            match_desc.group(1)
            .strip()
        )

    # =====================================================
    # DEBUG
    # =====================================================

    print(f"Proveedor detectado: {proveedor}")

    # =====================================================
    # VALIDAR
    # =====================================================

    if proveedor == "SIN_PROVEEDOR":

        print("Factura ignorada")
        continue

    print(f"OC: {oc}")
    print(f"HES: {hes}")
    print(f"Factura: {factura_numero}")
    print(f"Monto: {monto}")

    # =====================================================
    # GUARDAR
    # =====================================================

    facturas_por_proveedor[
        proveedor
    ].append({

        "archivo": archivo,
        "oc": oc,
        "hes": hes,
        "factura": factura_numero,
        "monto": monto,
        "descripcion": descripcion,
        "ruta_pdf": ruta_pdf

    })

# =====================================================
# GUARDAR JSON
# =====================================================

with open(
    SALIDA_JSON,
    "w",
    encoding="utf-8"
) as f:

    json.dump(
        facturas_por_proveedor,
        f,
        indent=4,
        ensure_ascii=False
    )

print("\nJSON generado")
