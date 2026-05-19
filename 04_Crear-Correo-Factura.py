import os
import json
import win32com.client

# =====================================================
# RUTAS
# =====================================================

JSON_FACTURAS = "facturas.json"

PASTA_OP = r"C:\Users\KevinDanielShirayVer\OneDrive - INPROLEC S.A\Escritorio\Gestion de Compras\Proveedores-OP"

# =====================================================
# LEER JSON
# =====================================================

with open(
    JSON_FACTURAS,
    "r",
    encoding="utf-8"
) as f:

    facturas_por_proveedor = json.load(f)

# =====================================================
# OUTLOOK
# =====================================================

outlook = win32com.client.Dispatch(
    "Outlook.Application"
)

# =====================================================
# RECORRER PROVEEDORES
# =====================================================

for proveedor, facturas in facturas_por_proveedor.items():

    print("\n" + "=" * 60)
    print(f"PROVEEDOR: {proveedor}")

    # =====================================================
    # BUSCAR ÚLTIMA OP
    # =====================================================

    carpeta_proveedor = os.path.join(
        PASTA_OP,
        proveedor
    )

    if not os.path.exists(
        carpeta_proveedor
    ):

        print(
            "No existe carpeta proveedor"
        )

        continue

    carpetas_op = [

        os.path.join(
            carpeta_proveedor,
            carpeta
        )

        for carpeta in os.listdir(
            carpeta_proveedor
        )

        if os.path.isdir(
            os.path.join(
                carpeta_proveedor,
                carpeta
            )
        )

        and carpeta.startswith(
            "OP-KS-"
        )
    ]

    if not carpetas_op:

        print(
            "No existen OPs"
        )

        continue

    carpeta_op = max(
        carpetas_op,
        key=os.path.getmtime
    )

    nombre_op = os.path.basename(
        carpeta_op
    )

    # =====================================================
    # DESCRIPCIONES
    # =====================================================

    descripciones = []

    for factura in facturas:

        descripcion = factura[
            "descripcion"
        ]

        if descripcion not in descripciones:

            descripciones.append(
                descripcion
            )

    descripcion_correo = ", ".join(
        descripciones
    )

    # =====================================================
    # CREAR CORREO
    # =====================================================

    try:

        mail = outlook.CreateItem(0)

        # =====================================================
        # DESTINATARIOS
        # =====================================================

        mail.To = (
            "Guido.Ferrufino@inprolec.cl"
        )

        mail.CC = ""

        # =====================================================
        # ASUNTO
        # =====================================================

        mail.Subject = (
            f"{nombre_op} - {proveedor}"
        )

        # =====================================================
        # MOSTRAR
        # =====================================================

        mail.Display()

        firma = mail.HTMLBody

        # =====================================================
        # CUERPO
        # =====================================================

        mail.HTMLBody = f"""
        <div style="font-family:Calibri; font-size:16px;">

        <p>Estimado Guido,</p>

        <p>
        Favor de revisar, aprobar y firmar las siguiente OP de {proveedor}
        referente a factura por {descripcion_correo}.
        </p>

        <p>
        Muchas gracias de antemano.
        </p>

        <br>

        </div>

        {firma}
        """

        print(
            f"Correo generado: {nombre_op}"
        )

    except Exception as e:

        print(
            f"Error correo: {e}"
        )

print("\nPROCESO TERMINADO")
