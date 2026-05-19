import subprocess
import os

# =====================================================
# CARPETA SCRIPTS
# =====================================================

BASE_DIR = r"C:\Users\KevinDanielShirayVer\OneDrive - INPROLEC S.A\Escritorio\Script\AutomatizacionOPs"

# =====================================================
# LISTA SCRIPTS
# =====================================================

scripts = [

    "00_Descargar-Correo.py",
    "01_Extraer-Facturas.py",
    "02_Odoo-SharePoint-Factura.py",
    "03_Generar-OP-Factura.py",
    "04_Crear-Correo-Factura.py"

]

# =====================================================
#                         EJECUTAR
# =====================================================

for script in scripts:

    ruta_script = os.path.join(
        BASE_DIR,
        script
    )

    print("\n" + "=" * 60)
    print(f"Ejecutando: {script}")
    print("=" * 60)

    resultado = subprocess.run(

        ["python", ruta_script],

        cwd=BASE_DIR

    )

    # =====================================================
    #                    VALIDAR ERROR
    # =====================================================

    if resultado.returncode != 0:

        print(f"\nERROR EN: {script}")

        break

print("\nFLUJO COMPLETO TERMINADO")
