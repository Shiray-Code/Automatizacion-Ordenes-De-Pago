import os
import re
import json
import zipfile

from playwright.sync_api import sync_playwright

# =====================================================
# RUTAS
# =====================================================

JSON_FACTURAS = "facturas.json"

JSON_ODOO = "odoo_data.json"

PASTA_TEMP = r"C:\Users\KevinDanielShirayVer\OneDrive - INPROLEC S.A\Escritorio\Gestion de Compras\_temp"

USER_DATA_DIR = r"C:\PlaywrightChrome"

# =====================================================
# CREAR TEMP
# =====================================================

os.makedirs(
    PASTA_TEMP,
    exist_ok=True
)

# =====================================================
# SHAREPOINT
# =====================================================

def descargar_sharepoint(page, browser, carpeta_oc):

    sharepoint = None

    try:

        print("\nBuscando link SharePoint...")

        # =====================================================
        # OBTENER LINKS
        # =====================================================

        links = page.locator("a").evaluate_all(
            """
            elements => elements.map(el => el.href)
            """
        )

        link_sharepoint = None

        for link in links:

            if not link:
                continue

            if "sharepoint" in link.lower():

                link_sharepoint = link
                break

        if not link_sharepoint:

            print("No se encontró SharePoint")

            return

        print(f"SharePoint encontrado:\n{link_sharepoint}")

        # =====================================================
        # ABRIR SHAREPOINT
        # =====================================================

        sharepoint = browser.new_page()

        sharepoint.goto(link_sharepoint)

        sharepoint.wait_for_load_state(
            "domcontentloaded"
        )

        sharepoint.wait_for_timeout(10000)

        print("SharePoint abierto")

        # =====================================================
        # SELECCIONAR TODO
        # =====================================================

        sharepoint.mouse.click(
            500,
            500
        )

        sharepoint.wait_for_timeout(2000)

        sharepoint.keyboard.press(
            "Control+A"
        )

        sharepoint.wait_for_timeout(5000)

        print("CTRL+A realizado")

        # =====================================================
        # DOWNLOAD
        # =====================================================

        boton_descargar = sharepoint.locator(
            'button:has-text("Download"), button:has-text("Descargar")'
        ).first

        boton_descargar.wait_for(
            timeout=15000
        )

        print("Botón download encontrado")

        # =====================================================
        # DESCARGAR ZIP
        # =====================================================

        with sharepoint.expect_download(
            timeout=120000
        ) as download_info:

            boton_descargar.click(
                force=True
            )

        download = download_info.value

        nombre_zip = download.suggested_filename

        ruta_zip = os.path.join(
            carpeta_oc,
            nombre_zip
        )

        download.save_as(
            ruta_zip
        )

        print(f"ZIP descargado: {nombre_zip}")

        # =====================================================
        # EXTRAER ZIP
        # =====================================================

        with zipfile.ZipFile(
            ruta_zip,
            'r'
        ) as zip_ref:

            zip_ref.extractall(
                carpeta_oc
            )

        print("ZIP extraído")

        # =====================================================
        # BORRAR ZIP
        # =====================================================

        os.remove(
            ruta_zip
        )

        print("ZIP eliminado")

    except Exception as e:

        print(f"Error SharePoint: {e}")

    finally:

        try:

            if sharepoint:

                sharepoint.close()

        except:

            pass

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
# ODOO DATA
# =====================================================

odoo_data = {}

# =====================================================
# PLAYWRIGHT
# =====================================================

with sync_playwright() as p:

    browser = p.chromium.launch_persistent_context(

        user_data_dir=USER_DATA_DIR,

        channel="chrome",

        headless=False,

        accept_downloads=True,

        args=[

            "--disable-popup-blocking",
            "--disable-notifications",
            "--disable-features=Translate",
            "--disable-translate",
            "--no-first-run",
            "--no-default-browser-check"
        ]
    )

    page = browser.new_page()

    # =====================================================
    # PROVEEDORES
    # =====================================================

    for proveedor, facturas in facturas_por_proveedor.items():

        for factura in facturas:

            oc_actual = factura["oc"]

            print(f"\nOC: {oc_actual}")

            # =====================================================
            # TEMP OC
            # =====================================================

            carpeta_oc = os.path.join(
                PASTA_TEMP,
                oc_actual
            )

            os.makedirs(
                carpeta_oc,
                exist_ok=True
            )

            # =====================================================
            # ODOO
            # =====================================================

            page.goto(
                "https://erp.inprolec.cl/web?#view_type=list&model=purchase.order&action=378&menu_id=262"
            )

            page.wait_for_load_state(
                "domcontentloaded"
            )

            page.wait_for_timeout(1700)

            # =====================================================
            # LIMPIAR FILTROS
            # =====================================================

            try:

                page.locator(
                    ".fa.fa-sm"
                ).nth(0).click()

                page.wait_for_timeout(300)

                page.locator(
                    ".fa.fa-sm"
                ).nth(1).click()

                page.wait_for_timeout(700)

            except:

                pass

            # =====================================================
            # BUSCAR OC
            # =====================================================

            buscador = page.get_by_role(
                "textbox",
                name="Buscar…"
            )

            buscador.click()

            buscador.press(
                "Control+A"
            )

            buscador.press(
                "Backspace"
            )

            buscador.type(
                f"OC {oc_actual}",
                delay=120
            )

            page.wait_for_timeout(300)

            buscador.press("Enter")

            page.wait_for_timeout(300)

            fila_oc = page.locator(
                f'td:has-text("{oc_actual}")'
            ).first

            fila_oc.click()

            page.wait_for_timeout(300)

            print("OC abierta")

            # =====================================================
            # PROYECTO
            # =====================================================

            proyecto = ""

            try:

                proyecto = page.locator(
                    '[name="analytic_line_id"]'
                ).inner_text().strip()

                print(f"Proyecto: {proyecto}")

            except:

                pass

            # =====================================================
            # CENTRO COSTO
            # =====================================================

            centro_costo = ""

            try:

                textos = page.locator(
                    'td.o_data_cell.o_required_modifier'
                ).all_inner_texts()

                for texto in textos:

                    primera_linea = texto.split(
                        "\n"
                    )[0].strip()

                    if re.search(
                        r'^([A-Z]{2}\d{2}/[A-Z]{2}-\d{3}|[A-Z]\d{2}[A-Z]{3}\d+)',
                        primera_linea
                    ):

                        centro_costo = primera_linea

                        break

                print(f"Centro costo: {centro_costo}")

            except:

                pass

            # =====================================================
            # GUARDAR JSON
            # =====================================================

            odoo_data[oc_actual] = {

                "proyecto": proyecto,
                "centro_costo": centro_costo

            }

            # =====================================================
            # SHAREPOINT
            # =====================================================

            descargar_sharepoint(
                page,
                browser,
                carpeta_oc
            )

# =====================================================
# GUARDAR JSON
# =====================================================

with open(
    JSON_ODOO,
    "w",
    encoding="utf-8"
) as f:

    json.dump(
        odoo_data,
        f,
        indent=4,
        ensure_ascii=False
    )

print("\nODOO JSON GENERADO")
