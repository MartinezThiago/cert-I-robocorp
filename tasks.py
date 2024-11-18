"""Modulos de robocorp"""

from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.PDF import PDF

# En task se referencia a cada funcion que el robot deba ejecutar secuencialmente


@task
def robot_spare_bin_python():
    """Bajamos la velocidad del navegador"""
    browser.configure(
        slowmo=50,
    )
    open_the_intranet_website()
    log_in()
    download_excel_file()
    fill_form_with_excel_data()
    export_as_pdf()
    log_out()


def open_the_intranet_website():
    """Determino a que URL va a ir el robot"""
    browser.goto("https://robotsparebinindustries.com/")


def log_in():
    """Inicia sesion"""
    username = "maria"
    password = "thoushallnotpass"
    page = browser.page()
    page.fill("#username", username)
    page.fill("#password", password)
    page.click("button:text('Log in')")


def download_excel_file():
    """Descarga el excel del cual vamos a importar las filas con RPA.HTTP"""
    http = HTTP()
    # Overwrite pisa el archivo en el sistema con el mismo nombre cada vez que lo vuelve a descargar
    http.download(
        url="https://robotsparebinindustries.com/SalesData.xlsx", overwrite=True
    )


def fill_form_with_excel_data():
    """Recorremos el excel y llenamos la pagina con los campos"""
    excel = Files()
    excel.open_workbook("SalesData.xlsx")
    worksheet = excel.read_worksheet_as_table("data", header=True)
    page = browser.page()
    for row in worksheet:
        page.fill("#firstname", row["First Name"])
        page.fill("#lastname", row["Last Name"])
        page.select_option("#salestarget", str(row["Sales Target"]))
        page.fill("#salesresult", str(row["Sales"]))
        page.click("button:text('SUBMIT')")
    excel.close_workbook()


def export_as_pdf():
    """Descargamos la informacion en formato PDF"""
    page = browser.page()
    sales_results_html = page.locator("#sales-results").inner_html()
    pdf = PDF()
    pdf.html_to_pdf(sales_results_html, "output/sales_results.pdf")


def log_out():
    """Cerramos sesion"""
    page = browser.page()
    page.click("#logout")
