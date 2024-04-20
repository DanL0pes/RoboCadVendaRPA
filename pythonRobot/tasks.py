from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.Excel.Files import Files

@task
def robot_tasks_bin_python():
    """Inserindo os dados de venda da semana e exportar em PDF"""
    abrir_intranet_site()
    log_in()
    download_excel_file()
    ler_vendas_excel()

def abrir_intranet_site():
    """Navega para o URL dado"""
    browser.goto("https://robotsparebinindustries.com/")

def log_in():
    """Preencher usuario e senha, depois click em logar"""
    page = browser.page()
    page.fill('#username', 'maria')
    page.fill('#password', 'thoushallnotpass')
    page.click("button:text('Log in')")

def download_excel_file():
    """Download excel file atraves de uma URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/SalesData.xlsx", overwrite=True)

def ler_vendas_excel():
    """Ler os dados do excel e guardar em uma variavel"""
    excel = Files()
    excel.open_workbook("SalesData.xlsx")
    worksheet = excel.read_worksheet_as_table("data", header=True)
    excel.close_workbook()
    for row in worksheet:
        cadastrar_vendas(row)
    #FIM FOR

def cadastrar_vendas(vendas_data):
    """Preencher os dados de venda e click no botao submit"""
    page = browser.page()
    page.fill('#firstname', vendas_data["First Name"])
    page.fill('#lastname', vendas_data["Last Name"])
    page.select_option('#salestarget', str(vendas_data["Sales Target"]))
    page.fill('#salesresult', str(vendas_data["Sales"]))
    page.click("text=Submit")