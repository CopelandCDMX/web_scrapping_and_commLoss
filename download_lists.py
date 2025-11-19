from selenium import webdriver 
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service # i added this line
import os
import time
import win32com.client as win32
# from selenium.webdriver.chrome.service import Service
# from selenium.common.exceptions import TimeoutException , NoSuchElementException


def inicio_pasword(file_path: str, driver: webdriver.Chrome, wait: WebDriverWait, lista_correo_errores: list):
    """This is a special function that introduces credentials in Connect +. 
    -----------------------------------------
    Parameters:
    ----------
    file_path : str
        Path where the credentials file is located
    driver : webdriver.Chrome
        Chrome webdriver instance   
    wait : WebDriverWait
        WebDriverWait instance for waiting for elements
    lista_correo_errores : list
        List of email addresses for error notifications
    -----------------------
    Returns:
    ------- 
    None
    """
    sesion_frame = wait.until(EC.presence_of_element_located((By.ID, "navId")))
    driver.switch_to.frame(sesion_frame)

    username_input = wait.until(EC.presence_of_element_located((By.NAME, "_UserName")))
    password_input = wait.until(EC.presence_of_element_located((By.NAME, "_Password")))
    login_button = wait.until(EC.element_to_be_clickable((By.NAME, "loginButton")))
    #password_input = driver.find_element(By.NAME, "_Password") i changed this
    #login_button = driver.find_element(By.NAME, "loginButton") i changed this
    print('found input fields\n')
    #k_path = path_programa + '\\txt\\Credenciales_Walmart.txt'
    k_path = file_path
    print(k_path)
    with open(k_path, 'r', encoding='utf-8') as file:
        lineas = file.readlines()

    print('Attempting login...\n')
    #cada línea a una variable
    u = lineas[0].strip()
    p = lineas[1].strip()

    username_input.send_keys(u)
    password_input.send_keys(p)
    login_button.click()

    #login_button = driver.find_element(By.ID, "ssoButton")
    #login_button.click()
    #print("Inicio de sesión exitoso!\n")
                #login_button.click()
    #switch back out of the frame after the clic, as the logging process usually loads a new page OUSIDE the frame.
    driver.switch_to.default_content()

    #print("Inicio de sesión exitoso!\n")
    try:
        element = driver.find_element(By.CLASS_NAME, "invalidFields")
        if "Password will expire in" in element.text:
            login_button.click()
            send_mail_app_escritorio(2, lista_correo_errores, lista_correo_errores, 'Contraseña de cuenta de descarga de alarmas por expirar', f'La contraseña de la cuenta {username_input_walmart} está por expirar', lista_correo_errores)
        else:
            print("El elemento existe, pero el texto no coincide.")
            time.sleep(2)
    except:
        pass


#Funcion para envio de alarmas 
def send_mail_app_escritorio(importancia: int,destinatario: list,copia:list ,subjet: list ,cuerpo, lista_correo_errores:list):
    destinatario = ';'.join(destinatario)
    copia = ';'.join(copia) 

    olApp = win32.Dispatch('Outlook.Application')
    mailItem = olApp.CreateItem(0)
    mailItem.Importance = importancia
    mailItem.to = destinatario
    mailItem.CC = copia
    mailItem.Subject = subjet
    # Creamos el objeto mensaje
    mailItem.HTMLbody = cuerpo

    try:
        mailItem.send
    except Exception as e:
        error_message = str(e)

        send_mail_app_escritorio(int(2), lista_correo_errores, lista_correo_errores, 'Fallo en envío de alarmas de Copeland', error_message)
        print('Error al enviar correo:',repr(e))
        print(f'Error: {e}\n')
        time.sleep(5)

def extraer_alarmas_connect(formatos_tienda: list, lista_correo_errores: list):
    for i in range(3):
    # try:
        path_programa = os.getcwd()   ## path where the code is living
        driver_path = os.path.join(path_programa, "chromedriver-win64", "chromedriver.exe")
        credentials_path = os.path.join(path_programa, "credentials.txt")
        # Configurar opciones de Chrome para abrir en inglés
        options = Options()
        options.add_argument("--lang=en")  # forza el idioma de chrome a inglés

        #service = webdriver.ChromeService(driver_path)  ### I changed this line
        service = Service(executable_path=driver_path) ### updated line
        driver = webdriver.Chrome(service=service, options=options)

        #tiempo de espera
        wait = WebDriverWait(driver, 20)

        try:
            # Abrir la página web
            #url_walmart = "https://walmartca.my-connectplus.com/walmartca/"   ###### una vez aqui, probably i need to put this as a parameter
            url_cpla = "https://cpla.my-connectplus.com/cpla/"     ###### una vez aqui, enfoque en este
            driver.get(url_cpla)
            #driver.get(url_walmart)
            wait.until(EC.presence_of_element_located((By.TAG_NAME, "html")))
            print("Página cargada\n")


           # def inicio_walmart():
            #    sesion_frame = wait.until(EC.presence_of_element_located((By.ID, "navId")))
            #    driver.switch_to.frame(sesion_frame)

            #    username_input_walmart = wait.until(EC.presence_of_element_located((By.NAME, "_UserName")))
            #    password_input_walmart = driver.find_element(By.NAME, "_Password")
            #    login_button = driver.find_element(By.ID, "ssoButton") #//*[@id="ssoButton"]

            #    login_button.click()

             #   print("Inicio de sesión exitoso!\n")

            #iniciar sesión
            try:
                print('Entering login function\n')
                inicio_pasword(file_path= credentials_path, driver=driver, wait=wait, lista_correo_errores=lista_correo_errores) 
                print('get out password function\n')
                print("Password function completed")

            except Exception as e:
                error_message = str(e)
                send_mail_app_escritorio(int(2), lista_correo_errores, lista_correo_errores, 'No se encontró el formulario de inicio de sesión.', error_message, lista_correo_errores)
                print("Fallo el inicio de sesión.\n")
                driver.quit()
                exit()

            #ir a la sección de alarmas   ###############
            try:
                driver.switch_to.default_content()
                sesion_frame = wait.until(EC.presence_of_element_located((By.ID, "mainId")))
                driver.switch_to.frame(sesion_frame)

                # Alarmas en instancia de Walmart
                wait.until(EC.element_to_be_clickable((By.XPATH, '//img[@title="Alarm"]'))).click() #<img class="header-img" src="/cpla/Images/alarm.png" title="Alarm">
                print('Clic en boton alarmas')

                # espera que la pantalla gris desaparezca
                wait.until(EC.invisibility_of_element((By.CLASS_NAME, "ext-el-mask"))) 

                # Expandir menú de alarmas activas
                wait.until(EC.element_to_be_clickable((By.NAME, "state"))).click()  
                print('Expandir menu alarmas All, Activas,  RTN')

                # espera que la pantalla gris desaparezca
                wait.until(EC.invisibility_of_element((By.CLASS_NAME, "ext-el-mask")))
                wait.until(EC.element_to_be_clickable((By.XPATH, '//div[contains(@class, "x-combo-list-item") and text()="Active"]'))).click() # Seleccionar opción de alarmas activas
                print('Clic en Alarmas Activas')

                # espera que la pantalla gris desaparezca
                wait.until(EC.invisibility_of_element((By.CLASS_NAME, "ext-el-mask")))

                # Busca la columna de directorio
                # columna = driver.find_element(By.ID, "sourceDir")
                columna = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="sourceDir"]')))
                print('Se encontro la columna de direccion')
                time.sleep(5)
                # Desplázate horizontalmente hasta esa columna usando JavaScript
                driver.execute_script("arguments[0].scrollIntoView(true);", columna)
                print('Scroll a la columna')

                # Ahora puedes hacer clic o interactuar con ella
                # wait.until(EC.element_to_be_clickable((columna))).click()
                time.sleep(5)
                # print('Clic en barra de formato')

                #busqueda por formato
                input_element = driver.find_element(By.XPATH, '//*[@id="sourceDir"]')

                lista_archivos_descargados = []
                for formato in formatos_tienda:
                    # limpiamos el campo de búsqueda
                    input_element.clear()
                    input_element.send_keys(formato)

                    # espera que la pantalla gris desaparezca
                    wait.until(EC.invisibility_of_element((By.CLASS_NAME, "ext-el-mask")))

                    #ruta de la carpeta de descargas
                    download_path = os.path.join(os.path.expanduser("~"), "Downloads")

                    #archivos xlsx existentes en la carpeta de descargas
                    archivos = os.listdir(download_path)
                    archivos_descargados_antes = [(f, os.path.getmtime(os.path.join(download_path, f))) for f in os.listdir(download_path)]

                    # Desplazarse y hacer clic en la opción de Excel
                    elemento_excel = wait.until(EC.presence_of_element_located((By.XPATH, '//li[@title="Excel Spreadsheet"]')))
                    driver.execute_script("arguments[0].scrollIntoView(true);", elemento_excel)
                    time.sleep(5)
                    elemento_excel.click()
                    print('Clic en boton de Excel')
                    print("Descarga iniciada...\n")
                    time.sleep(5)

                    try:
                        elementos_ok = driver.find_elements(By.XPATH, '//button[contains(@class, "x-btn-text") and text()="OK"]')
                        if elementos_ok:
                            boton_ok = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[contains(@class, "x-btn-text") and text()="OK"]')))
                            boton_ok.click()
                            print("Click en ok por muchas alarmas\n")
                    except TimeoutException:
                        pass

                    time.sleep(5)

                    for i in range(20):
                        #archivos xlsx existentes en la carpeta de descargas
                        archivos_descargados_nuevos = [(f, os.path.getmtime(os.path.join(download_path, f))) for f in os.listdir(download_path)]

                        # Obtener el archivo más reciente
                        archivo_reciente, tiempo_modificacion = max(archivos_descargados_nuevos, key=lambda x: x[1])

                        if len(archivos_descargados_nuevos) > len(archivos_descargados_antes) and archivo_reciente.endswith(".xlsx"):
                            break
                        time.sleep(5)

                    path_archivo_descargado = f'{download_path}\\{archivo_reciente}'
                    lista_archivos_descargados.append(path_archivo_descargado)

            except Exception as e:
                print("Fallo la descarga de alarmas desde Connect+.\n")
                error_message = str(e)
                send_mail_app_escritorio(int(2), lista_correo_errores, lista_correo_errores, f'Fallo en el intento {i+1} en descarga de alarmas de Copeland', error_message, lista_correo_errores)
                if i > 2:
                    return []
                continue

        finally:
            #cerrar navegador
            driver.quit()
            print("Proceso terminado.\n")

        return lista_archivos_descargados

if __name__ == "__main__":
    formato_tienda = ['MBG', 'BGA']
    lista_correo_errores = ['marcoantonio.tapia@copeland.com']

    path_archivo_descargado = extraer_alarmas_connect( formato_tienda, lista_correo_errores)
    print(f'Archivos descargados: {path_archivo_descargado}')