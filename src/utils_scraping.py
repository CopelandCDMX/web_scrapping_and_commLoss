from selenium import webdriver 
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service # i added this line
from selenium.webdriver.support.ui import Select
import os
import time
import datetime
import win32com.client as win32
from pathlib import Path
import logging


def get_chrome_options( download_dir: str):
    """This function defines pre-defined options for the Google chrome driver
    ------------
    Params:
    download_dir: str. directory where the downloads are going to be stored.
    -----------------
    Returns:
    chrome_options. Selenium.webdriver.chrome.options 
        Options for the Chrome driver
    """
    chrome_options = Options()
    ### add the language argument
    chrome_options.add_argument("--lang=en")
    ###add experimental options for file downloads
    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    })
    return chrome_options

def new_filename():
    """Generates a new filename based on the current date and time."""
    now = datetime.datetime.now()
    return f"Report_{now.strftime('%Y%m%d_%H%M')}.xlsx"


def rename_downloaded_file(download_dir: str, new_name: str):
    """Given a directory, this function finds the last file created in download_dir and changes the name 'new_name'.
    This function is designed to run after a download has taken place.  
    -------------------------------
    Params

    download_dir: str.
        path of directory where the latest file created is stored. 
    new_name: str.
        new name of the file.
    ----------------------------------------------
    Returns
    None
    """
    #Find the newest file in the directory
    # List all files and sort them by creation time (newest first)
    files = [os.path.join(download_dir, f) for f in os.listdir(download_dir)]
    files.sort(key=os.path.getmtime, reverse=True)
    
    if files: ## not empty list
        original_filepath = files[0] # Assumes the newest file is the one you just downloaded
        new_filepath = os.path.join(download_dir, new_name)
        #Rename the file
        try:
            os.rename(original_filepath, new_filepath)
            print(f"Successfully renamed '{os.path.basename(original_filepath)}' to '{new_name}'")
        except FileNotFoundError:
            print("Error: Downloaded file not found.")
        except Exception as e:
            print(f"Error renaming file: {e}")
    else:
        print("No files found in the download directory.")


def send_mail_app_escritorio(importancia: int, destinatario: list, copia:list , subjet: list , cuerpo: str, lista_correo_errores:list):
    """Function that sends an email. It's going to be used in case something goes wrong
    --------------
    Params

    importancia, int. 
        Describes importance of the email
    destinatario, list.
        list of people to send the alarm
    copia, list.
        list of people to copy the email
    subject, list.
        subject of the email
    cuerpo., str.
        String with the message included in the email
    lista_correo_errores, list.
        list of people to send the error message.
    ---------------------------------------
    Returns

    None
    """
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
        return None
    except Exception as e:
        error_message = str(e)

        send_mail_app_escritorio(int(2), lista_correo_errores, lista_correo_errores, 'Fallo en envío de alarmas de Copeland', error_message)
        print('Error al enviar correo:',repr(e))
        print(f'Error: {e}\n')
        time.sleep(5)
        return None


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
        None
    """
    
    ###goes to the required frame
    sesion_frame = wait.until(EC.presence_of_element_located((By.ID, "navId")))
    driver.switch_to.frame(sesion_frame)

    ## waits for the userid and password fields
    username_input = wait.until(EC.presence_of_element_located((By.NAME, "_UserName")))
    password_input = wait.until(EC.presence_of_element_located((By.NAME, "_Password")))
    login_button = wait.until(EC.element_to_be_clickable((By.NAME, "loginButton")))
    #print('found input fields\n')
    
    
    k_path = file_path  ## this is where the credentials are stored
    #print(k_path)  ### sanity check
    with open(k_path, 'r', encoding='utf-8') as file:
        lineas = file.readlines()

    #print('Attempting login...\n')
    logging.info("Attempting login")
    # first line is username
    # second line is password
    u = lineas[0].strip()
    p = lineas[1].strip()

    username_input.send_keys(u)
    password_input.send_keys(p)
    login_button.click()

    driver.switch_to.default_content()
    #print("Inicio de sesión exitoso!\n")
    
    try:
        element = driver.find_element(By.CLASS_NAME, "invalidFields")
        if "Password will expire in" in element.text:
            login_button.click()
            send_mail_app_escritorio(
                importancia=2,
                destinatario=lista_correo_errores, 
                copia=lista_correo_errores, 
                subjet='Contraseña de cuenta de descarga de alarmas por expirar', 
                cuerpo=f'La contraseña de la cuenta uestá por expirar', 
                lista_correo_errores=lista_correo_errores)
        else:
            logging.info("El elemento existe, pero el texto no coincide")
            #print("El elemento existe, pero el texto no coincide.")
            time.sleep(2)
    except:
        pass


#### need to describe better this function
def extraer_alarmas_connect(previous_days: int, driver_path: str, credentials_path: str, downloads_path:str,  lista_correo_errores: list):
    """ This function downloads the lists for the commLoss process using selenium. It follows a very specific process based on the connect+ 
    webpage and the list that wants to be downloaded. It storages the list in the directory given in the get_chrome_options function. 
    This function assumes the following directory structure:
    
    |-- chromedriver-win64/
    |-- data/
    |   |-- lists_downloaded/
    |   |-- results/
    |   |-- all_stores_list/
    |-- src/
    |   |-- main_etl.py
    |   |-- main_scraping.py
    |   |-- utils_etl.py
    |   |-- utils_scraping.py 
    |-- config.yaml
    |-- credentials.txt
    |-- main.bat
    -----------------------------------
    Params

    previous_day, int.
        This parameter specifies the number of days to be considered to download the lists. It will download the data with range
        [current_time - previous_days, current_time]
    driver_path, str. 
        Path of where the driver is located
    credentials_path, str.
        Path to the file where the credentials are located.
    downloads_path, str.
        Path to the directory where the downloaded lists are stored. 
    lista_correo_errores, list.
        If there is an error in the download process, it will send an email to the people in the list.
    ------------------------------------------
    Returns
    None
    """
    for i in range(3):
        #path_programa = os.getcwd()   ## path where the code is living
        #driver_path = os.path.join(path_programa, "chromedriver-win64", "chromedriver.exe")
        #credentials_path = os.path.join(path_programa, "credentials.txt")

        #service = webdriver.ChromeService(driver_path)  ### I changed this line
        service = Service(executable_path=driver_path) ### updated line
        driver = webdriver.Chrome(service=service, options=get_chrome_options(download_dir=downloads_path))
        driver.maximize_window()
        zoom_lever = '50%'
        driver.execute_script(f'document.body.style.zoom="{zoom_lever}"')

        #tiempo de espera
        wait = WebDriverWait(driver, 20)

        try:
            # Abrir la página web
            #url_walmart = "https://walmartca.my-connectplus.com/walmartca/"   ###### una vez aqui, probably i need to put this as a parameter
            url_cpla = "https://cpla.my-connectplus.com/cpla/"     ###### una vez aqui, enfoque en este
            driver.get(url_cpla)
            #driver.get(url_walmart)
            wait.until(EC.presence_of_element_located((By.TAG_NAME, "html")))
            #print("Página cargada\n")
            logging.info("Webpage loaded  https://cpla.my-connectplus.com/cpla/")
            
            inicio_pasword(file_path= credentials_path, driver=driver, wait=wait, lista_correo_errores=lista_correo_errores)
            logging.info("Succesfully log in")

        except Exception as e:
            error_message = str(e)
            #print('Error during login:', str(e))
            logging.error("Error during login: %s", e)
            send_mail_app_escritorio(int(2), lista_correo_errores, lista_correo_errores, 'No se encontró el formulario de inicio de sesión.', error_message, lista_correo_errores)
            print("Fallo el inicio de sesión.\n")
            driver.quit()
            exit()

        try:
            #print('Enter try to click the button to go to the data section\n')
            driver.switch_to.default_content()
            sesion_frame = wait.until(EC.presence_of_element_located((By.ID, "navId")))
            driver.switch_to.frame(sesion_frame)

            #wait.until(EC.element_to_be_clickable((By.XPATH, '//img[@title="Alarm"]'))).click() #<img class="header-img" src="/cpla/Images/alarm.png" title="Alarm">
            wait.until(EC.element_to_be_clickable((By.XPATH, '//img[@title="Graph/Watch"]'))).click() #<img class="header-img" src="/cpla/Images/graph.png" title="Graph/Watch">
            #wait.until(EC.element_to_be_clickable((By.XPATH, "//img[@id='clearCheckboxesButton']/following-sibling::img[@title='Graph/Watch']"))).click()
            #print('Clicked Graph/Watch image')
            logging.info("clicked Graph/Watch")
            # espera que la pantalla gris desaparezca
            #wait.until(EC.invisibility_of_element((By.CLASS_NAME, "ext-el-mask"))) 
            
            #################### selecting +Tiendas CommLoss from the dropdown #########################
            driver.switch_to.default_content()
            sesion_frame = wait.until(EC.presence_of_element_located((By.ID, "mainId")))
            driver.switch_to.frame(sesion_frame)
            #find the dropdown element by its id    
            dropdown_element = wait.until(EC.presence_of_element_located((By.ID, "listsselection")))
            
            #driver.find_element(By.ID, "listselection")
            
            #pass the located dropdown element to the select class constructor
            select_object = Select(dropdown_element)
            wait.until(
                EC.presence_of_element_located((By.XPATH, '//option[@value="599"]'))
            )
            select_object.select_by_value("599")
            #select_object.select_by_visible_text("+Tiendas CommLoss")
            #<option value="599" class="publiclistoption">+Tiendas CommLoss</option>

            ############## SELecting button Retrieve Logs  + Export ######################
            ##goes to the right frame
            driver.switch_to.default_content()
            sesion_frame = wait.until(EC.presence_of_element_located((By.ID, "mainId")))
            driver.switch_to.frame(sesion_frame)

            button_element = wait.until(
                EC.element_to_be_clickable((By.ID, "exportLogsButton"))
            )
            button_element.click()
            ##<button type="button" id="exportLogsButton" class="button contentNormal" onclick="showDateRangeSelection(true)">
                        # Retrieve Logs + Export
                     #</button>
            
            
            ##################### Select option Condense from the dropdown menu (stays in the same frame)########################
            dropdown_element = wait.until(EC.presence_of_element_located((By.ID, "exportFormat")))
            select_object = Select(dropdown_element)
            wait.until(
                EC.presence_of_element_located((By.XPATH, '//option[@value="2"]'))
            )
            select_object.select_by_visible_text("Condensed")
            ##<select id="exportFormat" class="controls">
		            #<option value="1">Comprehensive</option>
		            #<option value="2">Condensed</option>
		          #</select>

            ######################### start time and end time (stays in the same frame) ##########################
        
            now = datetime.datetime.now()
            end_time = now.strftime("%Y-%m-%d %H:%M:%S")
            
            initial_date = now - datetime.timedelta(days=previous_days)
            initial_date_str = initial_date.strftime("%Y-%m-%d %H:%M:%S")

            ### start time    
            start_time_input_field = wait.until(
                EC.presence_of_element_located((By.ID, "startTimeField"))
            )
            start_time_input_field.clear()
            start_time_input_field.send_keys(initial_date_str)
            #<input id="startTimeField" name="startTimeField" class="controls" type="text">
            
            ####end time 
            end_time_input_field = wait.until(
                EC.presence_of_element_located((By.ID, "endTimeField"))
            )
            end_time_input_field.clear()
            end_time_input_field.send_keys(end_time)
            #<input id="endTimeField" name="endTimeField" class="controls" type="text">

            ################## Click on Go button and download  (still in the same frame)#######################
            button_element = wait.until(
                EC.element_to_be_clickable((By.ID, "userDownloadStart"))
            )
            button_element.click()
            

            timeout = 600 ## ten minutes tolerance for downloads
            start = time.time()
            downloads_path = Path(downloads_path)
            
            #snapshop of files before download
            before = set(downloads_path.glob("*.xlsx"))

            #wait until at least one new file appears

            new_files = set()
            while time.time() - start < timeout:
                after = set(downloads_path.glob("*.xlsx"))
                new_files = after - before
                if new_files:
                    break
                time.sleep(20)
            else:
                raise TimeoutError(f"No new .xlsx files appeared after {timeout} seconds")
              
              
              
              
              # files = list(downloads_path.glob("*.xlsx"))
              # if files and not any(f.suffix == ".crdownload" for f in files):
                   #break
               #time.sleep(5)
            ### waits two minutes for the download to take place, Need to check how to improve the waiting for the download part 
            #time.sleep(300)
            ##<button id="userDownloadStart" type="button" class="controls dialogButton">Go</button>#


            #print('Closing the browser...')
            logging.info("Closing the browser...")
            driver.quit()
            return None

            #wait.until(EC.element_to_be_clickable((By.XPATH, '//img[@title="Alarm"]'))).click() #<img class="header-img" src="/cpla/Images/alarm.png" title="Alarm">
            #print('Clic en boton alarmas')
        except Exception as e:
                #print("Fallo la descarga de alarmas desde Connect+.\n")
                error_message = str(e)
                #print(str(e))
                logging.error("Fallo la descarga de alarmas desde Connect+: %s", e)
                send_mail_app_escritorio(int(2), lista_correo_errores, lista_correo_errores, f'Fallo en el intento {i+1} en descarga de alarmas de Copeland', error_message, lista_correo_errores)
                if i > 2:
                    return None
                continue  ################
                ##return None

