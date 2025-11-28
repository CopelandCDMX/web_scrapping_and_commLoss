import src. utils_scraping as scraping
import os
import yaml
import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
import logging 
from logging.handlers import RotatingFileHandler
from pathlib import Path

Path("logs").mkdir(exist_ok=True)

handler = RotatingFileHandler(
    "logs/scraping.log",
    maxBytes= 5_000_000, #5mb
    backupCount= 5
)

logging.basicConfig(
    handlers=[handler],
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)


BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
logging.info("base dir: %s", BASE_DIR)

### config.yaml has the relative paths of drivers, credentials, and where the data is stored
CONFIG_PATH = os.path.join(BASE_DIR, "config.yaml")
with open(CONFIG_PATH, "r") as f:
    config = yaml.safe_load(f)

###build absolute paths at runtime, you can change the paths/file in config.yaml
driver_path = os.path.join(BASE_DIR, config["driver"]["path"])
lists_downloaded_path = os.path.join(BASE_DIR, config["data"]["lists_downloaded"])
lists_downloaded_path = os.path.abspath(lists_downloaded_path)
results_path = os.path.join(BASE_DIR, config["data"]["results"])
all_stores_list = os.path.join(BASE_DIR, config["data"]["all_stores_fixed_list"] )
credentials_file = os.path.join(BASE_DIR, config["credentials"]["path"] )


####function for download
##changes the file name at the end of the download

###function for elt that will run after the two downloads take place
def run_scraping(previous_days= 1):
    """Function that runs the web scrapping process. Downloads a list, and renames it, in the directory list_downloaded_path.
    -----------------
    Params:
    previous_days, int. 
        Number of previous days for which the logs will be downloaded from the date that the function is run.
    -----------------------------------
    Returns:
    None
    """
    mail_list = ["adrian.perez1@copeland.com"]
    #print(lists_downloaded_path)
    logging.info('Scrapping script started')

    scraping.extraer_alarmas_connect(
        previous_days=previous_days, 
        driver_path=driver_path, 
        credentials_path=credentials_file,
        downloads_path=lists_downloaded_path,
        lista_correo_errores=mail_list
    )
    logging.info("Changing the name of the download")
    new_name = scraping.new_filename()
    scraping.rename_downloaded_file(download_dir=lists_downloaded_path, new_name=new_name)
    logging.info("Scraping process successful!!")
    return None

if __name__ == "__main__":
    run_scraping()    