import src.utils_etl as etl
import os
import yaml
import warnings
import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")


Path("logs").mkdir(exist_ok=True)

handler = RotatingFileHandler(
    "logs/etl.log",
    maxBytes= 5_000_000, #5mb
    backupCount= 5
)

logging.basicConfig(
    handlers=[handler],
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)



BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
print(BASE_DIR)

### config.yaml has the relative paths of drivers, credentials, and where the data is stored
CONFIG_PATH = os.path.join(BASE_DIR, "config.yaml")
with open(CONFIG_PATH, "r") as f:
    config = yaml.safe_load(f)

###build absolute paths at runtime, you can change the paths/file in config.yaml
driver_path = os.path.join(BASE_DIR, config["driver"]["path"])
lists_downloaded_path = os.path.join(BASE_DIR, config["data"]["lists_downloaded"])
results_path = os.path.join(BASE_DIR, config["data"]["results"])
all_stores_list = os.path.join(BASE_DIR, config["data"]["all_stores_fixed_list"] )
#credentials_file = os.path.join(BASE_DIR, config["credentials"]["path"] )


def run_etl():
    """This function takes the last two created lists in list_downloaded and do the etl procees give a diagnostic about CommLoss"""
    hours_considered = 1
    logging.info("Starting etl process")
    etl.online_offline_process(
        folder_connectplus_downloads_path= lists_downloaded_path, 
        table_all_stores_path= all_stores_list, 
        final_storage_path= results_path,
        N=hours_considered)
    
if __name__ == "__main__":
    run_etl()    





