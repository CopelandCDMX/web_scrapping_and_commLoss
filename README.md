# selenium_web_scrapping
This repository does webscraping and a etl processs that allows to detect Communication Loss for stores in an specific given list. 


For the webscraping part, the drivers can be found at:
https://googlechromelabs.github.io/chrome-for-testing/#stable

## structure of the repository

    |-- chromedriver-win64/
    |-- data/
    |   |-- lists_downloaded/
    |   |-- results/
    |   |-- all_stores_list/
    |-- logs/
    |-- src/
    |   |-- __init__.py
    |   |-- main_etl.py
    |   |-- main_scraping.py
    |   |-- utils_etl.py
    |   |-- utils_scraping.py 
    |-- config.yaml
    |-- credentials.txt
    |-- main.bat
