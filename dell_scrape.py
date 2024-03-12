from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
from pprint import pprint
from datetime import datetime 
from queue import Queue
from threading import Thread
from selenium.webdriver.common.keys import Keys
import traceback
import json


class Laptop:
    def __init__(self, serial_num, asset_id):
        self.serial_num = serial_num
        self.asset_id = asset_id
        self.waranty = None

    def set_waranty(self, date_str: str) -> None:
        '''
            NOTE
            when orignally fetching the data from the website
            the <span> tag had the text "Expires" in front of the date
            so I had to remove it and then parse the date
            
            the date is in the format of "DD MMM YYYY" which should work
            in onelist IF IT DOESNT => just re-open the json and re-construct the objects
            in and alter the strptime method into the appropriate format
        '''
        if date_str.startswith("Expires"):
            date_str = date_str.replace("Expires ", "")
        self.waranty = datetime.strptime(date_str, "%d %b %Y").date()
        
    def __str__(self):
        return f'Serial #{self.serial_num}\nAsset ID:{self.asset_id}\nWaranty Expiration:{self.waranty.strftime("%m/%d/%Y") if self.waranty else "None"}\n\n\n'

    def to_dict(self) -> dict:
        '''
                    NOTE
            [!] DO NOT USE THE UNPACKING OPERATOR - IT IS DESIGNED TO BE A CLEAN EXPORT TO AN EXCEL FILE
            [!] DO NOT USE A DATA CLASS - IT WILL NOT WORK (the warranty is not set at object instantion)
        '''
        return {
            'Serial Number': self.serial_num,
            'Asset ID': self.asset_id,
            'Waranty Expiration': self.waranty.strftime("%m/%d/%Y") 
        }
    
    
    
    
        

def get_data(excel_path) -> list[Laptop]:
    '''
        parses the asset id & serial number into object
        representation for the laptops in the excel file
        used to scrape from the dell website 
    '''
    # excel_path = 'data.xlsx'    
    df = pd.read_excel(excel_path)
    dell_assets = df[df['Manufacturer'].str.contains('Dell', case=False, na=False)]
    dell_laptops_info = dell_assets[['Asset ID', 'Serial number']].set_index('Asset ID').to_dict()['Serial number']
    obj_repr = [Laptop(serial_num=serial_num, asset_id=asset_id) for asset_id, serial_num in dell_laptops_info.items()]
    return obj_repr



def load_driver() -> webdriver.Chrome:
    '''
        NOTE
        loads the several instances of the chrome data 
        used to perform the scraping, if you end up re-running the script
        it'll take around 10mins to do everything since the driver isn't
        run in headless mode.
        
        If you do re-run it I highly consider adding headless mode as an option
        if your hardware isn't the best or you might max out your CPU lol
    '''
    options = webdriver.ChromeOptions()
    options.page_load_strategy = 'normal'
    driver = webdriver.Chrome()
    return driver


def process_laptop(laptop:Laptop) -> None:
    '''
        After doing some digging in the chrome console
        I found the url used to make the API call however there
        is a random delay, so you have to click on the More Details button
        on the URL to get the expiration date & then scrape the waranty info.
    '''
    
    print(f'*** Processing Laptop with AST_ID of { laptop.asset_id }***\n')
    driver = load_driver()
    url = f'https://www.dell.com/support/productsmfe/en-us/productdetails?selection={ laptop.serial_num }&assettype=svctag&appname=warranty&inccomponents=false&isolated=false'
    with driver:
        driver.get(url)
        print('*** Page loaded ***')    
        element = WebDriverWait(driver, 10).until(
            lambda x: x.find_element(By.CSS_SELECTOR, ".dds__button--secondary")
        )
        element.click()
        data = WebDriverWait(driver, 10).until(
            lambda x: x.find_element(By.CSS_SELECTOR, "#ps-inlineWarranty > div.flex-wrap.d-flex.flex-column.flex-lg-row.text-center.text-lg-left.mb-5.mb-lg-1.mt-lg-0.mt-6 > div > p")
        )
        expr_date = data.text.strip()
        laptop.set_waranty(expr_date)
        print(f'*** Waranty expires on { expr_date } ***')
        

        
def thread_worker(task_queue : Queue[Laptop]):
    ''''
        processes all threads using the queue
        generated from excel file.
        
        be careful modifying this or you may cause
        a concurrency issue with the threads 
    '''
    while not task_queue.empty():
        laptop = task_queue.get()
        try:
            process_laptop(laptop)
        except Exception as e:
            print(f'Error: { e }')
            log_error(e, laptop)
        finally:  
            task_queue.task_done()  # Signal that the thread bound to the task is done 

def log_error(e:Exception, laptop) -> None:
    '''
        log and provide stack trace when exceptions occur with file
    '''
    
    with open('error.log', 'a') as f:
        f.write(f'[LAPTOP { laptop.asset_id }]:  falled to proces...\n\n\n')
        f.write(f'[EXCEPTION]: { e }\n')
        f.write(traceback.format_exc())
        f.write('\n\n' + "*"*50)


def queue_objs(laptops: list[Laptop]):
    task_queue = Queue()
    for laptop in laptops:
        task_queue.put(laptop)
    return task_queue


def process_threads(laptops: Queue[Laptop], num_threads: int = 5) -> None:
    threads = []
    print('*** Starting threads ***')
    for _ in range(num_threads):
        thread = Thread(target=thread_worker, args=(laptops,))
        thread.start()
        threads.append(thread)
        
    print('*** Waiting for all threads to finish ***')
    for thread in threads:
        thread.join()


def save_in_json(laptops: list[Laptop]) -> None:
    data = [laptop.to_dict() for laptop in laptops]
    with open('laptops.json', 'w') as f:
        json.dump(data, f, indent=4)
    print('*** Sucessfully saved data collected in laptops.json ***')


def load_json_data(file: str):
    '''
        NOTE
        If you have to adjust the data again,
        pass in 'laptops.json' as the file parameter and then
        it will return a list of dictionaries that you can use to
        re-construct into laptop objects using good_construct method
    '''
    
    with open(file, 'r') as file:
        data = json.load(file)
    return data

def find_null_entries(data:dict):
    '''
        detects when the data in a SINGLE  dictionary  the json 
        or class attr is null and returns all of the 'null_entries'
    '''
    null_entries = []
    for entry in data:
        # Check if any value in the dictionary is None (null in JSON)
        if any(value is None for value in entry.values()):
            null_entries.append(entry)
    return null_entries


def construct_bad(laptops: dict) -> list[Laptop]:
    '''
        When the entry is null for waranty expiration,
        if the Serial Number or Asset ID is the laptop model
        it means that while fetching the data from the excel file
        the field for the entry wasn't there, disregard those as
        you can't find warranty information without the serial number
    '''
    
    item = Laptop(laptops['Serial Number'], laptops['Asset ID'])
    return item

def construct_good(laptops: dict) -> list[Laptop]:
    '''
        use this method if you need to reconstruct the objects 
    '''
    item = Laptop(laptops['Serial Number'], laptops['Asset ID'])
    item.set_waranty(laptops['Waranty Expiration'])
    return item


def update_dates() -> None:
    '''
        the format was orignally in the format of 
        "Expires DD MMM YYYY", I adjusted in the setter
        for the Laptop object & simply re-opened the json
        storing the objects and re-saved it.
        
        the unpacking operator doesnt work since the attributes 
        of the class are not photo copies. so you'll have to to use 
        construct_good method 
    '''
    raw = load_json_data('laptops.json')
    objs = [construct_good(r) for r in raw]
    save_in_json(objs)

def export_json_to_excel(json_path, output_dir) -> None:
    print('[i] Exporting JSON contents to excel... [i]')
    data_frame = pd.read_json(json_path)
    data_frame.to_excel(output_dir, index=False, engine='openpyxl')
    print(f'[i] Sucessfully exported JSON data to { output_dir } [i]')
    

def main() -> None:
    # change to whatever you named the excel file exported from sharepoint
    # excel_path = 'data.xlsx'
    # laptop_data = get_data(excel_path)
    # thread_queue = queue_objs(laptop_data)
    
    # # NOTE This uses FIVE threads by default to process the laptops, 
    # # if your laptop is slow consider passing in a lower value
    # # (it'll take longer to process but it won't explode ur computer)
    # # the thread count parameter is optional (why it isn't included in method call below)
     
    # process_threads(thread_queue)
    # # Lists are reference types, the objects are updated 
    # # and set with the waranty expiration date in process_threads
    # save_in_json(laptop_data)
    
    json_path = 'laptops.json'
    excel_output = 'waranty_info.xlsx'
    export_json_to_excel(json_path, excel_output)   
    
    # when i had to re-format the dates
    # json_data = load_json_data('laptops.json')
    # laptop_objs = [construct_good(entries) for entires in json_data]
    # queue = queue_objs(laptop_objs)
    
    
    
    
    
    
    


if __name__ == '__main__':
    main()