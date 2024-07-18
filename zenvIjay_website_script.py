import csv
import time
from seleniumwire import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

INPUT_FILE_NAME = 'input.csv'
OUTPUT_FILE_NAME = 'output.csv'
DIRECTOR_ARRAY = []
WAIT_DELAY = 3

#grabs ID's from first column of input file
def grab_input_IDs():
    id_array = []
    try:
        with open(INPUT_FILE_NAME, newline='') as csvfile:
            csv_IDs = csv.reader(csvfile, delimiter=' ')
            for row in csv_IDs:
                id_array.append(row[0])
    except Exception as error:
        print(error)
    return id_array

#prints the information gathered from website on each row
def print_results(file_name=None):
    out_file_name = OUTPUT_FILE_NAME
    if file_name is not None:
        out_file_name = file_name
    try:
        with open(out_file_name, 'w', newline='') as csvfile:
            writer = csv.writer(csvfile, delimiter=',')
            writer.writerow(['CIN', 'Registration Number', 'DIN', 'Director Name', 'Director Designation'])
            writer.writerows(DIRECTOR_ARRAY)
    except Exception as error:
        print(error)

#disables API which prevents website scrapping
def interceptor(request):
    if "disable-devtool" in request.path:
        request.abort()

#waits for wait delay time until desired element is found
def wait_til_ready(driver, selector_value: str):
    try:
        component = WebDriverWait(driver,WAIT_DELAY).until(EC.presence_of_element_located((By.CSS_SELECTOR, selector_value)))
    except TimeoutException:
        print_results(file_name='interrupted_output.csv')
        raise TimeoutException("Took too long to grab information. Please check internet connection and computer speeds. Try again.")
    return component

#solves the website's captcha
#reads the logs to find the captcha solution and enters it
def solve_captcha(driver: webdriver.Chrome):
    captcha_input = wait_til_ready(driver, '#customCaptchaInput')
    captcha_button = wait_til_ready(driver, '#check')

    captcha_input.send_keys('0')
    captcha_button.click()

    logs = driver.execute_script("return console.logs")
    captcha_answer = -1
    for log in logs:
        if len(log) == 2:
            if 'in validatecaptcha' in log[0]:
                captcha_answer = log[1]
    if captcha_answer == -1:
        raise Exception("Error, captcha result not found")
    
    captcha_input.clear()
    captcha_input.send_keys(captcha_answer)
    captcha_button.click()
    time.sleep(1)


#navigates the website
#deals with captchas, searching, and grabbing information
def navigate_website(driver: webdriver.Chrome, id):
    driver.get('https://www.mca.gov.in/content/mca/global/en/mca/master-data/MDS.html')
    driver.implicitly_wait(20)

    driver.execute_script("""
    console.stdlog = console.log.bind(console);
    console.logs = [];
    console.log = function(){
        console.logs.push(Array.from(arguments));
        console.stdlog.apply(console, arguments);    
    }
    """)
    
    solve_captcha(driver)

    master_input = wait_til_ready(driver, '#masterdata-search-box')
    master_input.clear()
    time.sleep(0.5)
    master_input.send_keys(id)
    time.sleep(0.5)
    master_input.send_keys(Keys.ENTER)
    time.sleep(0.5)
    solve_captcha(driver)    

    option_table = wait_til_ready(driver, 'table.table:nth-child(8)')
    if len(option_table.find_elements(by=By.CSS_SELECTOR, value='tr')) < 2:
        master_input.clear()
        time.sleep(0.5)
        master_input.send_keys(id)
        time.sleep(0.5)
        master_input.send_keys(Keys.ENTER)
        time.sleep(0.5)
        solve_captcha(driver)    

    #ensures enough attempts/time is given to going to new webpage
    retry_counter = 0
    while 'company-master-info' not in driver.current_url:
        if retry_counter > 3:
            raise Exception('Cannot get to company master info page.')
        select_result = wait_til_ready(driver,'table.table:nth-child(8) > tbody:nth-child(2)')
        select_result.click()
        time.sleep(0.5)
        solve_captcha(driver)
        retry_counter += 1
    
    registration_num = wait_til_ready(driver, '#registrationNumber').text

    director_tab = wait_til_ready(driver, '#formId > button.tablinks.directorData')
    director_tab.click()
    time.sleep(0.5)

    #formats direction information into an array
    cell_counter = 0
    director_table = wait_til_ready(driver, '#content')
    for row in director_table.find_elements(by=By.CSS_SELECTOR, value='tr'):
        director_info = []
        director_info.append(id)
        director_info.append(registration_num)
        for cell in row.find_elements(by=By.CSS_SELECTOR, value='td'):
            if cell_counter > 0 and cell_counter < 4: #DIN, name, designation
                director_info.append(cell.text)
            cell_counter += 1
        DIRECTOR_ARRAY.append(director_info)
        cell_counter = 0

if __name__ == '__main__':
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--log-level=3")
    driver = webdriver.Chrome(chrome_options=chrome_options)
    driver.request_interceptor = interceptor
    id_list = grab_input_IDs()
    for id in id_list:
        try:
            navigate_website(driver, id)
        except Exception as error:
            print_results('output_interrupted.csv')
            raise Exception(error)

    print_results()

    driver.close()
        