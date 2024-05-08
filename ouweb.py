from selenium import webdriver
from selenium.webdriver.common.by import By
import openpyxl
from selenium.common.exceptions import NoSuchElementException
# from selenium.webdriver.support import expected_conditions as EC
# import time
# from selenium.webdriver.common.action_chains import ActionChains
# from selenium.webdriver.common.keys import Keys
# from selenium.common.exceptions import NoSuchElementException

driver = webdriver.Firefox()

for i in range(100523733001, 100523733080):
    try:
        driver.get("http://202.63.117.72/result_april_2024/11/index.php")
        driver.implicitly_wait(0.5)

        wb = openpyxl.load_workbook("result.xlsx")

        ws = wb.active
        ws.append([])

        text_box = driver.find_element(By.NAME, value="htno")
        submit_button = driver.find_element(By.NAME, value="submit")
        text_box.send_keys(i)
        submit_button.click()
        Name = driver.find_element(By.XPATH, '/html/body/table[2]/tbody/tr[2]/td/table/tbody/tr[3]/td[2]').text
        RollNo = driver.find_element(By.XPATH, '/html/body/table[2]/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div' ).text
        CGPA = driver.find_element(By.XPATH, '/html/body/table[2]/tbody/tr[2]/td/span/table[2]/tbody/tr/td[2]' ).text
        ws.append([Name,RollNo,CGPA])
        mytable = driver.find_element(By.XPATH, '/html/body/table[2]/tbody/tr[2]/td/span/table[1]')

        data= []
        norow = 0
        for row in mytable.find_elements(By.CSS_SELECTOR, 'tr'): 
            norow +=1 
            for cell in row.find_elements(By.TAG_NAME,'td'):
                data.append(cell.text)
            ws.append(data)
            data = []
        if norow == 1:
            ws.append(["no result"]) 
        wb.save("result.xlsx")



        wb.save("result.xlsx")
    except NoSuchElementException:
        pass
