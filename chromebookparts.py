from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import requests
import json
from time import sleep
import csv
import pandas as pd

def wait_and_find_element(driver, by, value, timeout=20):
    """Wait for an element to be present and return it."""
    return WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((by, value))
    )

def wait_and_find_elements(driver, by, value, timeout=20):
    """Wait for elements to be present and return them."""
    return WebDriverWait(driver, timeout).until(
        EC.presence_of_all_elements_located((by, value))
    )
def safe_click(element):
    """Click on an element safely, retrying if a StaleElementReferenceException occurs."""
    for _ in range(3):  # Retry up to 3 times
        try:
            element.click()
            return True  # Click successful
        except StaleElementReferenceException:
            print("StaleElementReferenceException encountered. Retrying...")
            continue  # Retry clicking the element
    return False  # Click failed after retries
# Initialize the WebDriver with Chrome
driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))

try:
    # Navigate to the desired URL
    driver.get('https://go.boarddocs.com/wa/bisd/Board.nsf/Public')
    
    # Wait for the button to be clickable and then click it
    button = wait_and_find_element(driver, By.CSS_SELECTOR, "a[href='#tab-meetings']")
    if safe_click(button):
      time.sleep(2)
      # Wait for the meeting accordion to be present
      element = wait_and_find_element(driver, By.ID, 'meeting-accordion')
      
      selections = wait_and_find_elements(driver, By.TAG_NAME, "section")
      wrapYear = wait_and_find_elements(driver, By.CLASS_NAME, 'wrap-year')
      for index, selection in enumerate(selections):
          # Create a new dictionary for each selection
          # item1 = {"year": selection.text}
        year = selection.text
        if selection.get_attribute("aria-expanded") == "false":
            # If already expanded, we don't need to click again
            if index != 0:
              if safe_click(selection):
                time.sleep(2)
        meetings = wrapYear[index].find_elements(By.TAG_NAME, 'a')
        WebDriverWait(driver, 20).until(
            EC.visibility_of(wrapYear[index])  # Wait until the wrapYear is visible after clicking
        )
        meetings = wrapYear[index].find_elements(By.TAG_NAME, 'a')
        data=[]
        for subindex, meeting in enumerate(meetings):
          temp={}
          meeting.click()
          time.sleep(2)
          temp['title']=''
          temp['date']=''
          temp['description']=''
          temp['agendas'] = []
          meeting_content = wait_and_find_element(driver, By.ID, 'meeting-content')
          temp['title'] = temp['title'] = wait_and_find_element(meeting_content, By.CLASS_NAME, 'meeting-name').text
          temp['date'] = meeting_content.find_element(By.CLASS_NAME, 'meeting-date').text
          temp['description'] = meeting_content.find_element(By.CLASS_NAME, 'meeting-description').text
          
          agenda_btn = meeting_content.find_element(By.ID, 'btn-view-agenda')
          if safe_click(agenda_btn):
            time.sleep(2)
            agenda_content = wait_and_find_element(driver, By.ID, 'agenda')
            categories = agenda_content.find_elements(By.CLASS_NAME, 'wrap-category')
            for category in categories:
              items = category.find_elements(By.TAG_NAME, 'li')
              for item in items:
                agenda = {
                        'meeting': '',
                        'category': '',
                        'subject': '',
                        'type': '',
                        'description': '',
                        'files': []
                    }
                if safe_click(item):
                  time.sleep(2)
                  item_content = WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located((By.ID, 'view-agenda-item'))
                  )
                  public_body_div = wait_and_find_element(driver, By.XPATH, "//div[@key='publicbody']")
                  container = item_content.find_element(By.CLASS_NAME, 'container')
                  rows = container.find_elements(By.TAG_NAME, 'dl')
                  agenda['meeting']=rows[0].find_element(By.TAG_NAME, 'dd').text
                  agenda['category']=rows[1].find_element(By.TAG_NAME, 'dd').text
                  agenda['subject']=rows[2].find_element(By.TAG_NAME, 'dd').text
                  agenda['type']=rows[3].find_element(By.TAG_NAME, 'dd').text
                  agenda['description'] = public_body_div.text
                  file_tags = []
                  file_tags = container.find_elements(By.CLASS_NAME, 'public-file')
                  for file_tag in file_tags:
                    try:
                      file_url = file_tag.get_attribute('href')
                      file_name = file_tag.text
                      agenda['files'].append({'url': file_url, 'name': file_name})
                      file_response = requests.get(file_url)
                      file_content = file_response.content
                      with open(file_name, 'wb') as file:
                        file.write(file_content)
                    except Exception as e:
                      print(f"Error retrieving file URL: {e}")
                      continue
                  temp['agendas'].append(agenda)
            returnBtn = WebDriverWait(driver, 20).until(
              EC.element_to_be_clickable((By.CSS_SELECTOR, "a[href='#tab-meetings']"))
            )
            returnBtn.click()
            time.sleep(2)
            data.append(temp)
        rows = []
        for item in data:
            title = item['title']
            date = item['date']
            description = item['description']
            for agenda in item['agendas']:
                rows.append({
                    'title': title,
                    'date': date,
                    'description': description,
                    'meeting': agenda['meeting'],
                    'category': agenda['category'],
                    'subject': agenda['subject'],
                    'type': agenda['type'],
                    'agenda_description': agenda['description'],
                    'files': ', '.join(map(str, agenda['files'])) if len(agenda['files']) > 0 else ''
                })

        # Create a DataFrame from the flattened data
        df_agendas = pd.DataFrame(rows)

        # Save to Excel
        df_agendas.to_excel(year+'-school_board_meeting.xlsx', index=False)
      # with open('ssss.json', 'w', encoding='utf-8') as file:
      #   json.dump(data, file, ensure_ascii=False, indent=4)
finally:
    # Close the WebDriver
    driver.quit()