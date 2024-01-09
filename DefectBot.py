import os
import random
import sys
import re
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import ElementClickInterceptedException
import time

#EXECUTE COMMAND BELOW TO RUN CODE
#pip install google-api-python-client google-auth google-auth-httplib2 google-auth-oauthlib selenium webdriver-manager

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

#THIS IS FOUND ON THE GOOGLE SHEET URL
SPREADSHEET_ID = "1v1AWr798rWQZqQ6wGJhppRtkl53N2u0UWuxL23MDijg"
#https://docs.google.com/spreadsheets/d/  --RIGHT-> 1v1AWr798rWQZqQ6wGJhppRtkl53N2u0UWuxL23MDijg <-HERE--  /edit#gid=920483318
                  
credentials = None
spreadsheetDefects = []
websiteDefects = []
websiteData = []
newDefectsFound = []
templist = ['Attic Door Insulation Displacement', 'Missing Carpet', 'Pantry Ceiling Hole', 'Refrigerant Line Not Properly Sealed', 'Unsealed Dryer Vent Termination', 'Missing Condensate Drain Trap', 'Junction Box Cover Plate Missing']
position = 1
websiteURL = ""

def loadCreds():
    global credentials
    if os.path.exists("token.json"):
        credentials = Credentials.from_authorized_user_file("token.json", SCOPES)
    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(r"C:\Users\dlcou\coding\HomeCloudPython\Credentials2.json", SCOPES)
            credentials = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(credentials.to_json())

def loadDefects():
    global spreadsheetDefects
    service = build("sheets", "v4", credentials=credentials)
    sheets = service.spreadsheets().values()
    
    #"AllDefects" is the name of the sheet being scanned
    column_range = "AllDefects!A1:A"  # Replace this with the appropriate column range
    batch_get_request = {
        "ranges": [column_range],
        "majorDimension": "COLUMNS"
    }
    response = sheets.get(spreadsheetId=SPREADSHEET_ID, range=column_range).execute()
    column_values = response['values']
    spreadsheetDefects = [item for sublist in column_values for item in sublist]

def loadProjects():  
    try:
      global templist
      global websiteURL
      loadCreds()
      if(websiteURL == ""):
          websiteURL = input("Enter HomeDefect URL: ")
      driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
      driver.get(websiteURL)

      driver.maximize_window()
      username_input = driver.find_element("id", "username")
      password_input = driver.find_element("id", "password")
      username_input.send_keys("@gethomecloud.com")
      password_input.send_keys("")
      login_button = driver.find_element("xpath", "//button[contains(text(), 'Log In')]")
      login_button.click()
      wait = WebDriverWait(driver, 10)
      x = 0
      position = 1
      withoutProject = 1
      while x < len(newDefectsFound):
          wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.text-left.text-black.self-center.whitespace-normal")))
          element_xpath = f"//div[contains(@class, 'text-left text-black self-center whitespace-normal') and contains(text(), '{newDefectsFound[x]}')]"
          element = driver.find_element(By.XPATH, element_xpath)
          element.click()
          try:
            wait.until(EC.invisibility_of_element_located((By.CSS_SELECTOR, "div.loading")))
            # Now the target element should be clickable
            projectButton = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(@class, 'text-blue-600 text-base underline mr-1 cursor-pointer mb-1')]")))
            projectButton.click()
            driver.refresh()
          except TimeoutException as error:
            withoutProject = 0
            print()
          time.sleep(1)
          driver.execute_script("window.scrollTo(0, 0);")
          # typeOfProject = driver.find_element(By.CSS_SELECTOR, "div.text-black.text-base.mb-6")
          # print(typeOfProject.text)

          # description = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.text-black.text-base.mb-6")))
          # print(description.text)
          if(withoutProject == 1):
            wait = WebDriverWait(driver, 10)
            element1 = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@class='text-black text-base mb-6' and not(p)]")))
            element2 = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@class='flex pb-4 flex-row w-full' and not(p)]")))
            element3 = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@class='text-black text-base mb-6']/p")))
            element4 = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@class='text-black text-base mb-6' and not(p)][1]/following::div[contains(text(), '$')]")))
            element5 = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@class='text-black text-base mb-6' and not(p)][1]/following::div[contains(text(), 'Later') or contains(text(), 'Now')]")))
            element6 = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@class='text-black text-base mb-6'][contains(text(), 'Service Provider') or contains(text(), 'Homeowner')]")))       
            service = build("sheets", "v4", credentials=credentials)
            
            element2_text = element2.text
            second_line = element2_text.split('\n')[1]
            element4_text = element4.text

            numeric_value = element4_text[1:-3]
            try:
                #"NewDefects" is the name of the sheet being modified
                update_requests = [
                    {
                        "range": f"NewProjects!B{position+1}",
                        "values": [[f"{element1.text}"]]
                    },
                    {
                        "range": f"NewProjects!C{position+1}",
                        "values": [[f"{second_line}"]]
                    },
                    {
                        "range": f"NewProjects!D{position+1}",
                        "values": [[f"{element3.text}"]]
                    },
                    {
                        "range": f"NewProjects!F{position+1}",
                        "values": [[f"{numeric_value}"]]
                    },
                    {
                        "range": f"NewProjects!G{position+1}",
                        "values": [[f"{element5.text}"]]
                    },
                    {
                        "range": f"NewProjects!H{position+1}",
                        "values": [[f"{element6.text}"]]
                    }
                ]
                batch_update_request = {
                    "data": [
                        {
                            "range": req["range"],
                            "values": req["values"],
                            "majorDimension": "ROWS"
                        }
                        for req in update_requests
                    ],
                    "valueInputOption": "USER_ENTERED"
                }
                position+=1
                # Execute the batchUpdate request
                response = service.spreadsheets().values().batchUpdate(spreadsheetId=SPREADSHEET_ID, body=batch_update_request).execute()
            except HttpError as error:
              print("API call limit..")
          if(withoutProject == 0):
              position+=1
              withoutProject = 1
              
          
          time.sleep(1)
          link_text = "Home Defects"
          element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, link_text)))
          element.click()
          time.sleep(0.5)
          x+=1
    except TimeoutException as error:
      print()
        

def loadWebsite():
    loadCreds()
    global websiteURL
    global websiteDefects
    websiteDefects = []
    if(websiteURL == ""):
        websiteURL = input("Enter HomeDefect URL: ")
    
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
    driver.get(websiteURL)

    #SAMPLE HOME, USED FOR QUICK C/P
    #"https://app.gethomecloud.com/homes/4cae03ff-4a54-43ff-bb92-ecf4f5f47d7e/defectlist"
    driver.maximize_window()
    try:
      wait = WebDriverWait(driver, 25)
      username_input = driver.find_element("id", "username")
      password_input = driver.find_element("id", "password")
      username_input.send_keys("david@gethomecloud.com")
      password_input.send_keys("Hockey14dl!")
      time.sleep(10)
      login_button_xpath = "//button[contains(text(), 'Log In')]"
      login_button = wait.until(EC.element_to_be_clickable((By.XPATH, login_button_xpath)))
      login_button.click()
      wait = WebDriverWait(driver, 75)
      expand_all_span = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Expand All') and contains(@class, 'hidden md:inline')]")))
      expand_all_span.click()
      wait = WebDriverWait(driver, 10)
    except Exception as e:
      wait = WebDriverWait(driver, 5)

    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.text-left.text-black.self-center.whitespace-normal")))
    defect_elements = driver.find_elements(By.CSS_SELECTOR, "div.text-left.text-black.self-center.whitespace-normal")
    description_elements = driver.find_elements(By.XPATH, "//div[@class='mr-10']")
    priority_elements = driver.find_elements(By.XPATH, '//*[@id="root"]/div[1]/div/div/div[2]/div/div[2]/div/div/div/div[2]/div[*]/div[4]')
    tags_elements = driver.find_elements(By.XPATH, '//*[@id="root"]/div[1]/div/div/div[2]/div/div[2]/div/div/div/div[2]/div[*]/div[6]')
    implications_elements = driver.find_elements(By.XPATH, "//div[@class='mr-10 mt-2 md:mt-0']")
    service_provider_elements = driver.find_elements(By.XPATH, '//*[@id="root"]/div[1]/div/div/div[2]/div/div[2]/div/div/div/div[2]/div[*]/div[7]')
    # print("imp: ", implications_elements)
    # print("descript: ", description_elements)
    # print("priority: ", priority_elements)
    # print("service : ", service_provider_elements)
    services_2d_array = []
    service = build("sheets", "v4", credentials=credentials)
    sheets = service.spreadsheets().values()

    #"ServiceProvider" is the name of the sheet being scanned
    column_range = "ServiceProvider!A2:A" 
    batch_get_request = {
        "ranges": [column_range],
        "majorDimension": "COLUMNS"
    }
    response = sheets.get(spreadsheetId=SPREADSHEET_ID, range=column_range).execute()
    column_values = response['values']
    serviceproviderNum = [item for sublist in column_values for item in sublist]
    service = build("sheets", "v4", credentials=credentials)
    sheets = service.spreadsheets().values()

    #"ServiceProvider" is the name of the sheet being scanned
    column_range = "ServiceProvider!B2:B" 
    batch_get_request = {
        "ranges": [column_range],
        "majorDimension": "COLUMNS"
    }
    response = sheets.get(spreadsheetId=SPREADSHEET_ID, range=column_range).execute()
    column_values = response['values']
    serviceproviderName = [item for sublist in column_values for item in sublist]    
    services_2d_array = list(zip(serviceproviderNum, serviceproviderName))
    for i in range(len(defect_elements)):
        defect = defect_elements[i].text.strip()
        description = description_elements[i].text.replace("Description:", "").strip()
        priority = priority_elements[i + 1].text.strip()
        tags = tags_elements[i + 1].text.strip()
        service_provider_type = service_provider_elements[i + 1].text.strip()
        websiteData.append([defect, description, priority, tags, i, service_provider_type])
        print("Defect: ", defect, "Description: ", description, "Priority: ", priority, "Tags: ", tags, "Service provider: ", service_provider_type)

    target_text = "Implication:"
    counter = 0
    for imp in implications_elements: 
        imp_text = imp.text
        if target_text in imp_text:
            implication_text = imp_text.split(target_text, 1)[-1].strip()
            websiteData[counter][4] = implication_text
            counter += 1
    serviceProviderCounter = 0
    for serviceProvider in websiteData:
        foundService = False
        for service in services_2d_array:
            if(serviceProvider[5] == service[1]):
              foundService = True
              websiteData[serviceProviderCounter][5] = str(service[0])
              serviceProviderCounter = serviceProviderCounter + 1
        if(foundService == False):
            websiteData[serviceProviderCounter][5] = websiteData[serviceProviderCounter][5]
            serviceProviderCounter = serviceProviderCounter + 1
            foundService = True     
    div_elements = driver.find_elements("css selector", "div.text-left.text-black.self-center.whitespace-normal")
    for div in div_elements:
        try:
            content = div.text.strip()
            websiteDefects.append(content)
        except Exception as e:
            print("An error occurred:", e)
    driver.quit()

def fillSpreadsheet():
    loadCreds()
    loadDefects()
    global position
    spreadsheetDefects_set = set(spreadsheetDefects)
    websiteDefects_set = set(websiteDefects)
    uniqueDefectsList = websiteDefects_set - spreadsheetDefects_set
    try:
        service = build("sheets", "v4", credentials=credentials)
        for item in uniqueDefectsList:
            if(item != "Smoke Detector System Needed"):
              print("Defect not in config:", str(item))
              sheetCounter = 0
              for row in websiteData:
                  if(str(item) == str(row[0])):
                      while True:
                          try:
                              #"NewDefects" is the name of the sheet being modified
                              update_requests = [
                                  {
                                      "range": f"NewDefects!B{position+1}",
                                      "values": [[f"{websiteData[sheetCounter][0]}"]]
                                  },
                                  {
                                      "range": f"NewDefects!C{position+1}",
                                      "values": [[f"{websiteData[sheetCounter][1]}"]]
                                  },
                                  {
                                      "range": f"NewDefects!D{position+1}",
                                      "values": [[f"{websiteData[sheetCounter][2]}"]]
                                  },
                                  {
                                      "range": f"NewDefects!E{position+1}",
                                      "values": [[f"{websiteData[sheetCounter][3]}"]]
                                  },
                                  {
                                      "range": f"NewDefects!F{position+1}",
                                      "values": [[f"{websiteData[sheetCounter][4]}"]]
                                  },
                                  {
                                      "range": f"NewDefects!H{position+1}",
                                      "values": [[f"{websiteData[sheetCounter][5]}"]]
                                  }
                              ]
                              batch_update_request = {
                                  "data": [
                                      {
                                          "range": req["range"],
                                          "values": req["values"],
                                          "majorDimension": "ROWS"
                                      }
                                      for req in update_requests
                                  ],
                                  "valueInputOption": "USER_ENTERED"
                              }

                              # Execute the batchUpdate request
                              response = service.spreadsheets().values().batchUpdate(spreadsheetId=SPREADSHEET_ID, body=batch_update_request).execute()
                              newDefectsFound.append(websiteData[sheetCounter][0])
                              sheetCounter = 0
                              position = position + 1
                              break
                          except HttpError as error:
                            print("API call limit..")
                            time.sleep(30)
                      break
                  sheetCounter = sheetCounter + 1                
    except HttpError as error:
        print(error)
    print()
    print(position - 1, "Total New Defects")          

def clearSpreadsheet():
    loadCreds()
    loadDefects()
    global position
    spreadsheetDefects_set = set(spreadsheetDefects)
    websiteDefects_set = set(websiteDefects)
    uniqueDefectsList = websiteDefects_set - spreadsheetDefects_set
    try:
        service = build("sheets", "v4", credentials=credentials)
        for item in uniqueDefectsList:
            if(item != "Smoke Detector System Needed"):
                print("Clearing:", str(item))
                sheetCounter = 0
                for row in websiteData:
                    if(str(item) == str(row[0])):
                        while True:
                            try:
                                #"NewDefects" is the name of the sheet being modified
                                update_requests = [
                                    {
                                        "range": f"NewDefects!B{position+1}",
                                        "values": [[f""]]
                                    },
                                    {
                                        "range": f"NewDefects!C{position+1}",
                                        "values": [[f""]]
                                    },
                                    {
                                        "range": f"NewDefects!D{position+1}",
                                        "values": [[f""]]
                                    },
                                    {
                                        "range": f"NewDefects!E{position+1}",
                                        "values": [[f""]]
                                    },
                                    {
                                        "range": f"NewDefects!F{position+1}",
                                        "values": [[f""]]
                                    },
                                    {
                                        "range": f"NewDefects!H{position+1}",
                                        "values": [[f""]]
                                    },
                                    {
                                        "range": f"NewProjects!B{position+1}",
                                        "values": [[f""]]
                                    },
                                    {
                                        "range": f"NewProjects!C{position+1}",
                                        "values": [[f""]]
                                    },
                                    {
                                        "range": f"NewProjects!D{position+1}",
                                        "values": [[f""]]
                                    },
                                    {
                                        "range": f"NewProjects!F{position+1}",
                                        "values": [[f""]]
                                    },
                                    {
                                        "range": f"NewProjects!G{position+1}",
                                        "values": [[f""]]
                                    },
                                    {
                                        "range": f"NewProjects!H{position+1}",
                                        "values": [[f""]]
                                    }
                                ]
                                batch_update_request = {
                                  "data": [
                                      {
                                          "range": req["range"],
                                          "values": req["values"],
                                          "majorDimension": "ROWS"
                                      }
                                      for req in update_requests
                                  ],
                                  "valueInputOption": "USER_ENTERED"
                                }

                                # Execute the batchUpdate request
                                response = service.spreadsheets().values().batchUpdate(spreadsheetId=SPREADSHEET_ID, body=batch_update_request).execute()
                                sheetCounter = 0
                                position = position + 1
                                break
                            except HttpError as error:
                              print("API call limit..")
                              time.sleep(30)
                        break
                    sheetCounter = sheetCounter + 1                
    except HttpError as error:
        print(error)
    print()
    print(position - 1, "Cleared Defects")  

while True:
    print("-------------------------------")
    print("1: HomeCloud URL \n2: Fill Spreadsheet \n3: Clear Spreadsheet \n4: View All Scanned Defects \n5: View Spreadsheet defects \n6: Exit")
    userInput = eval(input("Input: "))
    if(userInput == 1):
        while True:
            try:
                loadWebsite()
                break
            except ElementClickInterceptedException as e:
                print()
    elif(userInput == 2):
        fillSpreadsheet()
        loadProjects()
        position = 1
    elif(userInput == 3):
        clearSpreadsheet()
        position = 1
    elif(userInput == 4):
        print(websiteDefects)
        print("Size: ", str(len(websiteDefects)))
    elif(userInput == 5):
      print(spreadsheetDefects)
      print("size: ", str(len(spreadsheetDefects)))
    else:
        sys.exit()
