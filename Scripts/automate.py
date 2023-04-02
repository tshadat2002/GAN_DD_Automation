# This program is only customized for Penoles DD Workflows
# The excel sheets have columns for --> Dynamic rule #, Approver, Category, and questionairre conditions
# Just open up a stop and go place your mouse on the first condition of 
# the first rule you want to start with and run
# The program will fill in conditions and dynamic approvers as needed and make a sound when it finishes a S&G
# After that you must open the next S&G and repeat the same process
# The point of this program is to avoid manual work as much as possible
# in the future we hope to generalize this script into a class that can be used for all clients


import pyautogui
import pandas as pd
import numpy as np
import os
import winsound
import time

def first_rule(category, df, start = 1.0, low_risk = True):
    x = df[df['Critical Categories'] == category]
    x = x[x['Dynamic Rule #'] == float(start)]
    area = int(x['Area'].iloc[0].split()[0].strip('.'))
    company = int(x['Company'].iloc[0].split()[0].strip('.'))
    

    pyautogui.PAUSE = 0.5  
    #first condition
    pyautogui.press('tab') 
    pyautogui.press('2')
    pyautogui.press('tab') 
    pyautogui.press('e')
    pyautogui.press('tab')

    #fill area
    for i in range(area):
        pyautogui.press('down')
    
    #choose area
    pyautogui.PAUSE = 0.8
    pyautogui.press('enter')
    
    #go to second condition
    pyautogui.PAUSE = 0.5
    pyautogui.press('tab') 
    pyautogui.press('tab') 
    pyautogui.press('tab')

    #fill second condition 
    pyautogui.press('c')
    pyautogui.press('tab') 
    pyautogui.press('2')

    # choose correct condition
    pyautogui.PAUSE = 0.8
    pyautogui.press('enter')
    pyautogui.press('down') 
    pyautogui.press('enter')

    # fill rest of condition
    pyautogui.PAUSE = 0.5
    pyautogui.press('tab') 
    pyautogui.press('e')
    pyautogui.press('tab')
    # fill company 
    for i in range(company):
        pyautogui.press('down')
    
    # choose company
    pyautogui.PAUSE = 0.8
    pyautogui.press('enter') 

    # go to approver
    pyautogui.PAUSE = 0.5
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    
    #fill approver
    fill_approver(df = x, low_risk = low_risk)

    # go to next rule
    pyautogui.PAUSE = 0.5
    for i in range(6):
        pyautogui.press('tab') 


def fill_conditions(category, df, num_rules, start = 2, low_risk = True):
    # read in excel file with area and company
    t = df[df['Critical Categories'] == category]
    for i in range(start, num_rules):
        tx = t[t['Dynamic Rule #'] == float(i)].copy()
        area = int(tx['Area'].iloc[0].split()[0].strip('.'))
        company = int(tx['Company'].iloc[0].split()[0].strip('.'))
        
        # fill first condition 
        pyautogui.PAUSE = 0.5
        pyautogui.press('c')
        pyautogui.press('tab')
        pyautogui.press('2')
        pyautogui.press('tab')
        pyautogui.press('e')
        pyautogui.press('tab')
        # fill area
        for i in range(area):
            pyautogui.press('down')
        
        #choose area
        pyautogui.PAUSE = 0.8
        pyautogui.press('enter')

        #go to second condition
        pyautogui.PAUSE = 0.5
        pyautogui.press('tab') 
        pyautogui.press('tab') 
        pyautogui.press('tab') 

        # fill second condition
        pyautogui.press('c')
        pyautogui.press('tab') 
        pyautogui.press('2')

        # choose correct condition
        pyautogui.PAUSE = 0.8
        pyautogui.press('enter')
        pyautogui.press('down') 
        pyautogui.press('enter')
        
        # fill rest of condition
        pyautogui.PAUSE = 0.5
        pyautogui.press('tab') 
        pyautogui.press('e')
        pyautogui.press('tab')
        # fill company 
        for i in range(company):
            pyautogui.press('down')

        # choose company
        pyautogui.PAUSE = 0.8
        pyautogui.press('enter') 
        
        # go to approver
        pyautogui.PAUSE = 0.5
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        
        #fill approver
        fill_approver(df = tx, low_risk = low_risk)

        # go to next rule
        pyautogui.PAUSE = 0.5
        for i in range(6):
            pyautogui.press('tab') 


def start():
    pyautogui.PAUSE = 0.8
    #start it off
    x, y = pyautogui.position()
    pyautogui.click(x=x, y=y, clicks=1, interval=2, button='left')
    pyautogui.press('enter')
    pyautogui.click(x = x + 43, y = y - 25, clicks=1, interval=1, button='left')
    pyautogui.press('tab') 


def fill_approver(df, low_risk = True):
    pyautogui.PAUSE = 0.6
    if low_risk:
        approver = df['Email  (Business Owner Manager)'].iloc[0]
    else:
        approver = df['Email  (Business Owner Manager + 1)'].iloc[0]

    # enter into approver box
    pyautogui.press('enter') 

    type_to_search = pyautogui.locateCenterOnScreen('type_to_search.png', confidence=0.8)

    
    #find the area to search for approver
    pyautogui.click(type_to_search)
    #put in approver
    pyautogui.write(approver, interval=0.1)

    #wait 3 seconds for approver to load
    time.sleep(3)

    checkbox = pyautogui.locateCenterOnScreen('checkbox.png', confidence=0.9)
    #click approver
    pyautogui.click(checkbox)
    #click out
    pyautogui.click(checkbox[0] - 30, checkbox[1])



os.chdir('C:\\Users\\David Kull\\Documents\\DD_Automation\\Scripts')
data = pd.read_excel('Dynamic Approvers_Peñoles_30122022.xlsx', 
                    sheet_name = 'Critical WF approvers - Penoles', 
                    skiprows = 1)


# # start
start()
first_rule(category = 'Category 4: Provision of technical professional services',
            df = data.copy(),
            start = 6,
            low_risk=True)
fill_conditions(category = 'Category 4: Provision of technical professional services',
                    df = data.copy(),
                    num_rules=24,
                    start = 7,
                    low_risk=True)
# notify when done
winsound.Beep(500, 1000)













