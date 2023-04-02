#to start this program make sure the workflow
# screen is half the screen and that the 
# stop and Go Approval box touches the grey box
# This Script Only Fills Conditions
# use on 50% zoomed screen on chrome
# only customized for Penoles DD at the moment

import pyautogui
import pandas as pd
import numpy as np
import os
import winsound
# print(pyautogui.size())
# print(pyautogui.position())


os.chdir('C:\\Users\\David Kull\\Desktop')
data = pd.read_excel('Dynamic Approvers_Peñoles_30122022.xlsx', 
                    sheet_name = 'Critical WF approvers - Penoles', 
                    skiprows = 1)

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
    pyautogui.PAUSE = 1.5
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
    pyautogui.write(approver)

    #add sleep function here

    checkbox = pyautogui.locateCenterOnScreen('checkbox.png', confidence=0.9)
    #click approver
    pyautogui.click(checkbox)
    #click out
    pyautogui.click(checkbox[0] - 30, checkbox[1])




# # start
start()
first_rule(category = 'Category 4: Provision of technical professional services',
            df = data.copy(),
            start = 3,
            low_risk=True)
# fill_conditions(category = 'Category 4: Provision of technical professional services',
#                     df = data.copy(),
#                     num_rules=24,
#                     start = 2,
#                     low_risk=True)
# # notify when done
winsound.Beep(500, 1000)











