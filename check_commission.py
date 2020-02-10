import pywinauto
import time
import win32com.client
from pywinauto import application
from tkinter import Tk

id_process = 6816
book = r'd:\Users\maksim.gvozdetskiy\Desktop\test_run.xlsx'
#file_input = r'd:\Users\maksim.gvozdetskiy\Desktop\check_applications.txt'
#file_output = r'd:\Users\maksim.gvozdetskiy\Desktop\application_commission.txt'
element_coordinates = {'application_large_table': (300, 152),
                       'payment': (604, 188),
                       'uah': (609, 683),
                       'commission': (741, 187)
                       } #for mouse

bad_result = 'no result'

fulfilled_applications = {}

class Pegasys():
    pass

def bad_application(P_application):
    if P_application in fulfilled_applications.keys():
        fulfilled_applications[P_application] += 1
    else:
        fulfilled_applications[P_application] = 1
        

def get_commission():
    pywinauto.mouse.click(button='left', coords=element_coordinates['commission'])
    commission = P.dialog[u'19'].texts()[0]
    P.dialog.close()
    P.winForm.Edit15.double_click()
    P.winForm.Edit15.type_keys('{BACKSPACE}')
    
    return commission

def get_uah():
    pywinauto.mouse.click(button='left', coords=element_coordinates['uah'])
    table = P.dialog[u'layoutControl12']
    table.type_keys('^c')
    
    return Tk().clipboard_get()

def payment():
    pywinauto.mouse.click(button='left', coords=element_coordinates['payment'])
    usd = P.dialog[u'30'].texts()[0]
    if usd == '': usd = P.dialog[u'31'].texts()[0]
    
    #get UAH
    
    uah = get_uah()
    if uah == 'UAH':
        old_x, old_y = element_coordinates['uah'][0], element_coordinates['uah'][1]
        element_coordinates['uah'] = (old_x + 147, old_y)
        uah = get_uah()
    
    return (usd, uah)

def pickUp_dialogue(P_application):
    text = u'\u0420\u0435\u0434\u0430\u043a\u0442\u0438\u0440\u043e\u0432\u0430\u043d\u0438\u0435 \u0437\u0430\u044f\u0432\u043a\u0438 #{}'.format(P_application)
    P.dialog = P.app[text]
    #P.dialog.wait()
    time.sleep(2)
    
    #coordinate adapting
    element_coordinates['payment'] = (P.dialog.Rectangle().left + 363, P.dialog.Rectangle().top + 54)
    element_coordinates['uah'] = (P.dialog.Rectangle().left + 355, P.dialog.Rectangle().top + 545)
    element_coordinates['commission'] = (P.dialog.Rectangle().left + 500, P.dialog.Rectangle().top + 54)
    
    if P.dialog[u'38'].texts()[0] != P_application:
        #this dialogue is not our application
        P.dialog.close()
        return 'Next'

def insert_application(P_application):
    '''Search by application. Opens an application for editing.
    '''
    P.app = application.Application().connect(process = id_process)
    P.winForm = P.app[u'WindowsForms10.Window.8.app.0.310f4af_r12_ad2']
    
    #check if any application is entered
    entry_field = P.winForm.Edit15.texts()[0] 
    if entry_field != '':
        P.winForm.Edit15.double_click()
        P.winForm.Edit15.type_keys('{BACKSPACE}')
        
    P.winForm.Edit15.type_keys(P_application)
    P.winForm.Edit15.type_keys('{ENTER}')
    time.sleep(2)
    
    #check if the application was found
    #for z in range(2):
        #pywinauto.mouse.click(button='left', coords=element_coordinates['application_large_table'])
        #large_table = P.winForm[u'WindowsForms10.Window.8.app.0.310f4af_r12_ad231']    
        #large_table.type_keys('^c')    
        #if Tk().clipboard_get() == P_application: #application found
            #break
        #elif z == 0:
            #time.sleep(2)
        #else:
            #return "Next"
        
    #open the editing dialog    
    P.winForm[u'\u0420\u0435\u0434\u0430\u043a\u0442\u0438\u0440\u043e\u0432\u0430\u0442\u044c'].click()

def start_check(sheet):
    '''Reads an excel file. Records the commission of the second crown. 
    If there is an error - leaves the second column empty
    '''
    P.launch = False
    
    #we get total columns and total rows
    total_columns = sheet.Cells(1,1).SpecialCells(11).Column
    total_rows = sheet.Cells(1,1).SpecialCells(11).Row
    
    #let's start sorting the plate
    for line_number in range(1, total_rows):
        line_number += 1 #the first line we do not need
        
        P_application = sheet.Cells(line_number,1).value.strip('№') #delete character
        P_application = P_application.strip(' ') #remove space
        if sheet.Cells(line_number,2).value != None or fulfilled_applications.get(P_application, -1) == 2:
            continue
                
        try:
            #insert a request
            #verification is not done, "Next" is not sent
            if insert_application(P_application) == "Next":
                P.launch = True
                bad_application(P_application)
            
            #pick up a dialogue    
            if pickUp_dialogue(P_application) == 'Next':
                P.launch = True
                bad_application(P_application)
                
            #get the amount of payment    
            usd, uah = payment()
            
            #get a commission
            commission = get_commission()
            
            sheet.Cells(line_number,2).value = str(((float(uah) / float(usd.replace(',', ''))) * float(commission)))
        except:
            P.launch = True
            bad_application(P_application)
                
if __name__ == '__main__':
    P = Pegasys() 
    P.launch = True
    Excel = win32com.client.Dispatch("Excel.Application")
    wb = Excel.Workbooks.Open(book)
    sheet = Excel.Sheets(1)
    
    while P.launch:
        start_check(sheet)
    
    wb.Save()
    wb.Close()
    Excel.Quit()    
    #Excel.Application.Quit()


#x, y = win32api.GetCursorPos()
#press_mouse(button='left', coords=(0, 0))