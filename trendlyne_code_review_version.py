import win32com.client
import pythoncom
import shutil
import psutil
import threading

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage

import sys
import time
import glob
import os

from datetime import datetime


def update_excel_links(excel_file):
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
    except Exception as e1:
        print("No existing Excel Window: ", e1)
        excel = win32com.client.Dispatch("Excel.Application")

    excel.Visible = True

    wb = excel.Workbooks.Open(excel_file, UpdateLinks=3)

    excel.DisplayAlerts = False
    wb.Save()
    wb.Close()
    excel.DisplayAlerts = True
    excel.Quit()

    del wb
    del excel


def run_vba_module(excel_file, module_name, macro_list, vba_code, refresh_link_stat):
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
    except Exception as e1:
        print("No existing Excel Window: ", e1)
        excel = win32com.client.Dispatch("Excel.Application")

    excel.Visible = True

    # Open the workbook
    # refresh_link_stat = 0
    wb = excel.Workbooks.Open(excel_file, UpdateLinks=refresh_link_stat)

    # Check if the module already exists and get the module object
    existing_module = None
    for vb_component in wb.VBProject.VBComponents:
        if vb_component.Name == module_name:
            existing_module = vb_component
            break

    # If the module doesn't exist, add it. If it exists, clear the existing code.
    if existing_module is None:
        existing_module = wb.VBProject.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
        existing_module.Name = module_name
    else:
        existing_module.CodeModule.DeleteLines(1, existing_module.CodeModule.CountOfLines)

    # Add or overwrite the VBA code in the module
    existing_module.CodeModule.AddFromString(vba_code)

    # Save the workbook
    excel.DisplayAlerts = False
    wb.Save()
    excel.DisplayAlerts = True

    # Run the macro
    for macro in macro_list:
        excel.Run(f'{module_name}.{macro}')

    # Close the workbook and quit Excel
    excel.DisplayAlerts = False
    wb.Save()
    wb.Close()
    excel.DisplayAlerts = True
    excel.Quit()

    del wb
    del excel


def run_vba_module_no_save(excel_file, module_name, macro_list, vba_code, refresh_link_stat):
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
    except Exception as e1:
        print("No existing Excel Window: ", e1)
        excel = win32com.client.Dispatch("Excel.Application")

    excel.Visible = True

    # Open the workbook
    # refresh_link_stat = 0
    wb = excel.Workbooks.Open(excel_file, UpdateLinks=refresh_link_stat)

    # Check if the module already exists and get the module object
    existing_module = None
    for vb_component in wb.VBProject.VBComponents:
        if vb_component.Name == module_name:
            existing_module = vb_component
            break

    # If the module doesn't exist, add it. If it exists, clear the existing code.
    if existing_module is None:
        existing_module = wb.VBProject.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
        existing_module.Name = module_name
    else:
        existing_module.CodeModule.DeleteLines(1, existing_module.CodeModule.CountOfLines)

    # Add or overwrite the VBA code in the module
    existing_module.CodeModule.AddFromString(vba_code)

    # Run the macro
    for macro in macro_list:
        excel.Run(f'{module_name}.{macro}')

    # Close the workbook and quit Excel
    excel.DisplayAlerts = False
    wb.Close(SaveChanges=False)
    excel.DisplayAlerts = True
    excel.Quit()

    del wb
    del excel


def click_para_limit(option_text, wait, driver):
    para_limit = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="paramlimit"]')))
    para_limit_select = para_limit.find_element(By.XPATH, './/*[@role="combobox"]')
    para_limit_select.click()

    result_list = driver.find_element(By.XPATH, '//ul[@class="select2-results__options"]')
    option_one = result_list.find_element(By.XPATH, f"//span[contains(text(), '{option_text}')]/ancestor::li[1]")
    option_one.click()


def sending_image_email(image_tags, sections, sender_info, recipients, change_prescripts):
    # Email settings
    smtp_server = 'smtp-mail.outlook.com'
    smtp_port = 587  # or 465 for SSL
    smtp_username = sender_info[0]
    smtp_password = sender_info[1]
    sender_email = sender_info[0]
    recipients = recipients

    # Create message container
    msg = MIMEMultipart('related')
    msg['Subject'] = f"{change_prescripts}MFW - " + datetime.now().strftime("%d/%m/%Y - %I:%M %p")
    msg['From'] = sender_email
    msg['To'] = ', '.join(recipients)

    # Create the body of the message (HTML version)
    html_body = """\
    <html>
      <head></head>
      <body>
    """

    current_datetime = "MFW - " + datetime.now().strftime("%d/%m/%Y - %I:%M %p")
    html_body += "<p>" + current_datetime + "</p>"

    # Embedding images in the email
    for i, (tag, section) in enumerate(zip(image_tags, sections)):
        image_cid = f"image{i}"  # Create a Content-ID for each image
        html_body += f"<h2><u>{section}</u></h2><br><img src='cid:{image_cid}'><br><br><br><hr/>"

    html_body += "</body></html>"
    part1 = MIMEText(html_body, 'html')
    msg.attach(part1)

    # Attach images to the email
    for i, tag in enumerate(image_tags):
        file_path = os.path.join(os.environ['TEMP'], f"{tag}.jpg")
        with open(file_path, 'rb') as img_file:
            img = MIMEImage(img_file.read(), name=os.path.basename(file_path))
            img.add_header('Content-ID', f"<image{i}>")
            msg.attach(img)

    # Send the message via local SMTP server
    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(smtp_username, smtp_password)
        server.sendmail(sender_email, recipients, msg.as_string())

    print('Email sent successfully' + datetime.now().strftime("%d/%m/%Y - %I:%M %p"))


def download_trendlyne(download_dir):
    fp = webdriver.FirefoxOptions()
    fp.set_preference("browser.download.folderList", 2)
    fp.set_preference("browser.download.manager.showWhenStarting", False)
    fp.set_preference("browser.download.dir", download_dir)
    fp.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/forced-download")
    fp.add_argument("--incognito")
    # fp.add_argument("--headless")
    driver = webdriver.Firefox(options=fp)

    wait = WebDriverWait(driver, 10)
    driver.maximize_window()

    website_url = "https://trendlyne.com/tools/data-downloader/"
    driver.get(website_url)

    user_name = driver.find_element(By.ID, 'id_login')
    user_name.send_keys("Username")

    password = driver.find_element(By.ID, 'id_password')
    password.send_keys("Password")

    button = driver.find_element(By.XPATH, "//button[contains(text(), 'Login')]")
    button.click()

    stock_group = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="stockgrouplimit"]')))
    stock_group_select = stock_group.find_element(By.XPATH, './/*[@role="combobox"]')
    stock_group_select.click()

    result_list = driver.find_element(By.XPATH, '//ul[@class="select2-results__options"]')
    option_one = result_list.find_element(By.XPATH, "//span[contains(text(), 'Coverage')]/ancestor::li[1]")
    option_one.click()

    option_text_list = ['Price and Changes', 'Scores and growth', 'Moving averages', 'Volume and Delivery',
                        'Technicals', 'F&O parameters']

    for option_text in option_text_list:
        click_para_limit(option_text, wait, driver)

    download_button = driver.find_element(By.XPATH, '//button[@class="btn dd-btn"]')
    download_button.click()

    new_file_name = fr"{download_dir}\Trendlyne Data.xlsx"

    if os.path.exists(new_file_name):
        os.remove(new_file_name)

    time.sleep(5)

    for _ in range(40):
        list_of_files = glob.glob(fr"{download_dir}\*")  # * means all. If we need specific format then *.csv

        if list_of_files:
            latest_file = max(list_of_files, key=os.path.getctime)
            if 'matsya' in latest_file.lower() and 'coverage' in latest_file.lower():
                os.rename(latest_file, new_file_name)
                break
            else:
                time.sleep(1)
        else:
            print("No files found in the directory.")

    driver.quit()


def vba_code_write():
    vba_code = """
    Sub create_jpg_image(awb As Workbook, SheetName As String, xRgAddrss As String, nameFile As String)
        Dim xRgPic As Range
        Dim xShape As Shape
        awb.Activate
        awb.Worksheets(SheetName).Activate
        DoEvents
        Set xRgPic = awb.Worksheets(SheetName).Range(xRgAddrss)
        DoEvents
    1:
        On Error GoTo 1
        xRgPic.CopyPicture xlScreen, xlPicture
        DoEvents
        On Error GoTo 0
        With awb.Worksheets(SheetName).ChartObjects.Add(xRgPic.Left, xRgPic.Top, xRgPic.Width, xRgPic.Height)
            .Activate
            For Each xShape In ActiveSheet.Shapes
                xShape.Line.Visible = msoFalse
            Next
            .Chart.Paste
            DoEvents
            .Chart.Export Environ$("temp") & "\\" & nameFile & ".jpg", "JPG"
        End With
       awb.Worksheets(SheetName).ChartObjects(awb.Worksheets(SheetName).ChartObjects.Count).Delete
    Set xRgPic = Nothing
    End Sub

    Sub create_images()
        Dim twb As Workbook
        Set twb = ThisWorkbook

        delete_file Environ$("temp") & "\\" & "MFW1.jpg"
        delete_file Environ$("temp") & "\\" & "MFW2.jpg"
        delete_file Environ$("temp") & "\\" & "MFW3.jpg"
        delete_file Environ$("temp") & "\\" & "MFW4.jpg"
        delete_file Environ$("temp") & "\\" & "MFW5.jpg"
        delete_file Environ$("temp") & "\\" & "MFW6.jpg"
        delete_file Environ$("temp") & "\\" & "MFW7.jpg"
        delete_file Environ$("temp") & "\\" & "MFW8.jpg"

        Call create_jpg_image(twb, "Summary", "I6:L40", "MFW1")

        Call create_jpg_image(twb, "Momentum FW & Filters", "AH3:AQ30", "MFW2")

        Call create_jpg_image(twb, "Output-Positive", "Y4:AI40", "MFW3")

        Call create_jpg_image(twb, "Output-Negative", "Y84:AI103", "MFW4")

        Call create_jpg_image(twb, "Momentum FW & Filters-MT", "AI3:AR30", "MFW5")

        Call create_jpg_image(twb, "Output-Positive-MT", "Y4:AJ44", "MFW6")

        Call create_jpg_image(twb, "Output-Negative-MT", "Y80:AJ100", "MFW7")
        Call create_jpg_image(twb, "Position", "B2:R45", "MFW8")

    End Sub

    Sub sort_sheets()
        Dim ws1 As Worksheet, ws2 As Worksheet
        Dim lastRow1 As Long, lastRow2 As Long

        Set ws1 = ThisWorkbook.Worksheets("Momentum FW & Filters")
        Set ws2 = ThisWorkbook.Worksheets("Momentum FW & Filters-MT")

        ' Find the last row with data in Column L in ws1
        lastRow1 = ws1.Cells(ws1.Rows.Count, "L").End(xlUp).Row

        ' Sort Column L in ws1 in descending order
        With ws1.Sort
            .SortFields.Clear
            .SortFields.Add Key:=ws1.Range("L5:L" & lastRow1), Order:=xlDescending
            .SetRange ws1.Range("A5:AF" & lastRow1)
            .Header = xlYes
            .Apply
        End With

        ' Find the last row with data in Column D in ws2
        lastRow2 = ws2.Cells(ws2.Rows.Count, "D").End(xlUp).Row

        ' Sort Column D in ws2 in descending order
        With ws2.Sort
            .SortFields.Clear
            .SortFields.Add Key:=ws2.Range("D5:D" & lastRow2), Order:=xlDescending
            .SetRange ws2.Range("A5:AF" & lastRow2)
            .Header = xlYes
            .Apply
        End With
    End Sub

    Sub create_history()

        Dim twb As Workbook
        Dim ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet, ws_momentum As Worksheet, ws_adx As Worksheet
        Dim lastCol As Long, col_index As Variant
        Dim i As Long
        Dim startDate As Date
        Dim lastRow As Long

        Set twb = ThisWorkbook
        Set ws1 = twb.Sheets("Log")
        Set ws2 = twb.Sheets("Log-MT")
        Set ws3 = twb.Sheets("Momentum FW & Filters")
        Set ws_momentum = twb.Sheets("Momentum History")
        Set ws_adx = twb.Sheets("ADX History")

        ' For ws1
        ' Set the formula in cell AE1
        ws1.Range("AE1").Formula = "=MATCH(TODAY(), AA2:XFD2, 0) + 26"

        ' Calculate to update the formula result
        Application.Calculate

        col_index = ws1.Range("AE1").Value

        If IsError(col_index) Then
        ' Find the last used column in row 2
        lastCol = ws1.Cells(2, ws1.Columns.Count).End(xlToLeft).Column

        ' Add today's date in the next column to the right
        ws1.Cells(2, lastCol + 1).Value = Date
        ws1.Cells(2, lastCol + 1).NumberFormat = "mm/dd/yyyy" ' Format as date

        Application.Calculate
        ' Recheck col_index
        col_index = ws1.Range("AE1").Value
        End If

        If col_index < 27 Then
            MsgBox "col_index is less than 27, macro will stop."
            Exit Sub
        End If

        ' Find the last used column in row 2
        lastCol = ws1.Cells(2, ws1.Columns.Count).End(xlToLeft).Column

        lastRow = ws1.Cells(ws1.Rows.Count, "K").End(xlUp).Row ' Last row in column K
        ws1.Range("K3:K" & lastRow).Copy
        ws1.Cells(3, col_index).PasteSpecial xlValues
        ws1.Range(ws1.Cells(2, col_index - 2), ws1.Cells(lastRow, col_index)).Copy
        ws1.Range("N2").PasteSpecial xlValues

        ' For ws2
        ' Set the formula in cell AE1
        ws2.Range("AE1").Formula = "=MATCH(TODAY(), AA2:XFD2, 0) + 26"

        ' Calculate to update the formula result
        Application.Calculate

        col_index = ws2.Range("AE1").Value

        If IsError(col_index) Then
        ' Find the last used column in row 2
        lastCol = ws2.Cells(2, ws2.Columns.Count).End(xlToLeft).Column

        ' Add today's date in the next column to the right
        ws2.Cells(2, lastCol + 1).Value = Date
        ws2.Cells(2, lastCol + 1).NumberFormat = "mm/dd/yyyy" ' Format as date

        Application.Calculate
        ' Recheck col_index
        col_index = ws2.Range("AE1").Value
        End If

        If col_index < 27 Then
            MsgBox "col_index is less than 27, macro will stop."
            Exit Sub
        End If

        lastRow = ws2.Cells(ws2.Rows.Count, "M").End(xlUp).Row ' Last row in column M
        ws2.Range("M3:M" & lastRow).Copy
        ws2.Cells(3, col_index).PasteSpecial xlValues
        ws2.Range(ws2.Cells(2, col_index - 2), ws2.Cells(lastRow, col_index)).Copy
        ws2.Range("P2").PasteSpecial xlValues

        ' For ws3 copying to ws_momentum
        ' Set the formula in cell C1
        ws_momentum.Range("C1").Formula = "=MATCH(TODAY(), 2:2, 0)"

        ' Calculate to update the formula result
        Application.Calculate

        col_index = ws_momentum.Range("C1").Value

        If IsError(col_index) Then
        ' Find the last used column in row 2
        lastCol = ws_momentum.Cells(2, ws_momentum.Columns.Count).End(xlToLeft).Column

        ' Add today's date in the next column to the right
        ws_momentum.Cells(2, lastCol + 1).Value = Date
        ws_momentum.Cells(2, lastCol + 1).NumberFormat = "mm/dd/yyyy" ' Format as date

        Application.Calculate
        ' Recheck col_index
        col_index = ws_momentum.Range("C1").Value
        End If

        lastRow = ws3.Cells(ws3.Rows.Count, "L").End(xlUp).Row ' Last row in column L
        ws3.Range("L6:L" & lastRow).Copy
        ws_momentum.Cells(3, col_index).PasteSpecial xlValues

        ' Apply percentage format to the pasted range
        With ws_momentum
            .Range(.Cells(3, col_index), .Cells(3 + (lastRow - 6), col_index)).NumberFormat = "0%"
        End With

        ' For ws3 copying to ws_adx
        ' Set the formula in cell C1
        ws_adx.Range("C1").Formula = "=MATCH(TODAY(), 2:2, 0)"

        ' Calculate to update the formula result
        Application.Calculate

        col_index = ws_adx.Range("C1").Value

        If IsError(col_index) Then
        ' Find the last used column in row 2
        lastCol = ws_adx.Cells(2, ws_adx.Columns.Count).End(xlToLeft).Column

        ' Add today's date in the next column to the right
        ws_adx.Cells(2, lastCol + 1).Value = Date
        ws_adx.Cells(2, lastCol + 1).NumberFormat = "mm/dd/yyyy" ' Format as date

        Application.Calculate
        ' Recheck col_index
        col_index = ws_adx.Range("C1").Value
        End If

        lastRow = ws3.Cells(ws3.Rows.Count, "W").End(xlUp).Row ' Last row in column W
        ws3.Range("W6:W" & lastRow).Copy
        ws_adx.Cells(3, col_index).PasteSpecial xlValues
    End Sub

    Sub create_images_end_day()
        Dim twb As Workbook
        Set twb = ThisWorkbook

        delete_file Environ$("temp") & "\\" & "MFW1.jpg"
        delete_file Environ$("temp") & "\\" & "MFW2.jpg"
        delete_file Environ$("temp") & "\\" & "MFW3.jpg"
        delete_file Environ$("temp") & "\\" & "MFW4.jpg"
        delete_file Environ$("temp") & "\\" & "MFW5.jpg"
        delete_file Environ$("temp") & "\\" & "MFW6.jpg"
        delete_file Environ$("temp") & "\\" & "MFW7.jpg"
        delete_file Environ$("temp") & "\\" & "MFW8.jpg"
        delete_file Environ$("temp") & "\\" & "MFW9.jpg"
        delete_file Environ$("temp") & "\\" & "MFW10.jpg"

        Call create_jpg_image(twb, "Summary", "I6:L40", "MFW1")
        Call create_jpg_image(twb, "Momentum FW & Filters", "AH3:AQ30", "MFW2")
        Call create_jpg_image(twb, "Output-Positive", "Y5:AI40", "MFW3")
        Call create_jpg_image(twb, "Output-Negative", "Y84:AI103", "MFW4")
        Call create_jpg_image(twb, "Momentum FW & Filters-MT", "AI3:AR30", "MFW5")
        Call create_jpg_image(twb, "Output-Positive-MT", "Y4:AJ44", "MFW6")
        Call create_jpg_image(twb, "Output-Negative-MT", "Y80:AJ100", "MFW7")
        Call create_jpg_image(twb, "Position", "B2:R45", "MFW8")
        Call createJpg(twb, "Log", "S4:V20", "MFW9")
        Call createJpg(twb, "log-MT", "U4:X20", "MFW10")

    End Sub

    Sub copy_paste_AA_value()
        Dim ws1 As Worksheet, ws2 As Worksheet
        Dim lastRow1 As Long, lastRow2 As Long

        Set ws1 = ThisWorkbook.Worksheets("Momentum FW & Filters")
        Set ws2 = ThisWorkbook.Worksheets("Momentum FW & Filters-MT")

        ' Copy from ws1 and paste in ws1
        lastRow1 = ws1.Cells(ws1.Rows.Count, "AA").End(xlUp).Row ' Last row in column AA
        ws1.Range("AA6:AB" & lastRow1).Copy
        ws1.Range("X6").PasteSpecial xlPasteValues
        Application.CutCopyMode = False

        ' Copy from ws2 and paste in ws2
        lastRow2 = ws2.Cells(ws2.Rows.Count, "AA").End(xlUp).Row ' Corrected to ws2
        ws2.Range("AA6:AB" & lastRow2).Copy
        ws2.Range("X6").PasteSpecial xlPasteValues
        Application.CutCopyMode = False

    End Sub

    Sub check_and_write()
        Dim ws1 As Worksheet, ws2 As Worksheet
        Dim allErrors As Boolean
        Dim filePath As String
        Dim fileNumber As Integer

        Set ws1 = ThisWorkbook.Worksheets("Momentum FW & Filters")
        Set ws2 = ThisWorkbook.Worksheets("Momentum FW & Filters-MT")

        ' Assume initially that all cells have errors
        allErrors = True

        ' Check cells in ws1
        If Not IsError(ws1.Range("AH6")) Then allErrors = False
        If Not IsError(ws1.Range("AI6")) Then allErrors = False
        If Not IsError(ws1.Range("AK6")) Then allErrors = False
        If Not IsError(ws1.Range("AL6")) Then allErrors = False

        ' Check cells in ws2
        If Not IsError(ws2.Range("AI6")) Then allErrors = False
        If Not IsError(ws2.Range("AJ6")) Then allErrors = False
        If Not IsError(ws2.Range("AL6")) Then allErrors = False
        If Not IsError(ws2.Range("AM6")) Then allErrors = False

        ' Define the file path for the text file (same directory as the Excel file)
        filePath = ThisWorkbook.Path & "\send_emal_trigger.txt"

        fileNumber = FreeFile()

        ' Open the file for writing
        Open filePath For Output As #fileNumber

        ' Write to the file based on the condition
        If allErrors Then
            Print #fileNumber, "not_send"
        Else
            Print #fileNumber, "do_send"
        End If

        ' Close the file
        Close #fileNumber

    End Sub

        """
    return vba_code


def check_weekend():
    today = datetime.now()

    # Get the day of the week as an integer (Monday=0, Sunday=6)
    day_of_week = today.weekday()

    # Check if it's Saturday or Sunday
    if day_of_week == 5:
        print("Today is Saturday. The program will not run")
        time.sleep(20)
        sys.exit(0)
    elif day_of_week == 6:
        print("Today is Sunday. The program will not run")
        time.sleep(20)
        sys.exit(0)


def check_delete_work_file(work_file):
    if os.path.exists(work_file):
        os.remove(work_file)


def create_work_file(original_file, work_file):
    try:
        shutil.copy2(original_file, work_file)
        print(f"Work file created successfully: {work_file}")
    except Exception as e:
        print(f"Error creating backup: {e}")


def replace_original_file(original_file, work_file):
    try:
        # Delete original file
        os.remove(original_file)
        print(f"Original file deleted: {original_file}")

        # Rename backup file to original file name
        os.rename(work_file, original_file)
        print(f"Work file file renamed to original file name: {original_file}")
    except Exception as e:
        print(f"Error creating backup: {e}")


def close_excel_instance():
    for process in psutil.process_iter(['pid', 'name', 'status']):
        process_info = process.as_dict(attrs=['pid', 'name', 'status'])

        # Check if the process is Excel
        if "excel" in process_info['name'].lower():
            try:
                # Terminate the process
                process.terminate()
                print(f"Terminated Excel instance with PID: {process.pid}")
            except Exception as e:
                print(f"Error while terminating Excel instance: {e}")


def monitor_function(timeout, excel_thread):
    start_time = time.time()
    while time.time() - start_time < timeout:
        if not excel_thread.is_alive():
            print("Excel thread has completed or stopped.")
            return
        time.sleep(1)

    # If timeout is reached, try to terminate Excel
    print("Function timed out, terminating Excel.")
    try:
        close_excel_instance()
    except psutil.NoSuchProcess:
        pass


def check_excel_validity(excel_work_file, vba_code):
    print("Testing validity of work file")
    time.sleep(5)
    module_name = "trendlyne_email"
    macro_list = ["sort_sheets"]

    run_vba_module_no_save(excel_work_file, module_name, macro_list, vba_code, 0)


def write_email_trigger(excel_work_file, vba_code):
    time.sleep(2)
    module_name = "trendlyne_email"
    macro_list = ["check_and_write"]

    run_vba_module_no_save(excel_work_file, module_name, macro_list, vba_code, 0)


def send_normal_email(download_dir, change_prescripts):
    image_tags = ["MFW1", "MFW2", "MFW3", "MFW4", "MFW5", "MFW6", "MFW7", "MFW8"]
    sections = [
        "Common Tickers", "Short Term Setup", "Positive Momentum Tickers", "Negative Momentum Tickers",
        "Medium Term Setup", "MT Positive Momentum Tickers", "MT Negative Momentum Tickers",
        "Positions"
    ]

    sender_info_file_path = fr"{download_dir}\sender_info.txt"
    recipients_file_path = fr"{download_dir}\recipients.txt"

    with open(sender_info_file_path, 'r') as file:
        sender_info = [line.rstrip('\n') for line in file]

    with open(recipients_file_path, 'r') as file:
        recipients = [line.rstrip('\n') for line in file]

    print("Start sending normal email")
    sending_image_email(image_tags, sections, sender_info, recipients, change_prescripts)


# Function to check and only send the email when Output tables changed
def check_and_send_email(download_dir):
    trigger_file = fr"{download_dir}\send_emal_trigger.txt"

    try:
        with open(trigger_file, 'r') as file:
            content = file.read().strip()

        if content == "do_send":
            send_normal_email(download_dir, '')
        elif content == "not_send":
            # send_normal_email(download_dir, 'No Change to Output Table - ')
            print('No Change to Output Table. Wont send email.')
            time.sleep(20)
            return
        else:
            print("Error: File content is neither 'do_send' nor 'not_send'.")
            return

    except FileNotFoundError:
        print(f"Error: The file {trigger_file} does not exist.")
    except Exception as e1:
        print(f"An error occurred: {e1}")


def excel_function(download_dir, return_dict):
    for _ in range(3):
        try:
            pythoncom.CoInitialize()

            # Section 1
            excel_file = fr"{download_dir}\Momentum Investing FW_with adx.xlsm"
            excel_work_file = fr"{download_dir}\Momentum Investing FW_with adx_work_file.xlsm"

            vba_code = vba_code_write()

            check_delete_work_file(excel_work_file)
            create_work_file(excel_file, excel_work_file)

            module_name = "trendlyne_email"
            macro_list = ["copy_paste_AA_value"]

            run_vba_module(excel_work_file, module_name, macro_list, vba_code, 0)

            time.sleep(3)

            # Section 2
            excel_data_source_file = fr"{download_dir}\MFW Data Source.xlsx"
            update_excel_links(excel_data_source_file)

            time.sleep(3)

            # Section 3
            excel_file = fr"{download_dir}\Momentum Investing FW_with adx.xlsm"
            excel_work_file = fr"{download_dir}\Momentum Investing FW_with adx_work_file.xlsm"

            module_name = "trendlyne_email"
            macro_list = ["sort_sheets", "create_images"]

            run_vba_module(excel_work_file, module_name, macro_list, vba_code, 3)

            # Write email trigger
            write_email_trigger(excel_work_file, vba_code)

            # Test if Excel file still works well
            check_excel_validity(excel_work_file, vba_code)

            # Send email
            check_and_send_email(download_dir)

            return_dict['pass_trigger'] = 1

            # Replace original Excel
            replace_original_file(excel_file, excel_work_file)

            pythoncom.CoUninitialize()

            time.sleep(2)
            break

        except Exception as e1:
            print("Script error", e1)
            time.sleep(10)

            # Close all frozen Excel instances to ensure fresh start
            close_excel_instance()

            time.sleep(2)

            excel_work_file = fr"{download_dir}\Momentum Investing FW_with adx_work_file.xlsm"
            check_delete_work_file(excel_work_file)

            time.sleep(2)


def main():
    print("Starting")

    def myexcepthook(type1, value, traceback, oldhook=sys.excepthook):
        oldhook(type1, value, traceback)
        input("Press Enter... ")

    sys.excepthook = myexcepthook

    if hasattr(sys, '_MEIPASS'):
        script_directory = os.path.dirname(os.path.abspath(sys.argv[0]))
    else:
        script_directory = os.path.dirname(os.path.abspath(__file__))

    # Make sure the program won't run on weekend
    check_weekend()

    download_dir = script_directory

    download_trendlyne(download_dir)

    return_dict = {'pass_trigger': 0}

    for _ in range(3):
        excel_thread = threading.Thread(target=excel_function, args=(download_dir, return_dict))
        excel_thread.start()

        monitor_thread = threading.Thread(target=monitor_function, args=(240, excel_thread))  # 240 seconds timeout
        monitor_thread.start()

        excel_thread.join()
        monitor_thread.join()

        time.sleep(0.5)

        # Use threading to catch cases where Excel not responding or froze
        if return_dict['pass_trigger'] == 1:
            break


if __name__ == "__main__":
    main()
