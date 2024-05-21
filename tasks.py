from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.Tables import Tables
from RPA.PDF import PDF
import pyautogui

@task
def order_robots_from_RobotSpareBin():
    download_csv_file()
    open_robot_order_website()
    complete_all_orders()

def open_robot_order_website():
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def input_order(robot_order):
    page = browser.page()
    page.click("text=OK")
    #input head
    page.click("id=head")
    if str(robot_order["Head"]) == "1": 
        pyautogui.typewrite(['down'])
    elif str(robot_order["Head"]) == "2": 
        pyautogui.typewrite(['down','down'])
    elif str(robot_order["Head"]) == "3": 
        pyautogui.typewrite(['down','down','down'])
    elif str(robot_order["Head"]) == "4": 
        pyautogui.typewrite(['down','down','down','down'])
    elif str(robot_order["Head"]) == "5": 
        pyautogui.typewrite(['down','down','down','down','down'])
    elif str(robot_order["Head"]) == "6": 
        pyautogui.typewrite(['down','down','down','down','down','down'])
    pyautogui.typewrite(["enter"])
    #input body
    if robot_order["Body"] == "1": page.click("text=Roll-a-thor body")
    elif str(robot_order["Body"]) == "2": 
        page.click("text=Peanut crusher body")
    elif str(robot_order["Body"]) == "3": 
        page.click("text=D.A.V.E body")
    elif str(robot_order["Body"]) == "4": 
        page.click("text=Andy Roid body")
    elif str(robot_order["Body"]) == "5": 
        page.click("text=Spanner mate body")
    elif str(robot_order["Body"]) == "6": 
        page.click("text=Drillbit 2000 body")
    #input legs
    if str(robot_order["Legs"]) == "1": 
        page.locator("xpath=//label[contains(.,'3. Legs:')]/../input").click()
        pyautogui.typewrite(['up'])
    elif str(robot_order["Legs"]) == "2": 
        page.locator("xpath=//label[contains(.,'3. Legs:')]/../input").click()
        pyautogui.typewrite(['up','up'])
    elif str(robot_order["Legs"]) == "3": 
        page.locator("xpath=//label[contains(.,'3. Legs:')]/../input").click()
        pyautogui.typewrite(['up','up','up'])
    elif str(robot_order["Legs"]) == "4": 
        page.locator("xpath=//label[contains(.,'3. Legs:')]/../input").click()
        pyautogui.typewrite(['up','up','up','up'])
    elif str(robot_order["Legs"]) == "5": 
        page.locator("xpath=//label[contains(.,'3. Legs:')]/../input").click()
        pyautogui.typewrite(['up','up','up','up','up'])
    elif str(robot_order["Legs"]) == "6": 
        page.locator("xpath=//label[contains(.,'3. Legs:')]/../input").click()
        pyautogui.typewrite(['up','up','up','up','up','up'])
    #input address
    page.locator("#address").fill(robot_order["Address"])
    page.click("text=Preview")
    screenshot_robot(robot_order)
    while page.locator("#order").is_visible()==True:
        page.locator("#order").click()
    save_as_PDF(robot_order)
    merge_PDF(robot_order)
    page.click("text=ORDER ANOTHER ROBOT")

def screenshot_robot(robot_order):
    page = browser.page()
    page.locator(".col-sm-5").screenshot(path="Order Number "+robot_order["Order number"]+".png")

def save_as_PDF(robot_order):
    page = browser.page()
    receipt_html = page.locator("#receipt").inner_html()
    pdf = PDF()
    pdf.html_to_pdf(receipt_html, "Order Number "+robot_order["Order number"]+".pdf")

def merge_PDF(robot_order):
    pdf=PDF()
    files_to_merge = ["Order Number "+robot_order["Order number"]+".pdf", "Order Number "+robot_order["Order number"]+".png"]
    pdf.add_files_to_pdf(files=files_to_merge, target_document = "Order number "+robot_order["Order number"]+".pdf")

def download_csv_file():
    """Downloads csv file from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

def complete_all_orders():
    table = Tables()
    orders = table.read_table_from_csv("orders.csv", columns=["Order number", "Head", "Body", "Legs", "Address"])
    for row in orders:
        input_order(row)

    