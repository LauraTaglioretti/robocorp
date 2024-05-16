from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.Browser.Selenium import Selenium
from RPA.PDF import PDF
from RPA.Archive import Archive
import time

seleniumBrowser = Selenium()

@task
def order_robots_from_RobotSpareBin():
    """    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images."""
      
    open_robot_order_website()
    orders = get_orders()
    order_robots(orders)
    archive_receipts()



def open_robot_order_website():
    """Navigates to the robot order website"""
    print("Navigating to robot order website")
    seleniumBrowser.open_available_browser("https://robotsparebinindustries.com/#/robot-order")
    seleniumBrowser.maximize_browser_window()


def get_orders():
    """Downloads the orders file csv file
    Reads the file and saves it in a table"""
    print("Obtaining orders to be made")
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True, 
    target_file="output")

    library = Tables()
    orders = library.read_table_from_csv("output/orders.csv")
    
    return orders 


def order_robots(orders):
    """Loops through the orders table"""
    for row in orders:
        fill_in_robot_order_form(row)


def fill_in_robot_order_form(orderDetails):
    """Fills in the robot order form"""
    print("Ordering robot" + str(orderDetails["Order number"]))
    time.sleep(0.5)
    seleniumBrowser.click_button("//button[@class='btn btn-dark']")
    seleniumBrowser.select_from_list_by_index("id:head", orderDetails["Head"])
    seleniumBrowser.select_radio_button("body", orderDetails["Body"])
    seleniumBrowser.input_text("//input[@placeholder='Enter the part number for the legs']", 
    orderDetails["Legs"])
    seleniumBrowser.input_text("address", orderDetails["Address"])
    time.sleep(0.5)
    seleniumBrowser.click_element_when_clickable("id:preview")
    time.sleep(0.5)
    
    screenshot_robot(orderDetails["Order number"])
    
    seleniumBrowser.click_button("id:order")
    time.sleep(0.5)
    
    while seleniumBrowser.does_page_contain_element("id:order"):
        print("Web error. Retrying")
        time.sleep(1)
        seleniumBrowser.click_element_when_clickable("id:order")
    
    store_receipt_as_pdf(orderDetails["Order number"])
    
    time.sleep(0.5)
    seleniumBrowser.click_button("id:order-another")



def screenshot_robot(orderNumber):
    """Takes a screenshot of the robot preview"""
    seleniumBrowser.capture_element_screenshot("id:robot-preview-image", 
    "output/robot screenshots/robot_" + orderNumber + ".png")



def store_receipt_as_pdf(orderNumber):
    """Saves the order as PDF"""
    print("Generating pdf file")
    receipt = seleniumBrowser.get_element_attribute("id:receipt", "innerHTML")
    pdf = PDF()
    pdf.html_to_pdf(receipt, "output/pdf files/robot_" + orderNumber + ".pdf")
    screenshot = ["output/robot screenshots/robot_" + orderNumber + ".png"]
    pdf.add_files_to_pdf(screenshot, "output/pdf files/robot_" + orderNumber + ".pdf", append = True)


def archive_receipts():
    """Generates a zip file with all the pdf files"""
    print("Generating zip file")
    zip = Archive()
    zip.archive_folder_with_zip("./output/pdf files", "./output/receipts.zip")