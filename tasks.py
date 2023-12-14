from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive
from RPA.FileSystem import FileSystem


@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    Open_the_robot_order_website()
    download_excel_file()
    CSV_to_Table()
    Create_ZIP_File()

def Open_the_robot_order_website():
    """Navigates to the given URL"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")
    page = browser.page()
    if page.is_visible("button:text('Yep')"):
        page.click("button:text('Yep')")


def download_excel_file():
    """Downloads excel file from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)


def Fill_the_form(fillorder):
    """Fills in the sales data and click the 'Submit' button"""
    page = browser.page()
    page.select_option("#head", str(fillorder["Head"]))
    page.set_checked("#id-body-" + fillorder["Body"],True)
    page.fill("input[placeholder='Enter the part number for the legs']", fillorder["Legs"]) 
    page.fill("#address", fillorder["Address"]) 


def Preview_the_order():
    page = browser.page()
    for x in range(10):
        blReceipt =   page.is_visible("#receipt")
        if blReceipt == False:
            if page.is_visible("button:text('Preview')"):
                page.click("button:text('Order')")
        else:
            break

    

def Store_the_receipt_as_a_PDF_file(orderno):
    page = browser.page()
    page.wait_for_selector("#receipt")
    orderhtml = page.locator("#receipt").inner_html()
    pdf = PDF()
    path = "output/Order_no" + str(orderno) + ".pdf"
    pdf.html_to_pdf(orderhtml, path)
    return path


def Take_a_screenshot_of_the_robot(orderno):
    """Take a screenshot of the page"""
    page = browser.page()
    sspath="output/Order_no" + str(orderno) + ".png"
    page.screenshot(path=sspath)
    return sspath

def Embed_the_robot_screenshot_to_the_receipt_PDF_file(screenshot,pdf):
    add_pdf = PDF()
    list_of_files = [
        screenshot
    ]
    add_pdf.add_files_to_pdf(files=list_of_files,
                             target_document=pdf,
                             append=True)

def Create_ZIP_File():
    lib = Archive()
    lib.archive_folder_with_zip('output/', 'PDF.zip')





def CSV_to_Table():
    page = browser.page()
    tables = Tables()
    table = tables.read_table_from_csv("orders.csv")
    for ordr in table:
        print(ordr["Address"])
        Fill_the_form(ordr)
        Preview_the_order()
        pdf = Store_the_receipt_as_a_PDF_file(ordr["Order number"])
        ss = Take_a_screenshot_of_the_robot(ordr["Order number"])
        Embed_the_robot_screenshot_to_the_receipt_PDF_file(ss,pdf)
        if page.is_visible("button:text('Order another robot')"):
            page.click("button:text('Order another robot')")
        if page.is_visible("button:text('Yep')"):
            page.click("button:text('Yep')")    
