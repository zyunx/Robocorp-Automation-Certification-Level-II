from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.PDF import PDF
from RPA.Archive import Archive
import csv

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images
    """
    # browser.configure(
    #     slowmo=1000,
    # )

    orders = get_orders()
    for order in orders:
        success = False
        while (not success):
            open_robot_order_website()
            close_annoying_modal()
            fill_the_form(order)
            preview_the_order()
            submit_the_order()
            success = is_submit_successful()

        store_receipt_as_pdf(order["Order number"])
        screenshot_robot(order["Order number"])
        embed_screenshot_to_receipt(f"output/receipts/{order['Order number']}.png",
                                    f"output/receipts/{order['Order number']}.pdf")
    
    archive_receipts()


def get_orders():
    """Download excel file from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)
    
    orders = []
    with open('orders.csv', newline='') as csvfile:
        csvreader = csv.DictReader(csvfile)
        for row in csvreader:
            orders.append(row)
    return orders


def open_robot_order_website():
    browser.goto('https://robotsparebinindustries.com/')
    browser.goto('https://robotsparebinindustries.com/#/robot-order')


def close_annoying_modal():
    page = browser.page()
    page.click(".modal-dialog  button:first-child")

def fill_the_form(order):
    page = browser.page()
    page.select_option("#head", str(order["Head"]))
    page.locator("#id-body-" + order["Body"]).set_checked(True)
    page.locator("input[placeholder='Enter the part number for the legs']").fill(order['Legs'])
    page.locator("#address").fill(order["Address"])


def preview_the_order():
    page = browser.page()
    page.locator("#preview").click()


def submit_the_order():
    page = browser.page()
    page.locator("#order").click()

def is_submit_successful():
    page = browser.page()
    return page.is_visible("#receipt")

def store_receipt_as_pdf(order_number):
    page = browser.page()
    receipt_html = page.locator("#receipt").inner_html()

    pdf = PDF()
    pdf.html_to_pdf(receipt_html, f"output/receipts/{order_number}.pdf")


def screenshot_robot(order_number):
    page = browser.page()
    page.screenshot(path=f"output/receipts/{order_number}.png", full_page=True)


def embed_screenshot_to_receipt(screenshot, pdf_file):
    pdf = PDF()
    pdf.add_files_to_pdf(files=[pdf_file, screenshot], target_document=pdf_file)


def archive_receipts():
    lib = Archive()
    lib.archive_folder_with_zip('output/receipts',
                                'output/receipts.zip', include='*.pdf')
