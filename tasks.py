from robocorp.tasks import task
from robocorp import browser
from pathlib import Path

from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive

OUTPUT = Path("output")
RECEIPTS = OUTPUT / "receipts"
IMAGES = OUTPUT / "images"
ZIP_TARGET = OUTPUT / "RobotSpareBin_Receipts.zip"


@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    browser.configure(
        slowmo=500,
    )

    open_robot_order_website()
    download_orders_csv()

    orders = get_orders()

    for order in orders:
        close_annoying_modal()
        process_single_order(order)
    
    archive_receipts()

    print("TASK COMPLETATO")
    
def open_robot_order_website():
    """Open robot order website"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def download_orders_csv():
    """Downloads csv file from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

def get_orders():
    """Takes orders from csv"""
    tables = Tables()
    orders = tables.read_table_from_csv("orders.csv", header=True)
    return orders

def close_annoying_modal():
    """Closes pop-up"""
    page = browser.page()
    if page.is_visible("text=OK"):
        page.click("text=OK")

def process_single_order(order: dict):
    """Complete pipeline for a single order"""
    order_no = order["Order number"]

    fill_the_form(order)
    img_path = screenshot_robot(order_no)
    submit_until_success()
    store_receipt_as_pdf(order_no, img_path)
    reset_form_for_next_order()

    print(f"Order {order_no} completed")

def fill_the_form(order: dict):
    """Fills the form with order data"""
    page = browser.page()

    page.select_option("#head", str(order["Head"]))
    page.click(f"#id-body-{order['Body']}")

    page.fill("input[placeholder*='part number for the legs']", str(order["Legs"]))

    page.fill("#address", order["Address"])

def screenshot_robot(order_no: str) -> Path:
    """Saves preview screenshot"""
    page = browser.page()
    page.click("text=Preview")
    page.wait_for_selector("#robot-preview-image")
    img_path = IMAGES / f"robot_{order_no}.png"
    # Screenshot robot image
    page.locator("#robot-preview-image").screenshot(path=str(img_path))
    return img_path

def submit_until_success(max_attempts: int = 5):
    """Submits until there's the receipt"""
    page = browser.page()

    for attempt in range(1, max_attempts + 1):
        page.click("#order")
        try:
            page.wait_for_selector("#receipt", timeout=3000)
            return
        except Exception:
            if attempt == max_attempts:
                raise RuntimeError("Impossible complete the order, too many tries")
            print(f"Submit failed: ({attempt}/{max_attempts})")

def store_receipt_as_pdf(order_no: str, img_path: Path) -> Path:
    """Exporting PDF receipt"""
    page = browser.page()
    receipt_html = page.inner_html("#receipt")

    pdf = PDF()
    pdf_path = RECEIPTS / f"receipt_{order_no}.pdf"
    pdf.html_to_pdf(receipt_html, str(pdf_path))
    # embedding screenshot to receipt
    pdf.add_files_to_pdf(files=[str(img_path)], target_document=str(pdf_path), append=True)

def reset_form_for_next_order():
    """Reset form"""
    browser.page().click("text=Order another robot")

def archive_receipts():
    """Final zip"""
    Archive().archive_folder_with_zip(str(RECEIPTS), str(ZIP_TARGET))



