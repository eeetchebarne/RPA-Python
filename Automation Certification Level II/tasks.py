from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.FileSystem import FileSystem
from RPA.Archive import Archive

@task
def order_robots_from_RobotSpareBin():
    """
    Main task that orchestrates the robot ordering process.
    Configures browser, opens order website, processes all orders,and archives receipts.
    """
    browser.configure(slowmo=100)
    open_robot_order_website()
    orders = get_orders()
    for o in orders:
        fill_the_form(o)
    archive_receipts("temp_output", "output/receipts.zip")

def open_robot_order_website():
    """Navigates to the RobotSpareBin order page."""
    browser.goto("https://robotsparebinindustries.com/")
    page = browser.page()
    page.click("//a[text() = 'Order your robot!']")

def get_orders():
    """
    Downloads and parses the orders CSV file and return a table.
    """
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)
    return Tables().read_table_from_csv(path="orders.csv", header=True)

def close_annoying_modal():
    """Closes the popup modal if it appears on the page."""
    page = browser.page()
    selector = "//button[text() = 'OK']"
    if page.is_visible(selector,timeout=3000):
        page.click(selector)

def fill_the_form(row):
    """Fills out the robot order form with data from a single order row."""
    page = browser.page()
    close_annoying_modal()

    page.select_option("#head",["Roll-a-thor head","Peanut crusher head","D.A.V.E head","Andy Roid head","Spanner mate head","Drillbit 2000 head"][int(row["Head"])-1])
    page.click(f"#id-body-{row['Body']}")
    page.fill("//input[@placeholder='Enter the part number for the legs']",row["Legs"])
    page.fill("#address",row["Address"])
    
    screenshot = screenshot_robot(row["Order number"])
    page.click("#order")
    
    while page.is_visible("//div[@class='alert alert-danger']",timeout=3000):
        page.click("#order")
    
    pdf_file = store_receipt_as_pdf(row["Order number"])
    embed_screenshot_to_receipt(screenshot,pdf_file)
    page.click("#order-another")

def store_receipt_as_pdf(order_number):
    """Converts the order receipt HTML to a PDF file."""
    fs = FileSystem()
    fs.create_directory("temp_output")
    pdf = PDF()
    receipt_html = browser.page().locator("#receipt").inner_html()
    pdf_output_path = f"temp_output/{order_number}.pdf"
    pdf.html_to_pdf(receipt_html, pdf_output_path)
    return pdf_output_path

def screenshot_robot(order_number):
    """Takes a screenshot of the robot preview image."""
    page = browser.page()
    page.click("#preview")
    page.wait_for_selector("#robot-preview", timeout=30000)
    screenshot_path = f"temp_output/screenshot_{order_number}.png"
    page.screenshot(path=screenshot_path)
    return screenshot_path

def embed_screenshot_to_receipt(screenshot, pdf_file):
    """Embeds the robot screenshot into the receipt PDF."""
    PDF().add_watermark_image_to_pdf(
        image_path=screenshot,
        source_path=pdf_file,
        output_path=pdf_file
    )

def archive_receipts(source_dir, zip_path):
    """Archives all receipt PDFs into a ZIP file and cleans up temporary files."""
    fs = FileSystem()
    Archive().archive_folder_with_zip(source_dir, zip_path, include="*.pdf")
    if fs.does_directory_exist(source_dir):
        files = fs.find_files(f"{source_dir}/*")
        for file in files:
            fs.remove_file(file)
        fs.remove_directory(source_dir)