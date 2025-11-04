from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
import time
import csv
import os


# Setup
userInput = input("Enter the product you want to search for: ")
# Construct the URL with user input
url = 'https://www.noon.com/egypt-ar/search/?q={userInput}'.format(userInput=userInput)
# List to store product data
productData = []
# Initialize the Chrome driver

service = Service(executable_path="chromedriver.exe")
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

# Function
def noon(link):
    try:
        driver.get(link)
        time.sleep(2)  # Wait for the page to load
        # Fix: Correct locator strategy
        # searchBar = driver.find_element(By.CLASS_NAME, "DesktopInput_searchInput__R44H1")
        # searchBar.send_keys(userInput)
        # time.sleep(2)
        productList = driver.find_elements(By.CLASS_NAME, "ProductBoxLinkHandler_linkWrapper__b0qZ9")
        for product in productList:
            html_code = product.get_attribute('outerHTML')
            soup = BeautifulSoup(html_code, 'html.parser')
            try:
                productName = soup.find(class_='ProductDetailsSection_title__JorAV').text.strip()
            except AttributeError:
                productName = "Product name not found"
            try:    
                productPrice = soup.find(class_='Price_amount__2sXa7').text.strip()
            except AttributeError:
                productPrice = "Price not found"
            try:
                productRating = soup.find(class_='RatingPreviewStar_textCtr__sfsJG').text.strip()
            except AttributeError:
                productRating = "Rating not found"
            try:
                productOldPrice = soup.find(class_='Price_oldPrice__ZqD8B').text.strip()
            except AttributeError:
                productOldPrice = "Old price not found"
            try:
                productDiscount = soup.find(class_='PriceDiscount_discount__1ViHb PriceDiscount_pBox__eWMKb').text.strip()
            except AttributeError:
                productDiscount = "No discount available"
            try:
                productLink = soup.find(class_='ProductBoxLinkHandler_productBoxLink__FPhjp')['href']  
            except TypeError:
                productLink = "Link not found"
            productData.append({
                'productName': productName,
                'productPrice': productPrice,
                'productRating': productRating,
                'productOldPrice': productOldPrice,
                'productDiscount': productDiscount,
                'productLink': productLink
            })

    except Exception as e:
        print(f"An error occurred: {e}")
        driver.quit()
        return


def export_csv():
    __path__ = f"C:/Users/DELL/Desktop/selenium/Noon/{userInput}.csv" 
    keys= productData[0].keys()
    # Write to CSV file
    try:
        with open(__path__,'w',newline='',encoding='utf-8') as output_file:
            dict_writer = csv.DictWriter(output_file, keys)
            dict_writer.writeheader()
            dict_writer.writerows(productData) 
        output_file.close()
        print(f"Data successfully written to {__path__}")       
    except IOError as e:
        print(f"Error writing to file: {e}")


# Call the function (Fix: removed extra ')')

noon(url)
export_csv()



def export_h():
    if not productData:
        print("No product data to export.")
        return

    html_head = """
    <html>
    <head>
        <meta charset='utf-8'>
        <title>Noon Product Results</title>
        <style>
            body { font-family: Arial, sans-serif; background-color: #f9f9f9; padding: 20px; }
            h2 { color: #333333; }
            table { width: 100%; border-collapse: collapse; margin-top: 20px; }
            th, td { border: 1px solid #dddddd; padding: 8px 12px; text-align: left; }
            th { background-color: #004080; color: white; }
            tr:nth-child(even) { background-color: #f2f2f2; }
            a { color: #0077cc; text-decoration: none; }
        </style>
    </head>
    <body>
    """

    html_body = f"<h2>Search Results for: <em>{userInput}</em></h2>\n"
    html_body += "<table>\n<tr>"
    html_body += "".join(f"<th>{key}</th>" for key in productData[0].keys()) + "</tr>\n"

    for item in productData:
        html_body += "<tr>"
        for key, value in item.items():
            if "Link" in key and value.startswith("/"):
                full_link = f"https://www.noon.com{value}"
                html_body += f'<td><a href="{full_link}" target="_blank">View Product</a></td>'
            else:
                html_body += f"<td>{value}</td>"
        html_body += "</tr>\n"

    html_footer = "</table>\n</body></html>"
    html_output = html_head + html_body + html_footer

    # Save to HTML file
    folder = "C:/Users/DELL/Desktop/selenium/Noon" 
    os.makedirs(folder, exist_ok=True)
    file_path = os.path.join(folder, f"{userInput}.html")

    try:
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(html_output)
        print(f"[✅] Data successfully written to {file_path}")
    except IOError as e:
        print(f"[❌] Error writing HTML file: {e}")
export_h()
