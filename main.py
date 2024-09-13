import bs4
import pandas as pd
import requests
import sys

df = pd.read_excel("result.xlsx")
def make_soup(url: str):
    response = requests.get(url)
    soup = bs4.BeautifulSoup(response.text, "html.parser")
    return soup
site_url = "https://liliome.ir/shop/?orderby=date"
main_soup = make_soup(site_url)
product_body = main_soup.find(name="div", class_="products")
if not product_body:
    print("couldnt find body")
    sys.exit()
all_products = product_body.find_all(name="div", class_="product")
i = 0
att_list = ["طبع", "جنسیت", "فصل"]
print("products len", len(all_products))
for product in all_products:
    product_url = all_products[i].find(name="a").get(key="href")
    product_soup = make_soup(product_url)
    title = product_soup.find(name="h1", class_="product_title")
    print("title =", title)
    if title:
        title = title.text
        title = title.strip()
        title_list = title.split("|")
        eng_title = title_list[0]
        per_title = title_list[1]
    else:
        eng_title = ""
        per_title = ""
    price = product_soup.find(name="p", class_="price")
    price = price.find(name="bdi")
    if price:
        price = price.text
        price = price.replace("تومان", "")
        price = price.strip()
        price = price.replace(",", "")
    else:
        price = "نا موجود"
    print("price =", price)

    att_value_list = []
    brand = product_soup.find(name="a", rel="tag")
    if brand:
        brand = brand.text
        att_value_list.append(brand)
    print("brand =", brand)
    table = product_soup.find(name="table", class_="table")
    table_rows = table.find_all(name="tr")
    for row in table_rows:
        if row:
            row_columns = row.find_all(name="span")
            if row_columns and row_columns[0].text in att_list:
                try:
                    text = row_columns[1].text.replace(r"\xa0", " ").strip()
                    att_value_list.append(text)
                except IndexError:
                    att_value_list.append("")
                    continue

    value_list_to_save = str(att_value_list).replace("]", "").replace("[", "").replace(r"\xa0", " ").strip()
    print("values =", value_list_to_save)
    df.loc[i] = [title, eng_title, per_title, price, value_list_to_save]
    print(f"saved row {i}")
    i += 1
# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('pandas_simple.xlsx', engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
df.to_excel(writer, sheet_name='Sheet1')
# Get the xlsxwriter workbook and worksheet objects.
workbook  = writer.book
worksheet = writer.sheets['Sheet1']

# Add the cell formats.
format_right_to_left = workbook.add_format({'reading_order': 2})

# Change the direction for the worksheet.
worksheet.right_to_left()

# Make the column wider for visibility and add the reading order format.
worksheet.set_column('B:B', 30, format_right_to_left)

# Close the Pandas Excel writer and output the Excel file.
writer._save()
