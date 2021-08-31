import pandas as pd
import xlsxwriter

from selenium import webdriver
from selenium.webdriver.chrome.options import Options

chrome_options = Options()
# chrome_options.add_argument('--headless')

chrome_path = r".\chromedriver.exe"
driver = webdriver.Chrome(options=chrome_options)

#  get the list of products to search
products = pd.read_excel('Products.xlsx', sheet_name='Tabelle1')

# #  create excel file for results
# workbook = xlsxwriter.Workbook('result.xlsx')
# worksheet = workbook.add_worksheet()

# # set headers
# worksheet.write("A1", "Product Name")
# worksheet.write("B1", "Price")
# worksheet.write("C1", "Shop Name")
# worksheet.write("D1", "Location")
# worksheet.write("E1", "Product url")
# worksheet.write("F1", "Description")
# worksheet.write("G1", "Features")
# worksheet.write("H1", "Inclusion")
# worksheet.write("I1", "Specification")
# worksheet.write("J1", "Others")
# worksheet.write("K1", "Image url")

def fix_url(links, website):
	urls = []
	for link in links:
		url = link.get_attribute("href")
		if url[0] == "/":
			url = ''.join(website) + ''.join(url)
		elif not "http" in url:
			url = ''.join(website) + ''.join(url)
		urls.append(url)
	return urls


def tbs(url):
	driver.get(url)

	product_name = driver.find_element_by_xpath('//*[@id="product_description"]/div[2]/h1').text
	price = driver.find_element_by_xpath('//*[@id="product_description"]/div[4]/div[1]/p').text
	main_desc = driver.find_element_by_xpath('//*[@id="product_description"]/div[3]').text
	gallery = driver.find_elements_by_xpath('//a[contains(@href, "/img/gallery/")]')
	image_url = fix_url(gallery, 'https://www.team-blacksheep.com/')
	product_info = driver.find_element_by_xpath('//*[@id="product_text"]').text
	print(product_info)

	# return data = {
	# 	"product_name": product_name,
	#	"price": price
	# 	"description": description,s
	# 	"features": features,
	# 	"inclusion": inclusion,
	# 	"specification": specification,
	# 	"others": others,
	#	"image_url": image_url
	# }

def main():
	for i, url in enumerate(products['Link to the product']):
		if products['Shopname'][i] == 'TBS':
			tbs(url)
	

if __name__ == '__main__':
	main()
	worksheet.close()
	driver.close()
	driver.quit()