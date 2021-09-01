import pandas as pd
import xlsxwriter

from selenium import webdriver
from selenium.webdriver.chrome.options import Options

chrome_options = Options()
chrome_options.add_argument('--headless')

chrome_path = r".\chromedriver.exe"
driver = webdriver.Chrome(options=chrome_options)

#  get the list of products to search
products = pd.read_excel('Products.xlsx', sheet_name='Tabelle1')

# #  create excel file for results
workbook = xlsxwriter.Workbook('Result.xlsx')
worksheet = workbook.add_worksheet()

# set headers
worksheet.write("A1", "Product Name")
worksheet.write("B1", "Price")
worksheet.write("C1", "Shop Name")
worksheet.write("D1", "Location")
worksheet.write("E1", "Product url")
worksheet.write("F1", "Description")
worksheet.write("G1", "Features")
worksheet.write("H1", "Inclusion")
worksheet.write("I1", "Specification")
worksheet.write("J1", "Image url")

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

	#get image urls
	gallery = driver.find_elements_by_xpath('//a[contains(@href, "/img/gallery/")]')
	image_url = fix_url(gallery, 'https://www.team-blacksheep.com/')

	features = ''
	specification = ''
	inclusion = ''

	new_product_infos = driver.find_elements_by_xpath('//*[@id="product_text"]//*')

	filter = 'desc'
	for info in new_product_infos:
		if info.text == 'FEATURES':
			filter = 'features'
		elif info.text == 'SPECIFICATION' or info.text == 'PRODUCT SPECIFICATIONS':
			filter = 'specs'
		elif 'INCLUDES' in info.text:
			filter = 'inclusion'
		elif info.text == 'MORE INFORMATION':
			filter == 'others'

		if filter == 'desc':
			main_desc = main_desc + '\n {}'.format(info.text)
		elif filter == 'features':
			features = features + '{}, '.format(info.text)
		elif filter == 'specs':
			specification = specification + '{}, '.format(info.text)
		elif filter == 'inclusion':
			inclusion = inclusion + '{}, '.format(info.text)

	data = {
		"product_name": product_name,
		"price": price,
		"description": main_desc,
		"features": features,
		"inclusion": inclusion,
		"specification": specification,
		"image_url": ','.join(map(str, image_url)) 
	}

	return data

def main():
	for i, url in enumerate(products['Link to the product']):
		if products['Shopname'][i] == 'TBS':
			res = tbs(url)
			worksheet.write(i+1, 0, res['product_name'])
			worksheet.write(i+1, 1, res['price'])
			worksheet.write(i+1, 2, products['Shopname'][i])
			worksheet.write(i+1, 3, products['Location'][i])
			worksheet.write(i+1, 4, url)
			worksheet.write(i+1, 5, res['description'])
			worksheet.write(i+1, 6, res['features'])
			worksheet.write(i+1, 7, res['inclusion'])
			worksheet.write(i+1, 8, res['specification'])
			worksheet.write(i+1, 9, res['image_url'])

if __name__ == '__main__':
	main()
	workbook.close()
	driver.close()
	driver.quit()