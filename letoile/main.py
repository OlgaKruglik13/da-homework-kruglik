from selenium import webdriver
from bs4 import BeautifulSoup
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait 
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import openpyxl
import re


driver = webdriver.Chrome()
service = Service(executable_path=driver)
chrome_options = Options()
user_agent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36'
chrome_options.add_argument(f"user-agent={user_agent}")


base_url = "https://www.letu.ru/browse/makiyazh"

driver.get(base_url)

try:
  element = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.CLASS_NAME, "results-listing-content"))
  )
finally:
  workbook = openpyxl.Workbook() #Создание Excel-файла
  sheet = workbook.active
  sheet.append([ "Название товара", "Название бренда", "Актуальная цена", "Старая цена", "Скидка", "Количество отзывов", "Оценка"])

  for page_num in range(1, 6): #Перебор по первым 5 страницам
    url = f"{base_url}/page-{page_num}"
    html_content = driver.page_source #Извлечение HTML
    soup = BeautifulSoup(html_content, "html.parser") #Парсинг с помощью BeautifulSoup
    products = soup.find_all("div", class_="product-tile results-listing-content__item") #Извлечение информации о товарах
    

    for product in products:
        product_title = product.select_one('.product-tile-name__text [data-at-product-tile-title]').text.strip() #Название товара
        brand = product.select_one('.product-tile-name__text--brand').text.strip()
        name = product_title.replace(brand, '').strip()

        brand = product.find("span", {'class': 'product-tile-name__text--brand'}).text.strip() #Название бренда

        actual_price_element = product.find('span', {'class': 'product-tile-price__text product-tile-price__text--actual'}) #Актуальная цена
        actual_price = float(actual_price_element.text.replace('&nbsp;', '').replace('\xa0', '').replace('₽', '').strip())

        old_price_element = product.find('span', {'class': 'product-tile-price__text product-tile-price__text--old'}) #Старая цена
        if old_price_element:
             old_price_text = old_price_element.text.replace('&nbsp;', '').replace('\xa0', '').replace('₽', '').strip()
             try:
                   old_price = float(old_price_text)
             except ValueError:
                   old_price = 0.0
        else:
             old_price = 0.0

        sale_element = product.find('span', {'class': 'product-tile-price__text product-tile-price__text--discount product-tile-price__text--has-description'}) #Скидка
        if sale_element:
             sale_text = sale_element.get_text(strip=True)
             match = re.search(r'\d+\.?\d*', sale_text) #Использование регулярного выражения для извлечения числового значения скидки
             if match:
                   sale = float(match.group())
             else:
                   sale = 0.0
        else:
             sale = 0.0

        review_element = product.find('div', {'class': 'product-tile-rating'}) #Количество отзывов
        if review_element:
             review = int(review_element.text.strip())
        else:
             review = 0
          
        stars = product.find_all('div', {'class': 'rating-symbol'}) #Оценка товара
        full_stars = 0
        half_stars = 0
        for star in stars:
             if 'icon-rating-star-v2' in str(star) and 'icon-rating-star-v2-half' not in str(star):
                   full_stars += 1
             elif 'icon-rating-star-v2-half' in str(star):
                   half_stars += 1
        rating = full_stars + (half_stars / 2)
          
        sheet.append([name, brand, actual_price, old_price, sale, review, rating])
            
driver.quit() #Закрытие браузера

workbook.save("ЛЭТУАЛЬ.xlsx") #Сохранение Excel-файла
print("Данные сохранены в файл ЛЭТУАЛЬ.xlsx")

