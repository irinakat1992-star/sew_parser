import cloudscraper
from bs4 import BeautifulSoup
import pandas as pd
import time
import json


scraper = cloudscraper.create_scraper(
    browser={
        'browser': 'chrome',
        'platform': 'windows',
        'mobile': False
    }
)

section = 'manekeny'

def collect_products_price(section):
   page_num = 1
   data = []

   headers = {
       'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
       'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
       'Accept-Language': 'en-US,en;q=0.9,ru;q=0.8',
   }

   while True:
       url = f'https://sew-irk.ru/catalog/{section}/?PAGEN_1={page_num}'
       # Отправляем запрос
       response = scraper.get(url, headers=headers)
       response.raise_for_status()

       soup = BeautifulSoup(response.text, 'html.parser')

       products = soup.find_all('div', class_='item product sku lol4')

       for products in products:
           name = products.find('span', class_='middle').text
           price = products.find('a', class_='price').text

           data.append({'product_name': name, 'product_price': price})

       next_button = soup.find('a', string='Вперед')
       if next_button:
           page_num += 1
           time.sleep(1)
       else:
           pagination = soup.find('div', class_='pagination')
           if pagination:
               has_next = False
               for link in pagination.find_all('a'):
                   if link.text.strip().isdigit() and int(link.text.strip()) > page_num:
                       has_next = True
                       break
               if has_next:
                   page_num += 1
                   time.sleep(1)
               else:
                   break
           else:
               break

   return data

print(f"\nСбор данных завершен")
products_price = collect_products_price(section)

if products_price:
    """Сохранение данных в Excel"""
    df = pd.DataFrame(products_price)
    df.to_excel('products_price.xlsx', index=False, engine='openpyxl')

    """Сохранение данных в CSV"""
    df = pd.DataFrame(products_price)
    df.to_csv('products_price.csv', index=False, encoding='utf-8-sig')

    """Сохранение данных в JSON"""
    with open('products_price.json', 'w', encoding='utf-8') as f:
        json.dump(products_price, f, ensure_ascii=False, indent=2)