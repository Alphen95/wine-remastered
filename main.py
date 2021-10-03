import datetime
import argparse
import pathlib
import os

import pandas

from http.server import HTTPServer, SimpleHTTPRequestHandler
from collections import defaultdict

from jinja2 import Environment, FileSystemLoader, select_autoescape

parser = argparse.ArgumentParser(
    description='Запускает сайт по продаже вина, с использованием .xlsx файла.'
)
#эта штука важная, без неё будет криво работать погдгрузка файлов, потому что не все люди будут напрямую из коммандой строки запускать, а через .bat или что-то похожее.
opening_year = 1920
parser.add_argument('xlsx_file_name', metavar="file", type=str, help='Пользовательский файл')
os.chdir(str(pathlib.Path(__file__).parent.absolute()))
args = parser.parse_args()
xlsx_table_filename = args.xlsx_file_name
enviroment = Environment(
    loader=FileSystemLoader('.'),
    autoescape=select_autoescape(['html', 'xml'])
)
organized_wine_cards = defaultdict(list)
template = enviroment.get_template("template.html")


def main():
  wine_table = pandas.read_excel(xlsx_table_filename, sheet_name='Лист1', usecols=['Категория','Название', 'Сорт','Цена','Картинка',"Акция"])
  wine_cards = wine_table.to_dict(orient='record')
  for bottle in wine_cards:
    organized_wine_cards[bottle["Категория"]].append(bottle)
      
  today = datetime.datetime.now()
  years_with_you = today.year - opening_year

  rendered_page = template.render(
      years_with_our_clients="Уже {} лет с вами.".format(years_with_you),
      wine_cards = organized_wine_cards
  )

  with open('index.html', 'w', encoding="utf8") as file:
      file.write(rendered_page)
  server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
  server.serve_forever()

if __name__ == "__main__":
  main()