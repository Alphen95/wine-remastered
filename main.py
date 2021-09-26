import datetime, pandas,argparse
from http.server import HTTPServer, SimpleHTTPRequestHandler
from collections import defaultdict
import pathlib
import os

from jinja2 import Environment, FileSystemLoader, select_autoescape

parser = argparse.ArgumentParser(
    description='Запускает сайт по продаже вина, с использованием .xlsx файла.'
)
path_folder = str(pathlib.Path(__file__).parent.absolute())
os.chdir(path_folder)
opening_year = 1920
parser.add_argument('xlsx_file', metavar="N", type=str, help='Пользовательский файл')
args = parser.parse_args()
xlsx_table_filename = args.xlsx_file
enviroment = Environment(
    loader=FileSystemLoader('.'),
    autoescape=select_autoescape(['html', 'xml'])
)

organized_wine_cards = defaultdict(list)
template = enviroment.get_template("template.html")

def main():
  excelfile = pandas.read_excel(xlsx_table_filename, sheet_name='Лист1', usecols=['Категория','Название', 'Сорт','Цена','Картинка',"Акция"])
  wine_cards = excelfile.to_dict(orient='record')
  for bottle in wine_cards:
    organized_wine_cards[bottle["Категория"]].append({"class":"Сорт винограда - {}".format(bottle["Сорт"]) if bottle["Сорт"] !="None" else " ","cost":"{} р.".format(str(bottle["Цена"])),"picture":"images/{}".format(bottle["Картинка"]),"name":bottle["Название"],"stock":bottle["Акция"]})
      
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

main()