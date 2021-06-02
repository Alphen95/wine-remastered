import datetime, pandas,argparse
from pprint import pprint
from http.server import HTTPServer, SimpleHTTPRequestHandler
from collections import defaultdict

from jinja2 import Environment, FileSystemLoader, select_autoescape

parser = argparse.ArgumentParser(
    description='Запускает сайт по продаже вина, с использованием .xlsx файла.'
)
first_opening_year = 1920
parser.add_argument('--xlsx_file', help='Пользовательский файл')
args = parser.parse_args()
if type(args.xlsx_file) != str:
  file = 'base_file.html'
else:
  print(str(args.xlsx_file))
  file = args.xlsx_file
enviroment = Environment(
    loader=FileSystemLoader('.'),
    autoescape=select_autoescape(['html', 'xml'])
)

organized_dict_wine_cards = defaultdict(list)
template = enviroment.get_template(file)

excelfile = pandas.read_excel('wine_placeholder.xlsx', sheet_name='Лист1', usecols=['Категория','Название', 'Сорт','Цена','Картинка',"Акция"])
wine_cards_dict = excelfile.to_dict(orient='record')
for bottle in wine_cards_dict:
  print
  try:
    bottle["Сорт"] = "Сорт винограда - " + bottle["Сорт"]
  except:
    bottle["Сорт"] = " "
  bottle["Цена"] = str(bottle["Цена"]) + "р."
  bottle["Картинка"] = "images/" + bottle["Картинка"]
  bottle_card_type =bottle["Категория"]
  organized_dict_wine_cards[bottle_card_type].append(bottle)
    
today = datetime.datetime.now()
years_with_you = today.year - first_opening_year

rendered_page = template.render(
    years_with_our_clients="Уже "+str(years_with_you)+" лет с вами.",
    wine_cards = organized_dict_wine_cards
)

with open('index.html', 'w', encoding="utf8") as file:
    file.write(rendered_page)
server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
pprint(organized_dict_wine_cards)
server.serve_forever()