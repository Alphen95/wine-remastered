import datetime, pandas
from pprint import pprint
from http.server import HTTPServer, SimpleHTTPRequestHandler
from collections import defaultdict

from jinja2 import Environment, FileSystemLoader, select_autoescape

enviroment = Environment(
    loader=FileSystemLoader('.'),
    autoescape=select_autoescape(['html', 'xml'])
)

organized_dict_wine_cards = defaultdict(list)
template = enviroment.get_template('base_file.html')

excelfile = pandas.read_excel('wine3.xlsx', sheet_name='Лист1', usecols=['Категория','Название', 'Сорт','Цена','Картинка',"Акция"])
wine_cards_dict = excelfile.to_dict(orient='record')
for bottle in wine_cards_dict:
  print(bottle["Название"])
  try:
    bottle["Сорт"] = "Сорт винограда - " + bottle["Сорт"]
  except:
    bottle["Сорт"] = " "
  bottle["Цена"] = str(bottle["Цена"]) + "р."
  bottle["Картинка"] = "images/" + bottle["Картинка"]
  if bottle["Категория"] == "Белые вина":
    organized_dict_wine_cards["Белые вина"].append(bottle)
  elif bottle["Категория"] == "Напитки":
    organized_dict_wine_cards["Напитки"].append(bottle)
  else:
    organized_dict_wine_cards["Красные вина"].append(bottle)
    
date_today = str(datetime.date.today())
year_today = date_today[:4]
years_with_you = int(year_today) - 1920

rendered_page = template.render(
    years_with_our_clients="Уже "+str(years_with_you)+" лет с вами.",
    wine_cards = organized_dict_wine_cards
)

with open('index.html', 'w', encoding="utf8") as file:
    file.write(rendered_page)
pprint(organized_dict_wine_cards)
server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
server.serve_forever()