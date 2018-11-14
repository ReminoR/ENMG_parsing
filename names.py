import re
import requests
from bs4 import BeautifulSoup


url_man = 'https://kakzovut.ru/man.html'
url_woman = 'https://kakzovut.ru/woman.html'


def get_names(url):
	r = requests.get(url)

	soup = BeautifulSoup(r.text, 'html.parser')
	names = soup.find_all("div", class_="nameslist")
	list_names = []

	for i in names:
		name = i.find('a').text
		if not ( re.search(r'\(|\)', name) ):
			list_names.append(name)
			

	return list_names

man_names = get_names(url_man) + \
	['Алекспандр', 
	'Аллександр', 
	'Анатлий', 
	'Артем', 
	'Атамурад', 
	'Батыр',
	'Валерний',
	'Валерний',
	'Владимр',
	'Гафарали',
	'Данила',
	'Жосурбек',
	'Зиннур',
	'Ильдар',
	'Ильнар',
	'Ильяс',
	'Ирек',
	'Левон',
	'Магеррам',
	'Масалим',
	'Неъматилло',
	'Ниязи',
	'Петр',
	'Раббон',
	'Рафис',
	'Рифгат',
	'Уткиржон',
	'Фаим',
	'Федор',
	'Хавадис',
	'Хайдар',
	'Ханиф',
	'Шамшитин',
	'Эдвард']

woman_names = get_names(url_woman) + \
	['Алена', 
	'Алефтина', 
	'Дэнсэма',
	'Зинаила',
	'Людмилла',
	'Магуфура',
	'Марьяна',
	'Наталия',
	'Нэля',
	'Просковья',
	'Сима',
	'Шалва']






