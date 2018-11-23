import re
import os, os.path

# Custom
import docx
import sound
import names

path_polyneuropathy_files = 'files/pnp'
number_docs = len([name for name in os.listdir(path_polyneuropathy_files) if os.path.isfile(os.path.join(path_polyneuropathy_files, name))])
files = os.listdir(path_polyneuropathy_files)

def main():
	create_csv()
	sound.play()

def parse(regular_expression, text, category, flag=0, end=False):
	search = re.search(regular_expression, text, flag)
	result = ''

	if search and category == 'age':
		result = search.group(1) + ';'
	elif not search and category == 'age':
		result = ';'

	if search and category == 'srv_motor_7':
		result = search.group(2) + ';' + search.group(3) + ';'
	elif not search and category == 'srv_motor_7':
		result = ';;' #0
	elif search and category == 'srv_motor_9':
		result = search.group(2) + ';' + search.group(3) + ';' + search.group(10) + ';'	
	elif not search and category == 'srv_motor_9':
		result = ';;;' #0

	if search and category == 'srv_sensor_9':
		result = search.group(3) + ';' + search.group(10) + ';'
	elif not search and category == 'srv_sensor_9':
		result = ';;' #0

	if search and category == 'f_wave_bool':
		result = '1;'
	elif not search and category == 'f_wave_bool':
		result = '0;'

	if search and category == 'f_wave':
		result = search.group(2) + ';'
	elif not search and category == 'f_wave':
		result = ';' #0

	if search and category == 'h_reflex_bool':
		result = '1;'
	elif not search and category == 'h_reflex_bool':
		result = '0;'

	if search and category == 'h_reflex':
		result = search.group(4) + ';' + search.group(7) + ';'
	elif not search and category == 'h_reflex':
		result = ';;' #0

	if search and category == 'name':
		result = search.group(2) + ';'
	elif not search and category == 'name':
		result = ';'

	if search and category == 'last_name':
		result = search.group(1) + ';'
	elif not search and category == 'last_name':
		result = ';'

	if category == 'gender':
		result = ';'

	if search and category == 'hands':
		result = 'hands;'
	elif not search and category == 'hands':
		result = 'legs;'

	if search and category == 'diagnosis_true':
		result = '1;'
	elif not search and category == 'diagnosis_true':
		result = '0;'

	if search and category == 'diagnosis_false':
		result = '1'

	#Удаляем последнюю ';'
	if end:
		result = result[:-1]

	return result

def create_csv():
	count = 1
	pattern_srv7 = '\n+([\d,]+)\n+([\d,]+)\n+([\d,]+)\n+([\d,]+)\n+([\d,]+)\n+([\d,]+)\n+([\d,]+)'
	pattern_srv9 = '\n+([\d,]+)\n+([\d,]+)\n+([\d,]+)\n+([\d,]+)\n+([\d,]+)\n+([\d,]+)\n+([\d,]+)\n+([\d,]+)\n+([\d,]+)'
	pattern_f_wave = '\n+([\d,]+)\n+([\d,]+)'
	pattern_hr = '\n+([\d,]+)\n+([\d,]+)\n+([\d,]+)\n+([\d,]+)\n+([\d,]+)\n+([\d,]+)'

	summary_file = open('files/summary.csv', 'w')
	summary_file.write(
		'id;' + 
		'age;' +
		'last_name;' +
		'name;' +
		'gender;' +
		
		#Ноги
		#СРВ моторная
		# Abductor hallucis, Tibialis
		'rAhT1_lat;' +
		'rAhT1_ampl;' +
		'rAhT2_lat;' +
		'rAhT2_ampl;' +
		'rAhT2_speed;' +

		'lAhT1_lat;' +
		'lAhT1_ampl;' +
		'lAhT2_lat;' +
		'lAhT2_ampl;' +
		'lAhT2_speed;' +

		# Extensor digitorum brevis, Peroneus
		'r_EdbP1_lat;' +
		'r_EdbP1_ampl;' +
		'r_EdbP2_lat;' +
		'r_EdbP2_ampl;' +
		'r_EdbP2_speed;' +

		'l_EdbP1_lat;' +
		'l_EdbP1_ampl;' +
		'l_EdbP2_lat;' +
		'l_EdbP2_ampl;' +
		'l_EdbP2_speed;' +


		#СРВ Сенсорная
		# Suralis
		'r_S_ampl;' +
		'r_S_speed;' +
		'l_S_ampl;' +
		'l_S_speed;' +

		# Peroneus superficialis
		'r_Ps_ampl;' +
		'r_Ps_speed;' +
		'l_Ps_ampl;' +
		'l_Ps_speed;' +

		# F-волна
		'f_wave;'
		'rAhT_f_wave;' +
		'lAhT_f_wave;' +
		'r_EdbP_f_wave;' +
		'l_EdbP_f_wave;' +

		# H-рефлекс 
		# Soleus, Tibialis
		'h_reflex;'
		'r_ST_hr_lat;' +
		'r_ST_hr_h/m;' +
		'l_ST_hr_lat;' +
		'l_ST_hr_h/m;' +

		'pnp_diagnosis' +

		'\n')

	for i in files:
		text = docx.get_docx_text('./' + path_polyneuropathy_files + '/' + i)

		age = parse(r'Пациент:.+?(\d+)', text, 'age')
		last_name = parse(r'Пациент:\s*(\w+)\s*(\w+)\s*(\w+)', text, 'last_name')
		name = parse(r'Пациент:\s*(\w+)\s*(\w+)\s*(\w+)', text, 'name')

		if name[:-1] in names.man_names:
			gender = '1;'
		elif name[:-1] in names.woman_names:
			gender = '0;'
		else:
			gender = '1;'
		
		#Находим руки
		hands_legs = parse(r'(Abductor pollicis brevis, Medianus, C8 T1)|(n. Medianus)|(Abductor digiti minimi, Ulnaris, C8 T1)|(n. Ulnaris)', text, 'hands', re.S)

		if hands_legs == 'hands;':
			continue

		#Ноги
		rAhT1 = parse(r'СРВ моторная.*?пр\., Abductor hallucis, Tibialis.*?(медиальная лодыжка|предплюсна)' + pattern_srv7, text, 'srv_motor_7', re.S)
		rAhT2 = parse(r'СРВ моторная.*?пр\., Abductor hallucis, Tibialis.*?(подколенная ямка)' + pattern_srv9, text, 'srv_motor_9', re.S)
		lAhT1 = parse(r'СРВ моторная.*?лев\., Abductor hallucis, Tibialis.*?(медиальная лодыжка|предплюсна)' + pattern_srv7, text, 'srv_motor_7', re.S)
		lAhT2 = parse(r'СРВ моторная.*?лев\., Abductor hallucis, Tibialis.*?(подколенная ямка)' + pattern_srv9, text, 'srv_motor_9', re.S)

		r_EdbP1 = parse(r'СРВ моторная.*?пр\., Extensor digitorum brevis, Peroneus.*?(предплюсна)' + pattern_srv7, text, 'srv_motor_7', re.S)
		r_EdbP2 = parse(r'СРВ моторная.*?пр\., Extensor digitorum brevis, Peroneus.*?(головка малоберцовой кости|подколенная ямка)' + pattern_srv9, text, 'srv_motor_9', re.S)
		l_EdbP1 = parse(r'СРВ моторная.*?лев\., Extensor digitorum brevis, Peroneus.*?(предплюсна)' + pattern_srv7, text, 'srv_motor_7', re.S)
		l_EdbP2 = parse(r'СРВ моторная.*?лев\., Extensor digitorum brevis, Peroneus.*?(головка малоберцовой кости|подколенная ямка)' + pattern_srv9, text, 'srv_motor_9', re.S)

		r_S = parse(r'СРВ сенсорная.*?пр\., n.Suralis, S1-S2.*?\d+\n+(1|Ср\. треть голени|Средняя треть голени|нижняя треть голени)' + pattern_srv9, text, 'srv_sensor_9', re.S)
		l_S = parse(r'СРВ сенсорная.*?лев\., n.Suralis, S1-S2.*?\d+\n+(1|Ср\. треть голени|Средняя треть голени|нижняя треть голени)' + pattern_srv9, text, 'srv_sensor_9', re.S)

		r_Ps = parse(r'СРВ сенсорная.*?пр\., n.Peroneus superficialis, L4-S1.*?(1|Ср\. треть голени|Уровень лодыжек)' + pattern_srv9, text, 'srv_sensor_9', re.S)
		l_Ps = parse(r'СРВ сенсорная.*?лев\., n.Peroneus superficialis, L4-S1.*?(1|Ср\. треть голени|Уровень лодыжек)' + pattern_srv9, text, 'srv_sensor_9', re.S)

		f_wave = parse(r'Параметры F-волны', text, 'f_wave_bool')
		rAhT_f_wave = parse(r'Параметры F-волны.*?пр\., Abductor hallucis, Tibialis.*?\n+(\d+)' + pattern_f_wave, text, 'f_wave', re.S)
		lAhT_f_wave = parse(r'Параметры F-волны.*?лев\., Abductor hallucis, Tibialis.*?\n+(\d+)' + pattern_f_wave, text, 'f_wave', re.S)
		r_EdbP_f_wave = parse(r'Параметры F-волны.*?пр\., Extensor digitorum brevis, Peroneus.*?\n+(\d+)' + pattern_f_wave, text, 'f_wave', re.S)
		l_EdbP_f_wave = parse(r'Параметры F-волны.*?лев\., Extensor digitorum brevis, Peroneus.*?\n+(\d+)' + pattern_f_wave, text, 'f_wave', re.S)

		h_reflex = parse(r'H-рефлекс', text, 'h_reflex_bool')
		r_ST_hr = parse(r'H-рефлекс.*?пр\., Soleus, Tibialis.*?(H-рефлекс)' + pattern_hr, text, 'h_reflex', re.S)
		l_ST_hr = parse(r'H-рефлекс.*?лев\., Soleus, Tibialis.*?(H-рефлекс)' + pattern_hr, text, 'h_reflex', re.S)

		pnp_diagnosis = parse(r'((ЭНМГ данные|ЭНМГданные)\s*(не исключают).+?(полинейропатию|полинейропатии))|(ЭНМГ признаки.+?(полинейропатии))|((ЭНМГ паттерн).+?(не исключает|характерен).+?(полинейропатии))|((Выявлены).+?(полинейропатии))', text, 'diagnosis_true', re.S, end=True)

		pnp_diagnosis_false = parse(r'(Не выявлены).+?(ЭНМГ признаки|признаки|критерии).+?(полинейропатии|полинейропатических нарушений)', text, 'diagnosis_false', re.S)

		if pnp_diagnosis_false == '1':
			pnp_diagnosis = '0'


		save_data(text, 'files/mainfile.txt', 'a', 'utf-8')
		# if hands_legs == 'legs;':
		# 	create_conclusion_file(text, last_name, name)


		summary_file.write(str(count) + ';'+

			age +
			last_name +
			name +
			gender +

			rAhT1 +
			rAhT2 + 
			lAhT1 +
			lAhT2 + 

			r_EdbP1 +
			r_EdbP2 +
			l_EdbP1 +
			l_EdbP2 +

			r_S + 
			l_S +

			r_Ps +
			l_Ps +

			f_wave + 
			rAhT_f_wave +
			lAhT_f_wave +
			r_EdbP_f_wave +
			l_EdbP_f_wave +

			h_reflex + 
			r_ST_hr + 
			l_ST_hr +

			pnp_diagnosis +

			'\n')

		print('Парсинг файлов: ' + str(count) + '/' + str(number_docs))
		count += 1

	summary_file.close()


# Сохранить все данные в один текстовый файл
def save_data(object, filename, mode, encoding):
	with open(filename, mode, encoding=encoding) as file:
		file.write(object)
	return

#Выбираем заключения
def create_conclusion_file(text, last_name, name):
	conclusion = text.replace('\n\n', '\n')
	conclusion = conclusion.split('Врач функциональной диагностики')
	conclusion = conclusion[0].split('\n')
	len_conslusion = len(conclusion)
	conclusion = conclusion[len_conslusion-6:len_conslusion]

	# Сохранить все данные в один текстовый файл
	with open('files/conclusion.txt', 'a', encoding='utf-8') as conclusion_file:
		conclusion_file.write(last_name[:-1] + ' ' + name[:-1] + '\n')
		for i in conclusion:
			conclusion_file.write(i + '\n')


if __name__ == '__main__':
	main()