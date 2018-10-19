import re
import os, os.path

# Custom
import docx
import sound

path_polyneuropathy_files = 'files/pnp'
number_docs = len([name for name in os.listdir(path_polyneuropathy_files) if os.path.isfile(os.path.join(path_polyneuropathy_files, name))])
files = os.listdir(path_polyneuropathy_files)

def main():
	create_csv()
	sound.play()

def parse(regular_expression, text, category, end=False):
	search = re.search(regular_expression, text, re.S)
	result = ''

	if search and category == 'age':
		result = search.group(1) + ';'
	elif not search and category == 'age':
		result = ';'

	if search and category == 'srv_motor_7':
		result = search.group(2) + ';' + search.group(3) + ';'
	elif not search and category == 'srv_motor_7':
		result = ';;'
	elif search and category == 'srv_motor_9':
		result = search.group(2) + ';' + search.group(3) + ';' + search.group(10) + ';'	
	elif not search and category == 'srv_motor_9':
		result = ';;;'

	if search and category == 'srv_sensor_9':
		result = search.group(3) + ';' + search.group(10) + ';'
	elif not search and category == 'srv_sensor_9':
		result = ';;'

	if search and category == 'f_wave_bool':
		result = '1;'
	elif not search and category == 'f_wave_bool':
		result = '0;'

	if search and category == 'f_wave':
		result = search.group(2) + ';'
	elif not search and category == 'f_wave':
		result = ';'

	if search and category == 'h_reflex_bool':
		result = '1;'
	elif not search and category == 'h_reflex_bool':
		result = '0;'

	if search and category == 'h_reflex':
		result = search.group(4) + ';' + search.group(7) + ';'
	elif not search and category == 'h_reflex':
		result = ';;'

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

	summary_file = open('files/summary.csv', 'w', encoding='utf-8')
	summary_file.write(
		'id;' + 
		'age;' +
		'last_name;'
		'name;' +
		'gender;' +
		'hands/legs;' +
		
		#Ноги
		#СРВ моторная
		# Abductor hallucis, Tibialis
		'rAhTt_lat;' +
		'rAhTt_ampl;' +
		'rAhTp_lat;' +
		'rAhTp_ampl;' +
		'rAhTp_speed;' +

		'lAhTt_lat;' +
		'lAhTt_ampl;' +
		'lAhTp_lat;' +
		'lAhTp_ampl;' +
		'lAhTp_speed;' +

		# Extensor digitorum brevis, Peroneus
		'r_EdbPt_lat;' +
		'r_EdbPt_ampl;' +
		'r_EdbPf_lat;' +
		'r_EdbPf_ampl;' +
		'r_EdbPf_speed;' +

		'l_EdbPt_lat;' +
		'l_EdbPt_ampl;' +
		'l_EdbPf_lat;' +
		'l_EdbPf_ampl;' +
		'l_EdbPf_speed;' +


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

		# #РУКИ
		# # СРВ моторная
		# # Abductor digiti minimi, Ulnaris
		# 'rAdmUw_lat;' +
		# 'rAdmUw_ampl;' +
		# 'rAdmUu_lat;' +
		# 'rAdmUu_ampl;' +
		# 'rAdmUu_speed;' +
		# 'rAdmU_f_wave;' +

		# 'lAdmUw_lat;' +
		# 'lAdmUw_ampl;' +
		# 'lAdmUu_lat;' +
		# 'lAdmUu_ampl;' +
		# 'lAdmUu_speed;' +
		# 'lAdmU_f_wave;' +

		# # Abductor pollicis brevis, Medianus
		# 'rApbMw_lat;' +
		# 'rApbMw_ampl;' +
		# 'rApbMu_lat;' +
		# 'rApbMu_ampl;' +
		# 'rApbMu_speed;' +
		# 'rApbM_f_wave;' +

		# 'lApbMw_lat;' +
		# 'lApbMw_ampl;' +
		# 'lApbMu_lat;' +
		# 'lApbMu_ampl;' +
		# 'lApbMu_speed;' +
		# 'lApbM_f_wave;' +

		# # СРВ сенсорная
		# # n. Medianus
		# 'r_M_ampl;' +
		# 'r_M_speed;' +
		# 'l_M_ampl;' +
		# 'l_M_speed;' +

		# # Ulnaris
		# 'r_U_ampl;' +
		# 'r_U_speed;' +
		# 'l_U_ampl;' +
		# 'l_U_speed;' +

		'\n')

	for i in files:
		text = docx.get_docx_text('./' + path_polyneuropathy_files + '/' + i)

		age = parse(r'Пациент:.+?(\d+)', text, 'age')
		last_name = parse(r'Пациент:\s*(\w+)\s*(\w+)\s*(\w+)', text, 'last_name')
		name = parse(r'Пациент:\s*(\w+)\s*(\w+)\s*(\w+)', text, 'name')
		gender = parse(r'', text, 'gender')

		hands_legs = parse(r'(Abductor pollicis brevis, Medianus, C8 T1)|(n. Medianus)|(Abductor digiti minimi, Ulnaris, C8 T1)|(n. Ulnaris)', text, 'hands')

		#Ноги
		rAhTt = parse(r'СРВ моторная.*?пр\., Abductor hallucis, Tibialis.*?(медиальная лодыжка|предплюсна)' + pattern_srv7, text, 'srv_motor_7')
		rAhTp = parse(r'СРВ моторная.*?пр\., Abductor hallucis, Tibialis.*?(подколенная ямка)' + pattern_srv9, text, 'srv_motor_9')
		lAhTt = parse(r'СРВ моторная.*?лев\., Abductor hallucis, Tibialis.*?(медиальная лодыжка|предплюсна)' + pattern_srv7, text, 'srv_motor_7')
		lAhTp = parse(r'СРВ моторная.*?лев\., Abductor hallucis, Tibialis.*?(подколенная ямка)' + pattern_srv9, text, 'srv_motor_9')

		r_EdbPt = parse(r'СРВ моторная.*?пр\., Extensor digitorum brevis, Peroneus.*?(предплюсна)' + pattern_srv7, text, 'srv_motor_7')
		r_EdbPf = parse(r'СРВ моторная.*?пр\., Extensor digitorum brevis, Peroneus.*?(головка малоберцовой кости)' + pattern_srv9, text, 'srv_motor_9')
		l_EdbPt = parse(r'СРВ моторная.*?лев\., Extensor digitorum brevis, Peroneus.*?(предплюсна)' + pattern_srv7, text, 'srv_motor_7')
		l_EdbPf = parse(r'СРВ моторная.*?лев\., Extensor digitorum brevis, Peroneus.*?(головка малоберцовой кости)' + pattern_srv9, text, 'srv_motor_9')

		r_S = parse(r'СРВ сенсорная.*?пр., n.Suralis.*?(1|Ср. треть голени|нижняя треть голени)' + pattern_srv9, text, 'srv_sensor_9')
		l_S = parse(r'СРВ сенсорная.*?лев., n.Suralis.*?(1|Ср. треть голени|нижняя треть голени)' + pattern_srv9, text, 'srv_sensor_9')

		r_Ps = parse(r'СРВ сенсорная.*?пр., n.Peroneus superficialis.*?(1|Ср. треть голени)' + pattern_srv9, text, 'srv_sensor_9')
		l_Ps = parse(r'СРВ сенсорная.*?лев., n.Peroneus superficialis.*?(1|Ср. треть голени)' + pattern_srv9, text, 'srv_sensor_9')


		f_wave = parse(r'Параметры F-волны', text, 'f_wave_bool')
		rAhT_f_wave = parse(r'Параметры F-волны.*?пр\., Abductor hallucis, Tibialis.*?\n+(\d+)' + pattern_f_wave, text, 'f_wave')
		lAhT_f_wave = parse(r'Параметры F-волны.*?лев\., Abductor hallucis, Tibialis.*?\n+(\d+)' + pattern_f_wave, text, 'f_wave')
		r_EdbP_f_wave = parse(r'Параметры F-волны.*?пр\., Extensor digitorum brevis, Peroneus.*?\n+(\d+)' + pattern_f_wave, text, 'f_wave')
		l_EdbP_f_wave = parse(r'Параметры F-волны.*?лев\., Extensor digitorum brevis, Peroneus.*?\n+(\d+)' + pattern_f_wave, text, 'f_wave')

		h_reflex = parse(r'H-рефлекс', text, 'h_reflex_bool')
		r_ST_hr = parse(r'H-рефлекс.*?пр., Soleus, Tibialis.*?(H-рефлекс)' + pattern_hr, text, 'h_reflex')
		l_ST_hr = parse(r'H-рефлекс.*?лев., Soleus, Tibialis.*?(H-рефлекс)' + pattern_hr, text, 'h_reflex', end=True)


		#Выбираем заключения
		conclusion = text.replace('\n\n', '\n')
		conclusion = conclusion.split('Врач функциональной диагностики')
		conclusion = conclusion[0].split('\n')
		len_conslusion = len(conclusion)
		conclusion = conclusion[len_conslusion-6:len_conslusion]

		# Сохранить все данные в один текстовый файл
		# print(hands_legs)
		if hands_legs == 'legs;':
			with open('files/conclusion.txt', 'a', encoding='utf-8') as conclusion_file:
				conclusion_file.write(last_name[:-1] + ' ' + name[:-1] + '\n')
				for i in conclusion:
					conclusion_file.write(i + '\n')

		# #Руки
		# rAdmUw = parse(r'СРВ моторная.*?пр\., Abductor digiti minimi, Ulnaris.*?(запястье)' + pattern_srv7, text, 'srv_motor_7')
		# rAdmUu = parse(r'СРВ моторная.*?пр\., Abductor digiti minimi, Ulnaris.*?(локтевой сгиб)' + pattern_srv9, text, 'srv_motor_9')
		# rAdmU_f_wave = parse(r'Параметры F-волны.*?пр\., Abductor digiti minimi, Ulnaris.*?\n+(\d+)' + pattern_f_wave, text, 'f_wave')
		# lAdmUw = parse(r'СРВ моторная.*?лев\., Abductor digiti minimi, Ulnaris.*?(запястье)' + pattern_srv7, text, 'srv_motor_7')
		# lAdmUu = parse(r'СРВ моторная.*?лев\., Abductor digiti minimi, Ulnaris.*?(локтевой сгиб)' + pattern_srv9, text, 'srv_motor_9')
		# lAdmU_f_wave = parse(r'Параметры F-волны.*?лев\., Abductor digiti minimi, Ulnaris.*?\n+(\d+)' + pattern_f_wave, text, 'f_wave')

		# rApbMw = parse(r'СРВ моторная.*?пр\., Abductor pollicis brevis, Medianus.*?(запястье)' + pattern_srv7, text, 'srv_motor_7')
		# rApbMu = parse(r'СРВ моторная.*?пр\., Abductor pollicis brevis, Medianus.*?(локтевой сгиб)' + pattern_srv9, text, 'srv_motor_9')
		# rApbM_f_wave = parse(r'Параметры F-волны.*?пр\., Abductor pollicis brevis, Medianus.*?\n+(\d+)' + pattern_f_wave, text, 'f_wave')
		# lApbMw = parse(r'СРВ моторная.*?лев\., Abductor pollicis brevis, Medianus.*?(запястье)' + pattern_srv7, text, 'srv_motor_7')
		# lApbMu = parse(r'СРВ моторная.*?лев\., Abductor pollicis brevis, Medianus.*?(локтевой сгиб)' + pattern_srv9, text, 'srv_motor_9')
		# lApbM_f_wave = parse(r'Параметры F-волны.*?лев\., Abductor pollicis brevis, Medianus.*?\n+(\d+)' + pattern_f_wave, text, 'f_wave')


		# r_M = parse(r'СРВ сенсорная.*?пр., n. Medianus.*?(локтевой сгиб)' + pattern_srv9, text, 'srv_sensor_9')
		# l_M = parse(r'СРВ сенсорная.*?лев., n. Medianus.*?(локтевой сгиб)' + pattern_srv9, text, 'srv_sensor_9')

		# r_U = parse(r'СРВ сенсорная.*?пр., n. Ulnaris.*?(запястье)' + pattern_srv9, text, 'srv_sensor_9')
		# l_U = parse(r'СРВ сенсорная.*?лев., n. Ulnaris.*?(запястье)' + pattern_srv9, text, 'srv_sensor_9')

		#Сохранить все данные в один текстовый файл
		# with open('mainfile.txt', 'a', encoding='utf-8') as main_file:
		# 	main_file.write(text)

		summary_file.write(str(count) + ';'+

			age +
			last_name +
			name +
			gender +
			hands_legs +

			rAhTt +
			rAhTp + 
			lAhTt +
			lAhTp + 

			r_EdbPt +
			r_EdbPf +
			l_EdbPt +
			l_EdbPf +

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

			# rAdmUw +
			# rAdmUu +
			# rAdmU_f_wave +
			# lAdmUw +
			# lAdmUu +
			# lAdmU_f_wave +

			# rApbMw +
			# rApbMu +
			# rApbM_f_wave +
			# lApbMw +
			# lApbMu +
			# lApbM_f_wave +

			# r_M +
			# l_M +
			# r_U +
			# l_U +

			'\n')

		print('Парсинг файлов: ' + str(count) + '/' + str(number_docs))
		count += 1


	summary_file.close()

if __name__ == '__main__':
	main()