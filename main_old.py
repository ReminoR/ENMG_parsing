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

def parse(regular_expression, text, number_of_groups):
	search = re.search(regular_expression, text, re.S)
	result = ''

	if search:
		for i in range(2, number_of_groups + 1):
			result = result + search.group(i)	+ ';'
		return result
	else:
		for i in range(number_of_groups):
			result = result + ';'
		return result

def create_csv():
	count = 1
	pattern_7 = '\n+([\d,]+)\n+([\d,]+)\n+([\d,]+)\n+([\d,]+)\n+([\d,]+)\n+([\d,]+)\n+([\d,]+)'
	pattern_9 = '\n+([\d,]+)\n+([\d,]+)\n+([\d,]+)\n+([\d,]+)\n+([\d,]+)\n+([\d,]+)\n+([\d,]+)\n+([\d,]+)\n+([\d,]+)'

	summary_file = open('files/summary.csv', 'w', encoding='utf-8')
	summary_file.write(
		'id;' + 

		# пр., Abductor hallucis, Tibialis, предплюсна
		'rAhTt_lat;' +
		'rAhTt_ampl;' +
		'rAhTt_dur;' +
		'rAhTt_sq;' +
		'rAhTt_stim1;' +
		'rAhTt_stim2;' +
		'rAhTt_dist;' +

		# пр., Abductor hallucis, Tibialis, подколенная ямка
		'rAhTt_lat;' +
		'rAhTt_ampl;' +
		'rAhTt_dur;' +
		'rAhTt_sq;' +
		'rAhTt_stim1;' +
		'rAhTt_stim2;' +
		'rAhTt_dist;' +
		'rAhTt_time;' +
		'rAhTt_speed;' +

		# лев., Abductor hallucis, Tibialis, предплюсна
		'lAhTt_lat;' +
		'lAhTt_ampl;' +
		'lAhTt_dur;' +
		'lAhTt_sq;' +
		'lAhTt_stim1;' +
		'lAhTt_stim2;' +
		'lAhTt_dist;' +

		# лев., Abductor hallucis, Tibialis, подколенная ямка
		'lAhTt_lat;' +
		'lAhTt_ampl;' +
		'lAhTt_dur;' +
		'lAhTt_sq;' +
		'lAhTt_stim1;' +
		'lAhTt_stim2;' +
		'lAhTt_dist;' +
		'lAhTt_time;' +
		'lAhTt_speed;' +

		# # пр., Extensor digitorum brevis, Peroneus, предплюсна
		# 'r_Extensor_digitorum_brevis_Peroneus_tarsus_lat;' +
		# 'r_Extensor_digitorum_brevis_Peroneus_tarsus_ampl;' +
		# 'r_Extensor_digitorum_brevis_Peroneus_tarsus_dur;' +
		# 'r_Extensor_digitorum_brevis_Peroneus_tarsus_sq;' +
		# 'r_Extensor_digitorum_brevis_Peroneus_tarsus_stim1;' +
		# 'r_Extensor_digitorum_brevis_Peroneus_tarsus_stim2;' +
		# 'r_Extensor_digitorum_brevis_Peroneus_tarsus_dist;' +

		# # пр., Extensor digitorum brevis, Peroneus, головка малоберцовой кости
		# 'r_Extensor_digitorum_brevis_Peroneus_fibula_lat;' +
		# 'r_Extensor_digitorum_brevis_Peroneus_fibula_ampl;' +
		# 'r_Extensor_digitorum_brevis_Peroneus_fibula_dur;' +
		# 'r_Extensor_digitorum_brevis_Peroneus_fibula_sq;' +
		# 'r_Extensor_digitorum_brevis_Peroneus_fibula_stim1;' +
		# 'r_Extensor_digitorum_brevis_Peroneus_fibula_stim2;' +
		# 'r_Extensor_digitorum_brevis_Peroneus_fibula_dist;' +
		# 'r_Extensor_digitorum_brevis_Peroneus_fibula_time;' +
		# 'r_Extensor_digitorum_brevis_Peroneus_fibula_speed;' +

		# # пр., Extensor digitorum brevis, Peroneus, подколенная ямка
		# 'r_Extensor_digitorum_brevis_Peroneus_popfos_lat;' +
		# 'r_Extensor_digitorum_brevis_Peroneus_popfos_ampl;' +
		# 'r_Extensor_digitorum_brevis_Peroneus_popfos_dur;' +
		# 'r_Extensor_digitorum_brevis_Peroneus_popfos_sq;' +
		# 'r_Extensor_digitorum_brevis_Peroneus_popfos_stim1;' +
		# 'r_Extensor_digitorum_brevis_Peroneus_popfos_stim2;' +
		# 'r_Extensor_digitorum_brevis_Peroneus_popfos_dist;' +
		# 'r_Extensor_digitorum_brevis_Peroneus_popfos_time;' +
		# 'r_Extensor_digitorum_brevis_Peroneus_popfos_speed;' +

		# # лев., Extensor digitorum brevis, Peroneus, предплюсна
		# 'l_Extensor_digitorum_brevis_Peroneus_tarsus_lat;' +
		# 'l_Extensor_digitorum_brevis_Peroneus_tarsus_ampl;' +
		# 'l_Extensor_digitorum_brevis_Peroneus_tarsus_dur;' +
		# 'l_Extensor_digitorum_brevis_Peroneus_tarsus_sq;' +
		# 'l_Extensor_digitorum_brevis_Peroneus_tarsus_stim1;' +
		# 'l_Extensor_digitorum_brevis_Peroneus_tarsus_stim2;' +
		# 'l_Extensor_digitorum_brevis_Peroneus_tarsus_dist;' +

		# # лев., Extensor digitorum brevis, Peroneus, головка малоберцовой кости
		# 'l_Extensor_digitorum_brevis_Peroneus_fibula_lat;' +
		# 'l_Extensor_digitorum_brevis_Peroneus_fibula_ampl;' +
		# 'l_Extensor_digitorum_brevis_Peroneus_fibula_dur;' +
		# 'l_Extensor_digitorum_brevis_Peroneus_fibula_sq;' +
		# 'l_Extensor_digitorum_brevis_Peroneus_fibula_stim1;' +
		# 'l_Extensor_digitorum_brevis_Peroneus_fibula_stim2;' +
		# 'l_Extensor_digitorum_brevis_Peroneus_fibula_dist;' +
		# 'l_Extensor_digitorum_brevis_Peroneus_fibula_time;' +
		# 'l_Extensor_digitorum_brevis_Peroneus_fibula_speed;' +

		# # лев., Extensor digitorum brevis, Peroneus, подколенная ямка
		# 'l_Extensor_digitorum_brevis_Peroneus_popfos_lat;' +
		# 'l_Extensor_digitorum_brevis_Peroneus_popfos_ampl;' +
		# 'l_Extensor_digitorum_brevis_Peroneus_popfos_dur;' +
		# 'l_Extensor_digitorum_brevis_Peroneus_popfos_sq;' +
		# 'l_Extensor_digitorum_brevis_Peroneus_popfos_stim1;' +
		# 'l_Extensor_digitorum_brevis_Peroneus_popfos_stim2;' +
		# 'l_Extensor_digitorum_brevis_Peroneus_popfos_dist;' +
		# 'l_Extensor_digitorum_brevis_Peroneus_popfos_time;' +
		# 'l_Extensor_digitorum_brevis_Peroneus_popfos_speed;' +

		# # пр., n.Suralis
		# 'r_Suralis_1_lat;' +
		# 'r_Suralis_1_ampl;' +
		# 'r_Suralis_1_dur;' +
		# 'r_Suralis_1_sq;' +
		# 'r_Suralis_1_stim1;' +
		# 'r_Suralis_1_stim2;' +
		# 'r_Suralis_1_dist;' +
		# 'r_Suralis_1_time;' +
		# 'r_Suralis_1_speed;' +

		# # лев., n.Suralis
		# 'l_Suralis_1_lat;' +
		# 'l_Suralis_1_ampl;' +
		# 'l_Suralis_1_dur;' +
		# 'l_Suralis_1_sq;' +
		# 'l_Suralis_1_stim1;' +
		# 'l_Suralis_1_stim2;' +
		# 'l_Suralis_1_dist;' +
		# 'l_Suralis_1_time;' +
		# 'l_Suralis_1_speed;' +

		# # пр., Peroneus superficialis, 1
		# 'r_Peroneus_superficialis_1_lat;' +
		# 'r_Peroneus_superficialis_1_ampl;' +
		# 'r_Peroneus_superficialis_1_dur;' +
		# 'r_Peroneus_superficialis_1_sq;' +
		# 'r_Peroneus_superficialis_1_stim1;' +
		# 'r_Peroneus_superficialis_1_stim2;' +
		# 'r_Peroneus_superficialis_1_dist;' +
		# 'r_Peroneus_superficialis_1_time;' +
		# 'r_Peroneus_superficialis_1_speed;' +

		# # лев., Peroneus superficialis, 1
		# 'l_Peroneus_superficialis_1_lat;' +
		# 'l_Peroneus_superficialis_1_ampl;' +
		# 'l_Peroneus_superficialis_1_dur;' +
		# 'l_Peroneus_superficialis_1_sq;' +
		# 'l_Peroneus_superficialis_1_stim1;' +
		# 'l_Peroneus_superficialis_1_stim2;' +
		# 'l_Peroneus_superficialis_1_dist;' +
		# 'l_Peroneus_superficialis_1_time;' +
		# 'l_Peroneus_superficialis_1_speed;' +



		'\n')

	for i in files:
		text = docx.get_docx_text('./' + path_polyneuropathy_files + '/' + i)

		r_Abductor_hallucis_Tibialis_tarsus = parse(r'СРВ моторная.*?пр\., Abductor hallucis, Tibialis.*?(медиальная лодыжка|предплюсна)' + pattern_7, text, 8)
		r_Abductor_hallucis_Tibialis_popfos = parse(r'СРВ моторная.*?пр\., Abductor hallucis, Tibialis.*?(подколенная ямка)' + pattern_9, text, 10)
		l_Abductor_hallucis_Tibialis_tarsus = parse(r'СРВ моторная.*?лев\., Abductor hallucis, Tibialis.*?(медиальная лодыжка|предплюсна)' + pattern_7, text, 8)
		l_Abductor_hallucis_Tibialis_popfos = parse(r'СРВ моторная.*?лев\., Abductor hallucis, Tibialis.*?(подколенная ямка)' + pattern_9, text, 10)
		
		# r_Extensor_digitorum_brevis_Peroneus_tarsus = parse(r'СРВ моторная.*?пр\., Extensor digitorum brevis, Peroneus.*?(предплюсна)' + pattern_7, text, 8)
		# r_Extensor_digitorum_brevis_Peroneus_fibula = parse(r'СРВ моторная.*?пр\., Extensor digitorum brevis, Peroneus.*?(головка малоберцовой кости)' + pattern_9, text, 10)
		# r_Extensor_digitorum_brevis_Peroneus_popfos = parse(r'СРВ моторная.*?пр\., Extensor digitorum brevis, Peroneus.*?(подколенная ямка)' + pattern_9, text, 10)
		# l_Extensor_digitorum_brevis_Peroneus_tarsus = parse(r'СРВ моторная.*?лев\., Extensor digitorum brevis, Peroneus.*?(предплюсна)' + pattern_7, text, 8)
		# l_Extensor_digitorum_brevis_Peroneus_fibula = parse(r'СРВ моторная.*?лев\., Extensor digitorum brevis, Peroneus.*?(головка малоберцовой кости)' + pattern_9, text, 10)
		# l_Extensor_digitorum_brevis_Peroneus_popfos = parse(r'СРВ моторная.*?лев\., Extensor digitorum brevis, Peroneus.*?(подколенная ямка)' + pattern_9, text, 10)

		# r_Suralis_1 = parse(r'СРВ сенсорная.*?пр\., n.Suralis.*?(1)' + pattern_9, text, 10)
		# l_Suralis_1 = parse(r'СРВ сенсорная.*?лев\., n.Suralis.*?(1)' + pattern_9, text, 10)

		# r_Peroneus_superficialis_1 = parse(r'СРВ сенсорная.*?пр\., n.Peroneus superficialis.*?(1)' + pattern_9, text, 10)
		# l_Peroneus_superficialis_1 = parse(r'СРВ сенсорная.*?лев\., n.Peroneus superficialis.*?(1)' + pattern_9, text, 10)

		with open('mainfile.txt', 'a', encoding='utf-8') as main_file:
			main_file.write(text)

		summary_file.write(str(count) + ';'+
			r_Abductor_hallucis_Tibialis_tarsus +
			r_Abductor_hallucis_Tibialis_popfos + 
			l_Abductor_hallucis_Tibialis_tarsus +
			l_Abductor_hallucis_Tibialis_popfos + 

			# r_Extensor_digitorum_brevis_Peroneus_tarsus +
			# r_Extensor_digitorum_brevis_Peroneus_fibula +
			# r_Extensor_digitorum_brevis_Peroneus_popfos +
			# l_Extensor_digitorum_brevis_Peroneus_tarsus +
			# l_Extensor_digitorum_brevis_Peroneus_fibula +
			# l_Extensor_digitorum_brevis_Peroneus_popfos +

			# r_Suralis_1 + 
			# l_Suralis_1 +

			# r_Peroneus_superficialis_1 +
			# l_Peroneus_superficialis_1 +

			'\n')

		print('Парсинг файлов: ' + str(count))
		count += 1


	summary_file.close()

if __name__ == '__main__':
	main()