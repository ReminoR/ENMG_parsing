import os, os.path
import shutil
import re

# Сustom
import docx
import custom_data
import sound

DIR = 'files/'
number_docs = len([name for name in os.listdir(DIR + 'pnp') if os.path.isfile(os.path.join(DIR + 'pnp', name))])
files = os.listdir(DIR + 'pnp')

def main():

	# Поиск по заданным ключевым словам
	folder_name = input('Введите название папки: ')
	create_folder(folder_name)
	custom_search(input('Введите ключевые слова для поиска: '), folder_name)
	sound.play()


	#Поиск по custom_data
	# for i in custom_data.data_nerves:
	# 	folder_name = i
	# 	pattern = i

	# 	create_folder(folder_name)
	# 	custom_search(pattern, folder_name)


def create_folder(folder_name):
	if not os.path.exists(DIR + folder_name):
		os.makedirs(DIR + folder_name)
	else:
		delete_files(folder_name)

def delete_files(folder_name):
	folder = './' + DIR + folder_name

	for the_file in os.listdir(folder):
		file_path = os.path.join(folder, the_file)
		os.unlink(file_path)

def custom_search(pattern, folder_name):
	count_found = 0
	count = 1
	print('Запрос: ' + pattern)

	for i in files:
		text = docx.get_docx_text('./' + DIR + 'pnp/' + i)
		search = re.search(r'' + pattern, text, re.S)
		print('Проверены файлы: ' + str(count) + ' из ' + str(number_docs) + '. ' + i)
		count += 1

		if search:
			count_found += 1
			shutil.copyfile(r'./' + DIR + 'pnp/' + i, r'./files/' + folder_name + '/' + i)

	print('Найдено по запросу: "' + pattern + '" – ' + str(count_found) + ' файлов.\r\n')

if __name__ == '__main__':
	main()