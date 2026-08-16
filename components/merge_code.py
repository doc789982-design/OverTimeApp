import os

# Папка, где лежат твои .qml файлы ('.' если в текущей)
folder_path = '.' 
# Сохраняем в .txt, чтобы было понятно, что это файл для чтения, а не для запуска
output_file = 'merged_qml_for_ai.txt' 

with open(output_file, 'w', encoding='utf-8') as outfile:
    for filename in os.listdir(folder_path):
        # Ищем только QML файлы
        if filename.endswith('.qml'):
            file_path = os.path.join(folder_path, filename)
            
            with open(file_path, 'r', encoding='utf-8') as infile:
                # В QML комментарии пишутся через //, используем их для красоты
                outfile.write(f"\n\n// {'='*40}\n")
                outfile.write(f"// НАЧАЛО ФАЙЛА: {filename}\n")
                outfile.write(f"// {'='*40}\n\n")
                
                outfile.write(infile.read())

print(f"Все QML компоненты успешно объединены в {output_file}")