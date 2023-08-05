from pathlib import Path
import csv
import sys
import os
import shutil


def make_wordfile(replace_string_csvfile, original_wordfile_path, processed_wordfile_common_name):
    #replace_string_csvfile = '**.csv'
    #original_wordfile_path = '**.docx'
    #processed_wordfile_common_name = '**'

    replace_string_lists = []

    with open(replace_string_csvfile, encoding = 'utf-8-sig') as f:
        reader = csv.reader(f)
        for row in reader :
            replace_string_lists.append(row)


    wildcard = replace_string_lists.pop(0)
    shutil.unpack_archive(original_wordfile_path,format='zip', extract_dir='archive-dir')

    #置換スクリプト
    for replace_string_list in replace_string_lists:
        dir_path = Path('ProcessedWordFile/' + processed_wordfile_common_name + ' ' + replace_string_list[0])

        shutil.copytree('./archive-dir', dir_path)

        with open(Path(dir_path.__str__() + '/word/document.xml'), mode='r', encoding='UTF-8') as doc_r:
            file_content = doc_r.read()
            for i in range(len(replace_string_list)):
                file_content = file_content.replace(wildcard[i], replace_string_list[i])
            with open(Path(dir_path.__str__() + '/word/document.xml'), mode='w', encoding='UTF-8') as doc_w:
                    doc_w.write(file_content)

        shutil.make_archive(Path(dir_path.__str__()+'.docx'), format='zip', root_dir=dir_path)


args = sys.argv

if not args[1] and  not args[2] and args[3]:
    print("Arguments are missing.")
    exit()

os.makedirs("./ProcessedWordFile", exist_ok=True)

try:
    make_wordfile(args[1],args[2],args[3])
except FileExistsError as e:
    print("error!")
    print(e)

shutil.rmtree(Path('./archive-dir'))
for f in os.listdir('./ProcessedWordFile'):
    if os.path.isdir(os.path.join('./ProcessedWordFile', f)):
        shutil.rmtree(os.path.join('./ProcessedWordFile', f))
    if os.path.isfile(os.path.join('./ProcessedWordFile', f)):
        try:
            os.rename(os.path.join('./ProcessedWordFile', f),os.path.join('./ProcessedWordFile', f).__str__().replace('.zip', ''))
        except FileExistsError as e:
            print("error!")
            print(e)
            os.remove(os.path.join('./ProcessedWordFile', f))
