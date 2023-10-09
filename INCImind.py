from collections import Counter
import io
import os
import re
import math
import queue
import threading
import pdfplumber
import pytesseract
import configparser
from sys import exit
from PIL import Image
from fuzzywuzzy import fuzz
from openpyxl import load_workbook
from pdf2image import convert_from_path
pytesseract.pytesseract.tesseract_cmd = ".\\Tesseract-OCR\\tesseract.exe"

def search_files(directory, filename, exact=False):
  res = []
  for root, _, files in os.walk(directory):
    if exact:
      if (filename + ".xlsx").lower() in (file.lower() for file in files):
        return os.path.join(root, filename + ".xlsx")
    else:
      for file in files:
        if fuzz.partial_ratio(filename.lower(), file.lower()) > 85:
          res.append(os.path.join(root, file))
          
  return res

def choose_file_menu(ingredient_file_paths, ingredient):
  print()
  menu = []
  for i, path in enumerate(ingredient_file_paths):
    menu.append(path)
    print(str(i+1) + " -> " + (os.path.join(os.path.basename(os.path.dirname(path)), os.path.basename(path))).replace("\\", " > "))

  print("0 -> Nessuno dei file precedenti")
  res = input("\nSelezione file per '" + ingredient + "': ")
  
  if int(res) > 0:
    print("File scelto per " + ingredient + ": " + os.path.join(os.path.basename(os.path.dirname(menu[int(res)-1])), os.path.basename(menu[int(res)-1])) + "\n")
    return menu[int(res)-1]
  return None

def read_xlsx(excel_file_path):
  workbook = load_workbook(excel_file_path)
  worksheet = workbook.active  

  result = []
  for row in worksheet.iter_rows(values_only=True, min_row=5):
    if not row[0]:
      break
    for i in range(2):
        if not i % 2:
          tmp = row[i]
        else:
          result.append((tmp, row[i]))
  return result

def check_in_file(ingredients, file_text, result_queue):
  res = []
  for ingredient in ingredients:
    pattern = re.compile(r"\b{}\b".format(re.escape(ingredient)), re.IGNORECASE)
    matches = pattern.findall(file_text)
    if matches:
      res.append((matches[0], len(matches)))  

  result_queue.put(res)

def result_processing(results):
  res = []

  for ingredient_x, rep_x in results:
    is_duplicate = False
    for ingredient_y, rep_y in results:
      if ingredient_x != ingredient_y and rep_x == rep_y:
        words = ingredient_y.split()
        if ingredient_x in words:
          is_duplicate = True
          break
    if not is_duplicate:
      res.append(ingredient_x)

  return res

if __name__ == "__main__":
  config = configparser.ConfigParser()
  config.read('./config.ini')
  tec_sheets_path = config['DEFAULT']['tec_sheets_path']

  if not tec_sheets_path:
    print("Errore: valore di configurazione 'tec_sheets_path' non trovato")
    input("Premere invio per continuare...")
    exit()

  filename_to_find = input("Nome del file Excel: ")
  file_path = search_files(os.path.expanduser("~"), filename_to_find, True)
  if not file_path:
    print("Errore: file '" + filename_to_find + "' non trovato")
    input("Premere invio per continuare...")
    exit()
    
  formula = read_xlsx(file_path)

  inci = []
  for ingredient in formula:
    if ingredient[0].startswith("-"):
      inci.append((ingredient[0], ["ignorato"]))
      continue
    ingredient_file_paths = search_files(tec_sheets_path, ingredient[0])

    if not ingredient_file_paths:
      inci.append((ingredient[0], ["ignorato"]))
      print("\nNessun file trovato per: " + ingredient[0])
      input("Premere invio per continuare")
      continue

    ingredient_file_path = choose_file_menu(ingredient_file_paths, ingredient[0])

    if not ingredient_file_path:
      print("\nFile ingoranto per: " + ingredient[0])
      inci.append((ingredient[0], ["ignorato"]))
      continue

    with pdfplumber.open(ingredient_file_path) as pdf:
      file_text = ' '.join([page.extract_text() for page in pdf.pages])

      if not file_text.strip():
        pages = convert_from_path(ingredient_file_path, poppler_path=".\\poppler-23.08.0\\Library\\bin")
        for page in pages:
          output = io.BytesIO()
          page.save(output, format='jpeg')
          output.seek(0)
          img = Image.open(output)
          file_text += pytesseract.image_to_string(img, lang='eng')

    with open("./ingredients.txt", "r") as f:
      ingredients = f.read().splitlines()

    lines_per_thread = 5000
    threads_num = math.ceil(len(ingredients) / lines_per_thread)
    threads = []
    result_queue = queue.Queue()
    for i in range(threads_num):
      thread = threading.Thread(target=check_in_file, args=(ingredients[i * lines_per_thread: (i+1) * lines_per_thread], file_text, result_queue))
      threads.append(thread)
      thread.start()

    for thread in threads:
      thread.join()

    result_queue_list = []
    while not result_queue.empty():
      result_queue_list.extend(result_queue.get())

    res = result_processing(result_queue_list)
    inci.append((ingredient[0], res))    

  for record in inci:
    print("(" + record[0] + ")")
    print(*[x.capitalize() for x in record[1]], sep=", ")
    print()

  merged_list = [item.lower().capitalize() for sublist in [ing[1] for ing in inci] for item in sublist if item != "ignorato"]
  duplicates = [item for item, count in Counter(merged_list).items() if count > 1]

  colors = ['\033[31m', '\033[32m', '\033[33m', '\033[34m', '\033[35m', '\033[36m']  
  color_idx = 0

  for item in merged_list:
    if item in duplicates:
      print(f'{colors[color_idx % len(colors)]}{item}\033[0m', end=' ')
      color_idx += 1
    else:
      print(item, end=', ')
