import io
import os
import math
import queue
import threading
import pdfplumber
import pytesseract
import configparser
from sys import exit
from PIL import Image
from openpyxl import load_workbook
from pdf2image import convert_from_path
pytesseract.pytesseract.tesseract_cmd = ".\\Tesseract-OCR\\tesseract.exe"

def search_file(directory, filename, extension):
  for root, _, files in os.walk(directory):
    if (filename + "." + extension).lower() in (file.lower() for file in files):
      return os.path.join(root, filename + "." + extension)
    if filename.lower() in (file.lower() for file in files):
      return os.path.join(root, filename)
  return None

def check_in_file(ingredients, file_text, result_queue):
  with open(file, 'r') as file:
    for ingredient in ingredients:
      if ingredient.lower() in file_text.lower():
        result_queue.put(ingredient.lower())

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

if __name__ == "__main__":
  config = configparser.ConfigParser()
  config.read('./config.ini')
  tec_sheets_path = config['DEFAULT']['tec_sheets_path']

  if not tec_sheets_path:
    exit("Error: config 'tec_sheets_path' not found")

  # filename_to_find = "PELLI MATURE+PREZZO+INCI-04.08.23"
  filename_to_find = input("Nome del file Excel:")
  file_path = search_file(os.path.expanduser("~"), filename_to_find, "xlsx")
  if not file_path:
    exit("Error: file '" + filename_to_find + "' not found")

  formula = read_xlsx(file_path)

  inci = []
  for ingredient in formula:
    ingredient_file_path = search_file(tec_sheets_path, ingredient[0], "pdf")

    if not ingredient_file_path:
      exit("Error: ingredient file '" + ingredient[0] + "' not found")

    with pdfplumber.open(ingredient_file_path) as pdf:
      first_page = pdf.pages[1]
      file_text = first_page.extract_text()

      if not file_text:
        pages = convert_from_path(file_path, poppler_path=".\\poppler-23.08.0\\Library\\bin")
        for page in pages:
          output = io.BytesIO()
          page.save(output, format='jpeg')
          output.seek(0)
          img = Image.open(output)
          file_text = pytesseract.image_to_string(img, lang='eng')
          print(file_text)

      else:
        print(file_text)

    with open("./ingredients.txt", "r") as f:
      ingredients = f.readlines()

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

    while not result_queue.empty():
      inci.extend(result_queue.get())

  print(inci, sep=", ")