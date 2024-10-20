from flask import Flask, render_template, request, send_file
import requests
from bs4 import BeautifulSoup
from fake_useragent import UserAgent
import xlsxwriter

app = Flask(__name__)

# Функция для парсинга результатов поиска в Google
def google_search(query):
    global results
    ua = UserAgent()  # Создание случайного User-Agent
    header = {"User-Agent": 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36 OPR/114.0.0.0'}  # Добавление заголовков с фейковым User-Agent
    url = f"https://www.google.com/search?q={query}"
    
    response = requests.get(url, headers=header)

    results =[]

    if response.status_code == 200:
        soup = BeautifulSoup(response.text, "html.parser")
        for result in soup.select('.tF2Cxc'):
            title = result.select_one('.DKV0Md').text
            description = result.select_one('.hJNv6b').text
            name = result.select_one('.VuuXrf').text
            link = result.select_one('.yuRUbf a')['href']
            results.append({"title": title, "description":description, "name":name, "link": link})
        return results
    return []

# Главная страница с формой поиска
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        query = request.form.get("query")
        results = google_search(query)  # Вызов функции парсинга
        return render_template("search.html", results=results)
    return render_template("search.html", results=None)

@app.route("/download_xlsx")
def download_xlsx():
    global results
    # Создаём XLSX файл
    xlsx_filename = "results.xlsx"
    
    # Создаем книгу и добавляем лист
    workbook = xlsxwriter.Workbook(xlsx_filename)
    worksheet = workbook.add_worksheet()
    
    # Записываем заголовки
    headers = ["Title", "Short Description", "Web name", "Link"]
    worksheet.write_row(0, 0, headers)  # Запись заголовков в первую строку

    # Записываем результаты
    for row_num, result in enumerate(results, start=1):  # Начинаем с 1, чтобы не перезаписать заголовки
        worksheet.write(row_num, 0, result['title'])  # Заголовок
        worksheet.write(row_num, 1, result['description'])  # Краткое описание
        worksheet.write(row_num, 2, result['name']) # Название сайта 
        worksheet.write(row_num, 3, result['link'])  # Ссылка на сайт

    workbook.close()  # Закрываем книгу

    return send_file(xlsx_filename, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)



