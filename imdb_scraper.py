from bs4 import BeautifulSoup
import requests
import openpyxl

excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = 'IMDb top 250'
sheet.append(['Rank', 'Movie Name', 'Release Year', 'IMDb rating'])

url = 'https://www.imdb.com/chart/top/'
page = requests.get(url).text
soup = BeautifulSoup(page, 'lxml')

table = soup.find('tbody', class_='lister-list')
movies = table.find_all('tr')

for movie in movies:
    rank = movie.find('td', class_='titleColumn').get_text(strip=True).split('.')[0]
    name = movie.find('td', class_='titleColumn').a.text
    year = movie.find('td', class_='titleColumn').span.text.strip('()')
    rating = movie.find('td', class_='ratingColumn imdbRating').strong.text

    print(rank, name, year, rating)
    sheet.append([rank, name, year, rating])

excel.save('IMDb ratings.xlsx')
