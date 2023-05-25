# import bs4
from bs4 import BeautifulSoup
import requests, openpyxl


excel = openpyxl.Workbook()
# print(excel.sheetnames)
sheet = excel.active
sheet.title = 'Top rated movies'
# print(print(excel.sheetnames))
sheet.append(['Movie Rank', 'Movie Name', 'Year of Release', 'Rating'])

try:
    source = requests.get('https://www.imdb.com/chart/top/')    # response is saved in source variable
    # print(source.text)
    # print(source.status_code)           # status code 
    # print(source.raise_for_status)      # raise to get an alert for incorrect urls

    soup = BeautifulSoup(source.text , 'html.parser')

    movies = soup.find('tbody' , class_ = "lister-list").find_all('tr')

    for movie in movies:
        name = movie.find('td', class_ = "titleColumn").a.text
        rank = movie.find('td', class_ = "titleColumn").get_text(strip=True).split('.')[0]
        year = movie.find('td', class_ = "titleColumn").span.text.strip('()')
        rating = movie.find('td', class_ = "ratingColumn imdbRating").strong.text

        # print(name, rank, year, rating)
        sheet.append([rank, name, year, rating])

except Exception as e:
    print(e)

excel.save('IMDB_MovieRating.xlsx')




