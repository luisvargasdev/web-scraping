from pathlib import Path
from bs4 import BeautifulSoup
import requests, openpyxl

# create the excel workbook
excel = openpyxl.Workbook()
# use the active sheet (in case there is more than one sheet)
sheet = excel.active
# change the title of that sheet
sheet.title = 'Top Rated Movies'
# create column names for every attribute
sheet.append(['Movie Rank', 'Movie Name', 'Year of Release', 'IMDB Rating'])

headers = {
    "User-Agent": "Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:109.0) Gecko/20100101 Firefox/117.0"
}

try:
    # webpage to get the data
    source = requests.get('https://www.imdb.com/chart/top/', headers=headers)
    # raise an error for the url
    source.raise_for_status()

    # accessing the html text and parser it
    soup = BeautifulSoup(source.text,'html.parser')

    # search movies // find fetch the first match
    movies = soup.find('ul', class_="ipc-metadata-list").find_all('li')
    
    # loop to get every attribute necessary for the scraping
    for movie in movies:
        # use the find function to access the tag and the text require
        name = movie.find('div', class_="ipc-title").a.text.split(". ")[1]
        rank = movie.find('div', class_="ipc-title").a.text.split(". ")[0]
        year = movie.find('div', class_="sc-b0691f29-7").span.text
        rating = movie.find('span', class_="sc-b0691f29-1").span.text.split("(")[0]

        # se imprimen en consola para revision
        print(rank, name, year, rating)

        # se guardan en el archivo de excel activo
        sheet.append([rank, name, year, rating])

except Exception as e:
    print(e)

# simple verification to save the file only if it does not exist
if(not (Path("IMDB Movie Ratings.xlsx").exists())):
    excel.save('IMDB Movie Ratings.xlsx')