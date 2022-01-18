import requests
from bs4 import BeautifulSoup
import numpy as np
import pandas as pd
import time
import random

headers = {"Accept-Language": "en-US,en;q=0.5"}


def parse():
    it = 1
    # only first 10 pages to start
    movies_info = pd.DataFrame(columns=['Countries', 'Languages'])
    for start in np.arange(1, 500, 50):
        url = f'https://www.imdb.com/search/title/?title_type=feature&genres' \
              f'=crime&sort=num_votes,desc&start={start}&ref_=adv_nxt'
        res = requests.get(url, headers=headers)
        page_html = BeautifulSoup(res.text, "lxml")
        movies = page_html.find_all('div', class_='lister-item')
        movies_with_links = {}
        for movie in movies:
            movie_name = movie.find('h3', class_='lister-item-header').find('a').text
            movies_with_links[movie_name] = movie.find('h3', class_='lister-item-header').find('a')['href']
        for movie in movies_with_links:
            print(f'scrapping data... iteration:{it} left: {451 - it}')
            time.sleep(random.randint(2, 4))
            res = requests.get('https://www.imdb.com' + movies_with_links[movie], headers=headers)
            page_html = BeautifulSoup(res.text, "lxml")
            try:
                details_section = page_html.find(attrs={'data-testid': 'Details'})
                # getting countries
                country_of_origin_li = details_section.find(attrs={'data-testid': 'title-details-origin'})
                countries_of_origin = country_of_origin_li.find_all('a',
                                                                    class_='ipc-metadata-list-item__list-content-item ipc-metadata-list-item__list-content-item--link')
                countries_of_origin = [country.contents[0] for country in countries_of_origin]

                # getting Languages
                languages_li = details_section.find(attrs={'data-testid': 'title-details-languages'})
                languages = languages_li.find_all('a',
                                                  class_='ipc-metadata-list-item__list-content-item ipc-metadata-list-item__list-content-item--link')
                languages = [country.contents[0] for country in languages]
                countries_of_origin = ' '.join([str(w) for w in countries_of_origin])
                languages = ' '.join([str(w) for w in languages])
                movies_info.loc[movie] = [countries_of_origin, languages]
            except Exception as ex:
                print(ex)
            it += 1
    # save collected data in excel
    writer = pd.ExcelWriter('movies_info.xlsx', engine='xlsxwriter')
    movies_info.to_excel(writer, startrow=1, sheet_name='Sheet1')
    worksheet = writer.sheets['Sheet1']

    for i, col in enumerate(movies_info.columns):
        column_len = movies_info[col].astype(str).str.len().max()
        column_len = max(column_len, len(col)) + 2
        worksheet.set_column(i, i, column_len)
    writer.save()


if __name__ == '__main__':
    parse()
