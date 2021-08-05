from random import randrange

import pandas as pd
import requests
from bs4 import BeautifulSoup
import time


def get_titles_and_urls():
    df = pd.read_excel('input.xlsx')
    df = df[['title', 'url']]
    df = df.sort_values('title', ascending=True)
    return df


def generate_reference(author, title, year, url):
    return "{}. {}. Forbes. {}. Disponível em: {}. Acesso em: 4 ago. 2021".format(author, title, year, url)


def get_author_and_year(url):
    html = requests.get(url).content
    soup = BeautifulSoup(html, 'html.parser')
    year = '0000'
    author = 'VOID'
    for div in soup.findAll('div', attrs={'class': 'content-data'}):
        if not div.find('time'):
            return year, author

        year = div.find('time').contents[0].split(' ')[2]
        year = year[:-1]  # drop the ',' at the end
        break

    # for div in soup.findAll('div', attrs={'class': 'contribs'}):
    for div in soup.findAll('a', attrs={'class': 'contrib-link--name remove-underline'}):
        if not div.contents:
            continue

        author = div.contents[0]
        break

    return year, author


def fetch_articles():
    out_file = 'output.csv'
    df = get_titles_and_urls()
    total_articles = df.shape[0]
    curr_article = 1
    try:
        for index, row in df.iterrows():
            print('Processing {}/{}'.format(curr_article, total_articles))
            year, author = get_author_and_year(row['url'])
            df.at[index, 'year'] = year
            df.at[index, 'author'] = author
            time_to_sleep = randrange(2, 10)
            time.sleep(time_to_sleep)
            curr_article += 1
    finally:
        df.to_csv(out_file, index=False)


def get_abnt_name(name):
    names = name.split(' ')
    first_name = names[0]
    last_name = names[-1]
    return last_name.upper() + ', ' + first_name


if __name__ == '__main__':
    df = pd.read_csv('output.csv')
    df = df[['year', 'author', 'title', 'url']]
    for index, row in df.iterrows():
        abnt_name = get_abnt_name(row['author'])
        df.at[index, 'author'] = abnt_name
        df.at[index, 'abnt'] = generate_reference(abnt_name, row['title'], row['year'], row['url'])

    df = df.sort_values(['year', 'author'], ascending=True)
    writer = pd.ExcelWriter("output.xlsx", engine='xlsxwriter')
    df.to_excel(writer, sheet_name='teste', index=False)
    writer.save()
