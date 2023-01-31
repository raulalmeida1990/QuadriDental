import pandas as pd
import numpy as np

import requests
import re
from bs4 import BeautifulSoup

df = pd.read_excel('QuadriMedical.xlsx')
df.dropna(subset=['Ref. '], inplace=True)
df.set_index('Ref. ', inplace=True)

link_ref = r'https://www.montellano.pt/pt/?_adin=11551547647'

df_search = pd.DataFrame()
for reference in df.index:
    price = np.nan
    link1 = fr'https://www.montellano.pt/pt/pesquisa_36.html?c=1&term={reference}'
    content = requests.get(link1)

    # Use regular expression to extract the link
    match = re.search("location='(.*)'", content.text)

    # Check if a match is found
    if match:
        link = match.group(1)
        print("Link:", link)
    else:
        price = np.nan
        print("No Price found for article: ")
        print(df.loc[reference, "Designação do artigo"])
        print('-' * 20)
        continue

    response = requests.get(link)

    # Check if the request is successful
    if response.status_code == 200:
        # Do something with the response

        # Get the HTML content
        content = response.text

        try:
            # Create a BeautifulSoup object
            soup = BeautifulSoup(content, 'html.parser')
            # Find the element with class "current"
            current_element = soup.find("span", class_="current")

            # Get the text of the element
            price = current_element.text
            price = float(price.replace(",", ".").replace("€", ""))

            print("Current value:", price)
        except:
            price = np.nan
            print("Current value:", price)
    else:
        price = np.nan
        print("Request failed")
    nome_artigo = df.loc[reference, "Designação do artigo"]
    print(df.loc[reference, "Designação do artigo"])
    print(price)
    print('-' * 20)
    data = {
        'Designacao': [nome_artigo],
        link_ref: [price]
    }
    df_temp = pd.DataFrame(data)
    df_temp.index = [reference]
    df_temp.index.name = "Reference"
    df_search = pd.concat([df_search, df_temp])

df_search.to_excel('dental.xlsx')
