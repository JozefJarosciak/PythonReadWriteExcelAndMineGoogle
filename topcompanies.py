import urllib
import pandas as pd
import requests as requests
from bs4 import BeautifulSoup

datasetLocation = "top-companies-in-the-world-by-market-value-2018.xlsx";
df = pd.read_excel(datasetLocation, "Sheet1")
userAgent = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.103 Safari/537.36'}
searchURL = "https://www.google.com/search?q="
for x in range(0, len(df.index)):
    if str(df.iloc[x, 2]) is None:
        print(str(df.iloc[x, 0]) + " | " + str(df.iloc[x, 2]) + " | " + str(df.iloc[x, 3]))
    else:
        companyName = df.iloc[x, 0]
        search1 = BeautifulSoup(requests.get(searchURL + urllib.parse.quote_plus(companyName) + '+headquarters ', headers=userAgent).text, 'html.parser')
        headquarters = search1.find('div', {"class": "Z0LcW"}).get_text()
        if not headquarters:
            headquarters = search1.find('div', {"class": "desktop-title-subcontent"}).get_text()
        search2 = BeautifulSoup(requests.get(searchURL + urllib.parse.quote_plus(headquarters) + '+coordinates', headers=userAgent).text, 'html.parser')
        coordinates = search2.find('div', {"class": "Z0LcW"}).get_text()
        df.loc[x, 'Location'] = headquarters;
        df.loc[x, 'Coordinates'] = coordinates
        print(df.iloc[x, 0] + " | " + headquarters + " | " + coordinates)
    df.to_excel(datasetLocation, index=False)
