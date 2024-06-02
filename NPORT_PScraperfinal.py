import requests
import json
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta


CIK_List = ['0001678124', '0001736510', '0001842754', '0001736035', '0001061630']


# Function to get the previous month last dat from the current date.
def get_previous_month_last_date(specific_date):
    if isinstance(specific_date, str):
        specific_date = datetime.strptime(specific_date, "%Y-%m-%d")
    first_day_of_current_month = specific_date.replace(day=1)
    last_day_of_previous_month = (first_day_of_current_month - relativedelta(days=1)).strftime("%Y-%m-%d")
    return last_day_of_previous_month

# function for getting the filing table data using the CIK. and then filtering for the NPORT-P and date greater than '2019-01-01
def getFilings(cik):
    tablerows = []

    # headers for the api. this is just the header that is passed to the api. it has nothing to do with google chrome or any browser. it will work in any os.
    header = {
        'Accept':'*/*',
        'Accept-Encoding':'gzip, deflate, br, zstd',
        'Accept-Language':'en-US,en;q=0.9',
        'Cache-Control':'no-cache',
        'Origin':'https://www.sec.gov',
        'Pragma':'no-cache',
        'Priority':'u=1, i',
        'Referer':'https://www.sec.gov/',
        'Sec-Ch-Ua':'"Google Chrome";v="125", "Chromium";v="125", "Not.A/Brand";v="24"',
        'Sec-Ch-Ua-Mobile':'?0',
        'Sec-Ch-Ua-Platform':'"Windows"',
        'Sec-Fetch-Dest':'empty',
        'Sec-Fetch-Mode':'cors',
        'Sec-Fetch-Site':'same-site',
        'User-Agent':'ashok_jj@hotmail.com'
    }
    resp2 = requests.get(f'https://data.sec.gov/submissions/CIK{cik}.json', headers=header) # sending a get request to search the cik
    if resp2.status_code == 200: # Checking the response status code. if status code = 200 it means data is available for that CIK.
        data = resp2.json()['filings']['recent'] # Converting the response into json
        df = pd.DataFrame() # Creating a dataFrame

        # Adding different Column into the dataFrame
        df['primaryDocument'] = data['primaryDocument']
        df['accessionNumber'] = data['accessionNumber']
        df['form type'] = data['form']
        df['filingDate'] = data['filingDate']
        df['reportDate'] = data['reportDate']
        df['cik'] = resp2.json()['cik']
        df['short name'] = resp2.json()['name']
        df = df[df['form type']=='NPORT-P'] # Filtering the dataFrame for NPORT-P
        df = df[df['filingDate'] >= '2019-01-01'] # Filtering the dataFrame for date grater than 2019-01-01
        tablerows = json.loads(df.to_json(orient='records')) # Converting the filtered dataFrame into json object.
    return tablerows

# Function for getting the xml of NPORT-P file and passing into the BeautifulSoup object
def getFileData(searchCik, accessionNumber, primaryDocument):

    # headers for the api. this is just the header that is passed to the api. it has nothing to do with google chrome or any browser. it will work in any os.
    header = {
        'Accept':'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
        'Accept-Encoding':'gzip, deflate, br, zstd',
        'Accept-Language':'en-US,en;q=0.9',
        'Cache-Control':'no-cache',
        'Pragma':'no-cache',
        'Priority':'u=0, i',
        'Sec-Ch-Ua':'"Google Chrome";v="125", "Chromium";v="125", "Not.A/Brand";v="24"',
        'Sec-Ch-Ua-Mobile':'?0',
        'Sec-Ch-Ua-Platform':'"Windows"',
        'Sec-Fetch-Dest':'document',
        'Sec-Fetch-Mode':'navigate',
        'Sec-Fetch-Site':'same-origin',
        'Upgrade-Insecure-Requests':'1',
        'User-Agent':'ashok_jj@hotmail.com'
    }
    resp3 = requests.get(f'https://www.sec.gov/Archives/edgar/data/{searchCik}/{accessionNumber}/{primaryDocument}', headers=header)
    soup = BeautifulSoup(resp3.content, 'lxml')  # Creating a BeautifulSoup object
    tr = soup.find_all('tr')  # getting all the table rows.
    return tr


rows = []
id = 1
for cik in CIK_List: #Looping through the CIK list
    tablerows = getFilings(cik) # Searhing the CIK. if data found it will return the filtered data.
    for tablerow in tablerows:
        accessionNumber = tablerow['accessionNumber'].replace('-','')
        primaryDocument = tablerow['primaryDocument']
        searchCik = tablerow['cik']
        tr = getFileData(searchCik, accessionNumber, primaryDocument)  # getting all the tables rows from the xml.
        EndDate = ''
        RefYear = ''
        RefMonth = ''
        TotalAssets = ''
        TotalLiabilities = ''
        NetAssets = ''
        Subs = []
        ReInvest = []
        Reds = []

        # Looping through all the table rows and getting these data points from it.
        for i in tr:
            if i.find('td',{'class':'label'}):
                heading_text = i.find('td',{'class':'label'}).text
                if 'Date as of which information is reported.' in heading_text:
                    EndDate = i.find('div',{'class':'fakeBox2'}).text.strip()
                    RefYear = EndDate.split('-')[0]
                    RefMonth = EndDate.split('-')[1]
                elif 'Total assets, including assets attributable to miscellaneous securities reported in Part D' in heading_text:
                    TotalAssets = i.find_all('td')[-1].text.strip()
                elif 'Total liabilities.' in heading_text:
                    TotalLiabilities = i.find_all('td')[-1].text.strip()
                elif 'Net assets.' in heading_text:
                    NetAssets = i.find_all('td')[-1].text.strip()
        # Again Looping through all the table rows and getting the Subs Reinvest and Reds and appending into a list.
        for i in tr:
            if i.find('td',{'class':'label'}):
                heading_text = i.find('td',{'class':'label'}).text
                if 'Total net asset value of shares sold (including exchanges but excluding reinvestment of dividends and distributions).' in heading_text:
                    Subs.append(i.find_all('td')[-1].text.strip())
                elif 'Total net asset value of shares sold in connection with reinvestments of dividends and distributions.' in heading_text:
                    ReInvest.append(i.find_all('td')[-1].text.strip())
                elif 'Total net asset value of shares redeemed or repurchased, including exchanges.' in heading_text:
                    Reds.append(i.find_all('td')[-1].text.strip())
        
        # Loop for formatting the Captured data into the desired excel format
        for i in range(1,4):
            row = {}
            row['id'] = id
            row['CIK'] = int(cik)
            row['shortName'] = tablerow['short name']
            if i==1:
                row['EndDate'] = EndDate
                row['RefYear'] = int(RefYear)
                row['RefMonth'] = int(RefMonth)
                row['QuarterlySubs'] = float(Subs[0])+float(Subs[1])+float(Subs[2]) # Calculating the QuarterlySubs
                row['QuarterlyReds'] = float(Reds[0])+float(Reds[1])+float(Reds[2]) # Calculating the QuarterlyReds
                row['QuarterlyFlow'] = row['QuarterlySubs'] - row['QuarterlyReds'] # Calculating the QuarterlyFlow
            else:
                EndDate = get_previous_month_last_date(EndDate)
                RefYear = EndDate.split('-')[0]
                RefMonth = EndDate.split('-')[1]
                row['EndDate'] = EndDate
                row['RefYear'] = int(RefYear)
                row['RefMonth'] = int(RefMonth)
                row['QuarterlySubs'] = ''
                row['QuarterlyReds'] = ''
                row['QuarterlyFlow'] = ''
            
            row['TotalAssets'] = float(TotalAssets)
            row['TotalLiabilities'] = float(TotalLiabilities)
            row['NetAssets'] = float(NetAssets)
            row['Subs'] = float(Subs[len(Subs)-i])
            row['Reds'] = float(Reds[len(Reds)-i])
            row['ReInvest'] = float(ReInvest[len(ReInvest)-i])
            rows.append(row)
            print(row)
        id = id+1


df = pd.DataFrame(rows)
df = df[['id', 'CIK', 'shortName', 'EndDate', 'RefYear', 'RefMonth', 'TotalAssets', 'TotalLiabilities', 'NetAssets', 'Subs', 'Reds', 'ReInvest', 'QuarterlySubs', 'QuarterlyReds', 'QuarterlyFlow']]
df.to_excel('output.xlsx', index=False)

display(df)