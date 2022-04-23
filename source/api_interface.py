import requests
import json
import pandas as pd
from pandas import DataFrame
from datetime import datetime, timedelta
import pgeocode
import folium
import time
import statistics
from tqdm import tqdm

class CHouse:
    api_calls = 0
    total_api_calls = 0
    current_time = datetime.now()
    
    def __init__(self, api_key):
        self.api_key = api_key        
        
    @classmethod
    def api_guard(cls):
        call_time = datetime.now()
        buffer_delta = timedelta(seconds=420)
        
        if call_time > (CHouse.current_time + buffer_delta):
            CHouse.api_calls = 0
            CHouse.current_time = datetime.now()
            
        if CHouse.api_calls >= 599:
            print("API cool down 5 min")
            time.sleep(310)
            print(f"Restarting, current total API calls: {CHouse.total_api_calls}")
            CHouse.api_calls = 0
        else:
            CHouse.api_calls += 1
            CHouse.total_api_calls += 1
            
        CHouse.current_time = datetime.now()
            
    def _enrich_geo_loc(self, df: DataFrame) -> DataFrame:
        nomi = pgeocode.Nominatim('gb')
        output_lat = []
        output_lon = []
        postal_codes = []
        for i in range(df.shape[0]):
            try:
                temp = (df['registered_office_address'][i]['postal_code'])
                output_lat.append(nomi.query_postal_code(temp)[['latitude', 'longitude']][0])
                output_lon.append(nomi.query_postal_code(temp)[['latitude', 'longitude']][1])
                postal_codes.append(temp)
            except:
                output_lat.append('NaN')
                output_lon.append('NaN')
                postal_codes.append('NaN')
                
        df['Latitude'] = output_lat
        df['Longitude'] = output_lon
        #df['registered_office_address'] = postal_codes
        
        return df

    def _founder_api(self,company_number: str) -> (list, list, list):
        url_founder = "https://api.company-information.service.gov.uk/company/{}/persons-with-significant-control"
        birth_year = []
        name = []
        residency = []
        
        response = requests.get(url_founder.format(company_number),auth=(self.api_key,''))
        json_search_result = response.text
        search_result = json.JSONDecoder().decode(json_search_result)

        for i in range(len(search_result['items'])):
            birth_year.append(search_result['items'][i]['date_of_birth']['year'])
            name.append(search_result['items'][i]['name'])
            residency.append(search_result['items'][i]['country_of_residence'])

        return birth_year, name, residency
        
    def filter_sic(self, sic_code: int, status: str, start_index: int) -> DataFrame:
        url_companies = "https://api.company-information.service.gov.uk/advanced-search/companies?sic_codes={}&start_index={}&company_status={}"
        output = []
        while True:
            CHouse.api_guard()
            response = requests.get(url_companies.format(sic_code, start_index, status),auth=(self.api_key,''))
            json_search_result = response.text
            search_result = json.JSONDecoder().decode(json_search_result)
            if len(search_result) == 4:
                break
            output += search_result['items']
            index += 20
            df = pd.DataFrame(output)[['company_name', 'company_number', 'company_type', 'date_of_creation', 'registered_office_address', 'sic_codes']]
            df['date_of_creation'] =  pd.to_datetime(df['date_of_creation'], format='%Y-%m-%d')
                                
        return df
        
    def create_map(self, df: DataFrame, show: bool = False) -> object:
        df_temp = self._enrich_geo_loc(df)
        m = folium.Map(location=[51.5072, 0])

        for i in tqdm(range(df_temp.shape[0])):
            try:
                folium.Marker(
                    [df_temp['Latitude'][i], df_temp['Longitude'][i]], popup=f"<i>{df_temp['company_name'][i]}"
        ).add_to(m)
            except:
                continue
        
        if show == True:
            return m
        elif show == False:
            m.save("CompanyMap.html")
        

    def enrich_founder(self, df: DataFrame) -> None:
        today_year = CHouse.current_time.year
        result_year = []
        result_name = []
        result_residency = []
        
        for i in tqdm(range(len(df['company_number']))):
            try:
                CHouse.api_guard()
                export = self._founder_api(df['company_number'][i])
                result_year.append(export[0])
                result_name.append(export[1])
                result_residency.append(export[2])
            except:
                result_year.append('NaN')
                result_name.append('NaN')
                result_residency.append('NaN')
                
        result_year_median = []
        result_year_min = []
        result_year_max = []
        
        for i in tqdm(range(len(result_year))):
            try:
                result_year_median.append(today_year - statistics.median(result_year[i]))
                result_year_min.append(today_year - min(result_year[i]))
                result_year_max.append(today_year - max(result_year[i]))
            except:
                result_year_median.append('NaN')
                result_year_min.append('NaN')
                result_year_max.append('NaN')
            
        df['Significant Person Birth Year/s'] = result_year
        df['Median Person Age'] = result_year_median
        df['Oldest Person Age'] = result_year_min
        df['Youngest Person Age'] = result_year_max
        df['Significant Person Name/s'] = result_name
        df['Significant Person Residency'] = result_residency
        
    def retrieve_filings(self, company_number: str) -> DataFrame:
        url_filings = "https://api.company-information.service.gov.uk/company/{}/filing-history?items_per_page=200"
        
        CHouse.api_guard()
        response = requests.get(url_filings.format(company_number),auth=(self.api_key,''))
        json_search_result = response.text
        search_result = json.JSONDecoder().decode(json_search_result)
        
        
        date = []
        action = []
        des = []
        
        for i in range(len(search_result['items'])):
            date.append(search_result['items'][i]['date'])
            action.append(search_result['items'][i]['category'])
            des.append(search_result['items'][i]['description'])

        data = {'Date': date, 'Action': action, 'Details': des}
        search_result = pd.DataFrame(data)
        
        return search_result
    
    def export_excel(self, df:DataFrame) -> None:
        df.to_excel("CompaniesExport.xlsx")
