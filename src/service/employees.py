import requests
from src.general.constants import PROD_BASE_URL as base_url

def get_employees_by_area(area_id):
    url = base_url + 'personal/excel/personal'
    querystring = {'idArea' : area_id}
    payload = ''
    headers = {
        'cache-control': 'no-cache'
    }

    response = requests.request('GET', url, data=payload, headers=headers, params=querystring)
    json_data = response.json()
    return json_data
