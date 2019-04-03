import requests
from src.general.constants import LOCAL_BASE_URL

def get_employees_by_area(area_id):
    url = LOCAL_BASE_URL + 'personal/excel/personal'
    querystring = {'idArea' : area_id}
    payload = ''
    headers = {
        'cache-control': 'no-cache'
    }

    response = requests.request('GET', url, data=payload, headers=headers, params=querystring)
    json_data = response.json()
    return json_data
