from geopy.geocoders import Nominatim
from geopy.distance import geodesic
from math import floor
import pandas as pd
import time
from fake_useragent import UserAgent

def get_distance(out, to):
    out_cord = {'Hungary Budapest': (47.4978789, 19.0402383), 'Germany Duisburg': (51.434999, 6.759562), 'Germany Hamburg': (53.550341, 10.000654), 'Poland Malaszewicze': (52.0240425, 23.5310557), 'Austria Vienna': (48.2083537, 16.3725042), 'Spain Madrid': (40.4167047, -3.7035825), 'France Lyon': (45.7578137, 4.8320114), 'Italy Milan': (45.4641943, 9.1896346), 'Czech Republic Ceska-Treshebova': (49.90436, 16.44413)}
    user_agent = UserAgent().random
    geolocator = Nominatim(user_agent=user_agent, timeout=200)
    
    try:
        to_city = geolocator.geocode(to)
    except: 
        print("Many request")
        time.sleep(5)
        to_city = geolocator.geocode(to)

    if to_city == None:
        print(to)
        return
    
    from_city_lat = out_cord[out]
    to_city_lat = (to_city.latitude, to_city.longitude)

    return floor(geodesic(from_city_lat, to_city_lat).kilometers)

def work_with_excel(filename):
    df = pd.read_excel(filename, header=None, engine='openpyxl')
    df = df.drop(columns=[0])
    df.columns = df.iloc[1]
    df = df.drop(1)
    df = df.dropna(how='all')
    print("Start")
    for index, row in df.iterrows():
        out = row['Страна отправления'] + ' ' + row['Город отправления']
        to = row['Страна назначения'] + ' ' + row['Город назначения']

        distance = get_distance(out,to)
        row['Расстояние, km'] = distance

    df.to_excel("all_res3.xlsx", index=False)
    

work_with_excel('3.xlsx')
