from geopy.geocoders import Nominatim
from geopy.distance import geodesic
from math import floor
import pandas as pd
import time

def get_distance(out, to):
    user_a = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
    geolocator = Nominatim(user_agent=user_a)
    try:
        from_city = geolocator.geocode(out)
    except :
        print("Много запросов")
        time.sleep(2)
        from_city = geolocator.geocode(out)

    try:
        to_city = geolocator.geocode(out)
    except:
        print("Много запросов")
        time.sleep(2)
        to_city = geolocator.geocode(out)

    if to_city == None or from_city == None:
        print(to_city)
        print(from_city)
        return
    from_city_lat = (from_city.latitude, from_city.longitude)
    to_city_lat = (to_city.latitude, to_city.longitude)

    return floor(geodesic(from_city_lat, to_city_lat).kilometers)

def work_with_excel(filename):
    df = pd.read_excel(filename, header=None)
    df = df.drop(columns=[0])
    df.columns = df.iloc[1]
    df = df.drop(1)
    df = df.dropna(how='all')
    print("Начат обход")
    for index, row in df.iterrows():
        out = row['Страна отправления'] + ' ' + row['Город отправления']
        to = row['Страна назначения'] + ' ' + row['Город назначения']

        distance = get_distance(out,to)

        row['Расстояние, km'] = distance

    df.to_excel(filename, index=False)
    

work_with_excel('test.xlsx')
