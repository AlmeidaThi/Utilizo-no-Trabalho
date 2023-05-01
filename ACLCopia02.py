import numpy as np
import pandas as pd
import pytz
from geopy.point import Point
from datetime import date, datetime
from timezonefinder import TimezoneFinder


tf = TimezoneFinder()
today = date.today()

time_inicio = datetime.now()


# //////////////////////////////////////////////////////

import os

path = os.getcwd()
dir_list = os.listdir(path)

print("Arquivos e diretórios em '", path, "' :")

for arquivo in os.listdir():
    if arquivo.endswith(".xlsx"):
        print(arquivo)

# //////////////////////////////////////////////////////

# import nomation
from geopy.geocoders import Nominatim

geolocator = Nominatim(user_agent="Python")

# Nome do Arquivo
DB_df = pd.read_excel(arquivo)


def ReverseGeoCode():
    #DB_df[['País', 'Estado', 'Cidade', 'CEP', 'GMT', 'Exclusão']] = ''

    for i in range(0, len(DB_df)):
        progress = np.round(((i + 1) * 100) / len(DB_df), 2)

        time_fim = (datetime.now() - time_inicio).total_seconds()


        text_progress = "Endereço capturado: " + str(progress) + "%"
        print(text_progress + f" - {time_fim:.03f} segundos")

        #Colunas de latitude e longitude no excel
        lat = DB_df.iloc[i, 13]
        lng = DB_df.iloc[i, 14]

        coord = Point(lat, lng)

        timezone_str = tf.certain_timezone_at(lat=lat, lng=lng)
        IST = pytz.timezone(timezone_str)
        datetime_ist = datetime.now(IST)

        valida_gmt = datetime_ist.strftime('%z')
        #Coluna do GMT Localização no excel
        DB_df.iloc[i, 15] = str(valida_gmt[0:3]) + ":" + str(valida_gmt[3:5])


        location = geolocator.reverse(coord, exactly_one=True,timeout=10)
        address = location.raw['address']

        #Coluna Google Maps

        DB_df.iloc[i, 16] = '=HYPERLINK("{}", "{}")'.format('https://www.google.com.br/maps/place/' + str(lat) + '+' + str(lng) + '/@' + str(lat) + ',' + str(lng), "Ver no mapa")
        
        #Colunas de País, Estado e Cidade no Excel
        DB_df.iloc[i, 17] = address.get('country', '')
        DB_df.iloc[i, 18] = address.get('state', '')
        DB_df.iloc[i, 19] = address.get('city', '')

        #DB_df.iloc[i, 21] = address.get('postcode', '')

        coluna_clockin_gmt = DB_df.iloc[i,9]
        coluna_gmt_lat_long = DB_df.iloc[i, 15]

        #Coluna de Exclusão no excel
        DB_df.iloc[i,20] = coluna_clockin_gmt == coluna_gmt_lat_long

    #Filtra pela coluna Exclusão os diferentes de verdadeiro
    DB_df_02 = DB_df.loc[DB_df['Exclusão']!=True]

    #Excluí a coluna "Exclusão"
    DB_df_03 = DB_df_02.drop(labels='Exclusão', axis=1)

    #Salva excel completo
    file_name = str("Convertido_Total_" + str(today.strftime("%d_%m_%Y")) + ".xlsx")
    DB_df.to_excel(file_name, index=False)

    #Salva excel com filtro na coluna Exclusão
    file_name2 = str("Convertido_Filtrado_" + str(today.strftime("%d_%m_%Y")) + ".xlsx")
    DB_df_02.to_excel(file_name2, index=False)

    #Salva excel filtrado excluindo a coluna Exclusão
    file_name3 = str("Convertido_Excluído_" + str(today.strftime("%d_%m_%Y")) + ".xlsx")
    DB_df_03.to_excel(file_name3, index=False)

    print("Salvo com sucesso!!!")


ReverseGeoCode()
