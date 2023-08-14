import os
import time
import openpyxl
from datetime import datetime
import requests

import key
import location

# Substitua 'YOUR_API_KEY' pelo seu próprio chave de API do Google Maps
api_key = key.key_pass()   # Substitua pela sua API_KEY
base_url = 'https://maps.googleapis.com/maps/api/place' 

# Define a localização central (latitude e longitude) da sua redondeza
latitude = location.location_latitude()  #Substitua pela sua Latitude
longitude = location.location_longitude() #Substitua pela sua longitude

# Define o raio inicial em metros
radius_increment = 1000
max_radius = 5000

# Lista para armazenas todos os lugares
all_places = []

current_radius = radius_increment

while current_radius <= max_radius:
    # Faz a solicitação à API do Google Maps para empresas próximas
    places_url = f'{base_url}/nearbysearch/json'
    params = {
        'location': f'{latitude}, {longitude}',
        'radius': current_radius,
        #'type': 'establishment',
        'key': api_key 
    }
    response = requests.get(places_url, params=params)
    places_result = response.json()

    if 'results' in places_result:
        all_places.extend(places_result['results'])

    # Verifica se há mais páginas de resultados
    while 'next_page_token' in places_result:
        next_page_token = places_result['next_page_token']

        # Aguarda um pouco antes de solicitar a próxima página
        time.sleep(5)

        params['pagetoken'] = next_page_token
        response = requests.get(places_url, params=params)
        places_result = response.json()

        if 'results' in places_result:
            all_places.extend(places_result['results'])
        else:
            break

        # Incrementa a distância em 1000 metros
        current_radius += radius_increment

# Obtém a data atual
current_date = datetime.now().strftime('%Y-%m-%d')

# Encontra o nome de arquivo disponível para o dia atual
file_name = f'empresas_{current_date}.xlsx'
counter = 1
while os.path.exists(file_name):
    file_name = f'empresas_{current_date}_{counter}.xlsx'
    counter += 1

# Cria um arquivo Excel com as informações das empresas
wb = openpyxl.Workbook()
ws = wb.active
ws.append(['Nome da Empresa', 'Endereço', 'Número de Telefone', 'Categoria'])

for place in all_places:
    place_id = place['place_id']
    place_details_url = f'{base_url}/details/json'
    details_params = {
        'place_id': place_id,
        'fields': 'name,formatted_address,formatted_phone_number',
        'key': api_key
    }
    details_response = requests.get(place_details_url, params=details_params)
    place_details = details_response.json().get('result', {})

    name = place_details.get('name', 'N/A')
    address = place_details.get('formatted_address', 'N/A')
    phone = place_details.get('formatted_phone_number', 'N/A' )

    # Determina a categoria com base nos tipos de lugar
    types = place.get('types', [])
    category = ', '.join(types)

    ws.append([name, address, phone, category])

# Formata as colunas para ajustar automaticamente as larguras
for column_cells in ws.columns:
    max_length = 0
    column = column_cells[0].column  # Get the column name
    for cell in column_cells:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        except:
            pass
    adjusted_width = (max_length + 2) * 1.2
    column_letter = openpyxl.utils.get_column_letter(column) # Converte o número da coluna em letra
    ws.column_dimensions[column_letter].width = adjusted_width

# Salva o arquivo Excel
wb.save(file_name)

print(f'Informações das empresas salvas no arquivo "{file_name}"')
