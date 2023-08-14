import os
import time
import googlemaps
import openpyxl
from datetime import datetime
import key
import location

# Substitua 'YOUR_API_KEY' pelo seu próprio chave de API do Google Maps
api_key = key.key_pass()
gmaps = googlemaps.Client(key=api_key)

# Define a localização central (latitude e longitude) da sua redondeza
latitude =  location.location_latitude()
longitude = location.location_longitude()

# Define o número máximo de resultados por página
results_per_page = 60

# Faz a solicitação à API do Google Maps para empresas próximas
places_result = gmaps.places_nearby(
    location=(latitude, longitude),
    radius=50000,  # Raio em metros
    #type='business'  # Tipo de lugares que você está buscando
)

# Lista para armazenas todos os lugares
all_places = places_result['results']

# Verifica se há mais páginas de resultados e coleta todas as páginas
while 'next_page_token' in places_result:
    next_page_token = places_result['next_page_token']

    # Aguarda um pouco antes de solicitar a próxima página
    time.sleep(2)

    # Faz a solicitação para a próxima página
    places_result = gmaps.places_nearby(
        location=(latitude, longitude),
        radius=50000,
        #type='business'
        page_token=next_page_token,
    )

    all_places.extend(places_result['results'])

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

for place in places_result['results']:
    place_details = gmaps.place(place['place_id'], fields=['name', 'formatted_address', 'formatted_phone_number'])
    name = place_details['result']['name']
    address = place_details['result'].get('formatted_address', 'N/A')
    phone = place_details['result'].get('formatted_phone_number', 'N/A')

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
