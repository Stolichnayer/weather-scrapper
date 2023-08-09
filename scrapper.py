__author__      = "Alex Perrakis"
__date__        = "30-03-2023"

import requests
from bs4 import BeautifulSoup
import win32com.client as win32

# Returns the wind direction based on its degrees
def get_wind_direction(degrees):

    if (degrees >= 0 and degrees <= 22.5):
        direction = "Β"
    elif (degrees > 22.5 and degrees <= 67.5):
        direction = "ΒΑ"
    elif (degrees > 67.5 and degrees <= 112.5):
        direction = "Α"
    elif (degrees > 112.5 and degrees <= 157.5):
        direction = "ΝΑ"
    elif (degrees > 157.5 and degrees <= 202.5):
        direction = "Ν"
    elif (degrees > 202.5 and degrees <= 247.5):
        direction = "ΝΔ"
    elif (degrees > 247.5 and degrees <= 292.5):
        direction = "Δ"
    elif (degrees > 292.5 and degrees <= 337.5):
        direction = "ΒΔ"
    elif (degrees > 337.5 and degrees <= 360):
        direction = "Β"
    else:
        direction = "Unknown"
    
    return direction

def get_sea_state(wind_degree):

    sea_state = {
        0: "ΓΑΛΗΝΗ",
        1: "ΓΑΛΗΝΗ",
        2: "ΗΡΕΜΗ",
        3: "ΛΙΓΟ ΤΑΡΑΓΜΕΝΗ",
        4: "ΛΙΓΟ ΤΑΡΑΓΜΕΝΗ",
        5: "ΤΑΡΑΓΜΕΝΗ",
        6: "ΚΥΜΑΤΩΔΗΣ",
        7: "ΠΟΛΥ ΚΥΜΑΤΩΔΗΣ",
        8: "ΤΡΙΚΥΜΙΩΔΗΣ",
        9: "ΤΡΙΚΥΜΙΩΔΗΣ",
        10: "ΠΟΛΥ ΤΡΙΚΥΜΙΩΔΗΣ",
        11: "ΑΓΡΙΑ",
        12: "ΠΟΛΥ ΑΓΡΙΑ"
    }

    return sea_state[wind_degree]

def get_weather_state(icon_number): 

    # Dictionary of weather states
    weather_states = {
        "1": "ΑΙΘΡΙΟΣ ΚΑΙΡΟΣ",
        "2": "ΑΡΑΙΗ ΣΥΝΝΕΦΙΑ",
        "3": "ΑΥΞΗΜΕΝΗ ΣΥΝΝΕΦΙΑ",
        "4": "ΣΥΝΝΕΦΙΑ",
        "5": "ΠΙΘΑΝΗ ΒΡΟΧΗ",
        "6": "ΑΣΘΕΝΗΣ ΒΡΟΧΗ",
        "7": "ΒΡΟΧΗ",
        "10": "ΚΑΤΑΙΓΙΔΑ",
        "26": "ΧΙΟΝΟΠΤΩΣΗ",
        "28": "ΧΙΟΝΟΚΑΤΑΙΓΙΔΑ",
        "30": "ΑΙΘΡΙΟΣ ΚΑΙΡΟΣ",
        "31": "ΑΣΘΕΝΗΣ ΒΡΟΧΗ"
    }
    
    return weather_states[icon_number]

def create_island_weather_info_row(url):   

    # Send an HTTP GET request to the URL
    response = requests.get(url)

    # Parse the HTML content of the page using BeautifulSoup
    soup = BeautifulSoup(response.content, 'html.parser')

    #--------------------------- Title -----------------------------

    # Find all elements with class "wind" and extract their text
    titles = [get_weather_state(title['data-icon']) for title in soup.find_all('span', class_='wicon w78x73')]

    titles = list(set(titles))

    # Format the titles string to the final printable string

    titles_formatted = " - ".join(titles)

    #----------------------- Wind direction------------------------
    wind_degrees = [get_wind_direction(int(wind.text[:-1])) for wind in soup.find_all('div', class_='wind-popinfo')]
    wind_degrees = list(set(wind_degrees))

    wind_degrees_formatted = " - ".join(wind_degrees)

    #------------------------- Wind speed--------------------------

    # Find all elements with class "wind" and extract their text
    wind_speed = [int(wind.text[:-2]) for wind in soup.find_all('span', class_='wind')]

    # Calculate the minimum and maximum wind degrees
    min_wind = min(wind_speed)
    max_wind = max(wind_speed)

    # Format the wind degrees string based on the minimum and maximum values
    if min_wind == max_wind:
        wind_speed_str = f"{min_wind} Bf"
    else:
        wind_speed_str = f"{min_wind} - {max_wind} Bf"

    #------------------------- Sea state --------------------------
    sea_states_min = get_sea_state(min_wind)
    sea_states_max = get_sea_state(max_wind)

    if (sea_states_min == sea_states_max):
        sea_state_formatted = sea_states_min
    else:
        sea_state_formatted = f"{sea_states_min} - {sea_states_max}"

    #------------------------ Temperature -------------------------

    # Find all elements with class "wind" and extract their text
    temp_degrees = [int(temp.text[:-2]) for temp in soup.find_all('span', class_='temp')]

    # Calculate the minimum and maximum wind degrees
    min_temp = min(temp_degrees)
    max_temp = max(temp_degrees)

    # Format the wind degrees string based on the minimum and maximum values
    if min_temp == max_temp:
        temp_degrees_str = f"{min_temp} °C"
    else:
        temp_degrees_str = f"{min_temp} - {max_temp} °C"

    return [titles_formatted, wind_degrees_formatted, wind_speed_str, sea_state_formatted, temp_degrees_str]

def create_msword_table_row(row_num, table, text_list):

    for i, cell in enumerate(text_list):
        table.Cell(row_num, i + 1).Range.Text = text_list[i]; 

def create_msword_table(data):

    # Create a new MS Word document
    word = win32.gencache.EnsureDispatch('Word.Application')
    doc = word.Documents.Add()

    # Create the MS Word table with 4 rows and 6 columns
    table = doc.Tables.Add(doc.Range(0, 0), 4, 6)

    create_msword_table_row(1, table, data[0])
    create_msword_table_row(2, table, data[1])
    create_msword_table_row(3, table, data[2])
    create_msword_table_row(4, table, data[3])

    # save the document and close Word
    doc.SaveAs('weather.docx')
    doc.Close()
    word.Quit()
    
def create_text_file(data):

    with open('weather.txt', 'w') as f:
        f.write('Ν.ΡΟΔΟΣ\t' + '\t'.join(data[0]) + '\n')
        f.write('Ν.ΜΕΓΙΣΤΗΣ\t' + '\t'.join(data[1]) + '\n')
        f.write('Ν.ΚΑΡΠΑΘΟΣ\t' + '\t'.join(data[2]) + '\n')
        f.write('Ν.ΣΥΜΗ\t' + '\t'.join(data[3]) + '\n')

# URL of the page to be scraped
rhodes_url = 'https://freemeteo.gr/kairos/rodos/imerisia-provlepsi/aurio/?gid=400666'
megisti_url = 'https://freemeteo.gr/kairos/nisos-megisti/imerisia-provlepsi/aurio/?gid=257079'
karpathos_url = 'https://freemeteo.gr/kairos/nisos-karpathos/imerisia-provlepsi/aurio/?gid=260893'
sumi_url = 'https://freemeteo.gr/kairos/sumi/imerisia-provlepsi/aurio/?gid=253858'

data = []
data.append(create_island_weather_info_row(rhodes_url))
data.append(create_island_weather_info_row(megisti_url))
data.append(create_island_weather_info_row(karpathos_url))
data.append(create_island_weather_info_row(sumi_url))

#print(data)

#create_msword_table(data)

create_text_file(data)

