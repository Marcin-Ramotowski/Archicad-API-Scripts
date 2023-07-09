from data_tools.connection_init import acc, acu
from data_tools.functions import is_all_properties_available
import matplotlib.pyplot as plt
import numpy as np
import os


''' STREFA KONFIGURACJI '''
precision = 1 # do tylu miejsc po przecinku będą zaokrąglane wartości na wykresie
checked_stories = ['Parter'] # nazwy sprawdzanych kondygnacji, jeśli pusta sprawdzane są wszystkie kondygnacje
plot_folder = 'Inne/Wykresy' # ścieżka do folderu, do którego mają trafiać wygenerowane wykresy.
prefix_mode = True # jeśli True, wyświetla na wykresie nazwę razem z numerem strefy, jeśli False bez numeru.

properties_names = ['Zone_ZoneName', 'Zone_NetArea', ['Pozycja', 'Nazwa kondygnacji'], 'Zone_ZoneNumber']
property_path = 'Inne/Właściwości/pozycja.xml'
run = is_all_properties_available(properties_names, property_path)

if run:
    zones = acc.GetElementsByType('Zone')
    properties_ids = [acu.GetBuiltInPropertyId(name) if type(name) is str else acu.GetUserDefinedPropertyId(*name)
                      for name in properties_names]
    properties_values = acc.GetPropertyValuesOfElements(zones, properties_ids)
    elements_values = [cell.value for row in properties_values for cell in row]
    story_zones = {}

    for element in elements_values:
        zone_name = element[0]
        area = element[1]
        story = element[2]
        zone = f"{element[3]} {zone_name}" if prefix_mode else zone_name
        if len(checked_stories) == 0 or story in checked_stories:
            zones_in_story, areas = story_zones.get(story, [[], []])
            zones_in_story.append(zone)
            areas.append(area)
            story_zones[story] = [zones_in_story, areas]

    for story in story_zones:
        zones, areas = story_zones[story]
        total_area = sum(areas)
        percentages = list(map(lambda x: x/total_area*100, areas))
        smallest_percent = min(percentages)
        plt.figure()
        if smallest_percent < 5:
            # ustawienie podziałki i etykiet
            fig, ax = plt.subplots()
            ticks = np.arange(0, 101, 1)
            labels = [f'{x}%' if x % 5 == 0 else '' for x in ticks]
            ax.set_xticks(ticks)
            ax.set_xticklabels(labels)
            plt.barh(zones, percentages)

            # dołączenie wartości do słupków
            for index, value in enumerate(percentages):
                text = f"{round(value, precision)}%"
                plt.text(value+0.1, index, text)
            ax.set_xlim(min(percentages), max(percentages)+precision+2.5)
        else:
            plt.pie(percentages, labels=zones, autopct=f'%1.{precision}f%%')
        plt.title(f'Procentowe udziały powierzchni pomieszczeń \n {story}')
        exit_path = os.path.join(plot_folder, f'{story}.png')
        plt.savefig(exit_path, bbox_inches='tight')
        acu.OpenFile(exit_path)
