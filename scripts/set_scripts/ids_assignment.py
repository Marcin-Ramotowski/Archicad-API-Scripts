from data_tools.connection_init import acc, act, acu

allProperties = acc.GetAllPropertyNames()
types = ['Ściana', 'Słup', 'Przegroda strukturalna', 'Strop', 'Belka', 'Dach', 'Drzwi',
          'Okno', 'Schody', 'Obiekt', 'Kształt', 'Lampa', 'Balustrada', 'Strefa']
idsForElements = {'Słup': 'S', 'Przegroda strukturalna': 'SK', 'Strop': 'P', 'Belka': 'B',
                    'Dach': 'DA', 'Drzwi': 'D', 'Okno': 'O', 'Schody': 'SCH', 'Obiekt': 'OB',
                    'Kształt': 'KS', 'Lampa': 'LP', 'Balustrada': 'BL', 'Strefa': 'POM',
                    'Ściana zewnętrzna': 'SZ', 'Ściana wewnętrzna': 'SW', 'Ściana niezidentyfikowana': 'SN'}
translatorPositions = {'Interior': 'Ściana wewnętrzna', 'Exterior': 'Ściana zewnętrzna', 
                       'Undefined': 'Ściana niezidentyfikowana'}
elementsCounter = {} # dict what count number of elements by type

elements = acc.GetAllElements()
propertyId = acu.GetBuiltInPropertyId('General_ElementID')
propertyId2 = acu.GetBuiltInPropertyId('Category_Position')
propertyId3 = acu.GetBuiltInPropertyId('General_Type')
positions = acc.GetPropertyValuesOfElements(elements, [propertyId2])
type_objects = acc.GetPropertyValuesOfElements(elements, [propertyId3])
ElementPropertyValues = []
a = 0

for type_object in type_objects:
    type = type_object.propertyValues[0].propertyValue.value
    elementId = elements[a].elementId
    if type in types:
        if type =='Ściana':
            ''' Ustala konkretny typ ściany na podstawie zdefiniowanej przez modelera pozycji
               pobranej z właściwości Category_Position. '''
            position = positions[a].propertyValues[0].propertyValue.value.nonLocalizedValue
            type = translatorPositions[position]
        amount = elementsCounter.get(type, 0) + 1
        elementsCounter[type] = amount
        value = f"{idsForElements[type]}{amount}"
        wrappedValue = act.NormalStringPropertyValue(value, 'string', 'normal')
        newValue = act.ElementPropertyValue(elementId, propertyId, wrappedValue)
        ElementPropertyValues.append(newValue)
    a += 1

logs = acc.SetPropertyValuesOfElements(ElementPropertyValues)
errors = [f"Error code {log.error.code}: {log.error.message}" for log in logs if not log.success]
if errors:
    print('The following problems have occurred:')
    for error in errors:
        print(error)
else:
    print('All operations successfully completed')
