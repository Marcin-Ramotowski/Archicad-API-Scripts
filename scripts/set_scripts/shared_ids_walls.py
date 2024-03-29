from data_tools.connection_init import acc, act, acu

walls = acc.GetElementsByType('Wall')
compositeNameId = acu.GetBuiltInPropertyId('Construction_CompositeName')
element_id = acu.GetBuiltInPropertyId('General_ElementID')
values = acc.GetPropertyValuesOfElements(walls, [compositeNameId])
ids = acc.GetPropertyValuesOfElements(walls, [element_id])
printedTexts = []
ElementsPropertyValues = []
compositeNames = []
checked = []
indexes_occurrences = []
i = 0
print("Wykaz struktur warstwowych ścian w projekcie:")

for Propertyvalue in values:
    status = Propertyvalue.propertyValues[0].propertyValue.status
    id = ids[i].propertyValues[0].propertyValue.value
    if status == 'normal':
        value = Propertyvalue.propertyValues[0].propertyValue.value
        text = f"ID elementu:  {id}  Struktura warstwowa: {value}"
        if text not in printedTexts:
            print(text)
        printedTexts.append(text)
    else:
        value = f"Nie określono {i}"
    compositeNames.append(value)
    i += 1
i = j = l = 0
k = 0

for compositeName in compositeNames:
    quantity = compositeNames.count(compositeName)
    if quantity > 1:
        element = ids[i].propertyValues[0].propertyValue.value
        if compositeName not in checked:
            for x in compositeNames:
                if x == compositeName:
                    indexes_occurrences.append(j)
                j += 1
            j = 0
            first = indexes_occurrences[0]
            #newID = ids[first].propertyValues[0].propertyValue.value
            newID = f"SZ{l}"
            l += 1
            while k < len(indexes_occurrences):
                change = indexes_occurrences[k]
                toChange = walls[change].elementId
                propertyValue = act.NormalStringPropertyValue(newID, 'string', 'normal')
                x = act.ElementPropertyValue(toChange, element_id, propertyValue)
                ElementsPropertyValues.append(x)
                k += 1
            k = 0
            indexes_occurrences = []
        checked.append(compositeName)
    i += 1

if len(ElementsPropertyValues) > 0:
    y = acc.SetPropertyValuesOfElements(ElementsPropertyValues)
    if y[0].success is True:
        print('Operacja zakończyła się powodzeniem. Ujednolicono ID dla ścian o identycznych strukturach warstwowych.')
    else:
        print(f'Operacja zakończona niepowodzeniem.\n{y[0].error}')