from archicad import ACConnection
conn = ACConnection.connect()
assert conn

acc = conn.commands
acu = conn.utilities

'''PANEL KONFIGURACJI'''
elements = acc.GetElementsByType("Column")
a = acu.GetBuiltInPropertyId("General_Height")
b = acu.GetBuiltInPropertyId("General_ElementID")
print("Zestawienie słupów w projekcie:")
sequence = True  #jeśli True - sortowanie malejąco, jeśli False - rosnąco

heights = acc.GetPropertyValuesOfElements(elements, [a])
ids = acc.GetPropertyValuesOfElements(elements, [b])
wykaz = {}
i = 0
for Propertyvalue in heights:
    status = Propertyvalue.propertyValues[0].propertyValue.status
    id = ids[i].propertyValues[0].propertyValue.value
    if status == 'normal':
        value = Propertyvalue.propertyValues[0].propertyValue.value
        wykaz[id] = value
    else:
        value = f"Nie określono"
        wykaz[id] = value
    i += 1

lista = sorted(wykaz.items(), key=lambda x: x[1], reverse=sequence)
for element in lista:
    text = f'ID: {element[0]} Wysokość: {element[1]}'
    print(text)
