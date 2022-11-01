from archicad import ACConnection
conn = ACConnection.connect()
assert conn

acc = conn.commands
act = conn.types
acu = conn.utilities

columns=(acc.GetElementsByType('Column'))
a = acu.GetBuiltInPropertyId('General_Height')
b = acu.GetBuiltInPropertyId('General_ElementID')
values = acc.GetPropertyValuesOfElements(columns, [a])
IDs = acc.GetPropertyValuesOfElements(columns, [b])
propertyValue = act.NormalLengthPropertyValue(1.0, 'length', 'normal')
i=0
for Propertyvalue in values:
    value = Propertyvalue.propertyValues[0].propertyValue.value
    id = IDs[i].propertyValues[0].propertyValue.value
    print(f"Wysokość słupa {id} w metrach: {value}")
    if value < 1:
        elementId = columns[i].elementId
        x = act.ElementPropertyValue(elementId, a, propertyValue)
        ElementPropertyValues = [x]
        y = acc.SetPropertyValuesOfElements(ElementPropertyValues)
        if y[0].success == True:
            print('Operacja zakończyła się powodzeniem.\n',f"Zmieniono wysokość słupa {id} na 1m. ")
        else:
            print(f'Operacja zakończona niepowodzeniem.\n{y[0].error}')
    i+=1
