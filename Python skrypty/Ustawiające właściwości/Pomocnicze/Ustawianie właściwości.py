from archicad import ACConnection
conn = ACConnection.connect()
assert conn

acc = conn.commands
act = conn.types
acu = conn.utilities

columns = acc.GetElementsByType('Column')
propertyId = acu.GetBuiltInPropertyId('General_Height')
propertyValue = act.NormalLengthPropertyValue(5.0, 'length', 'normal')
ElementPropertyValues = []
for column in columns:
    elementId = column.elementId
    x = act.ElementPropertyValue(elementId, propertyId, propertyValue)
    ElementPropertyValues = [x]
    y = acc.SetPropertyValuesOfElements(ElementPropertyValues)
    if y[0].success == True:
        print('Operacja zakończyła się powodzeniem.')
    else:
        print(f'Operacja zakończona niepowodzeniem.\n{y[0].error}')

#Wartość właściwości tworzymy poprzez użycie metody act.Normal[Typ]PropertyValue(wartość, typ, status)